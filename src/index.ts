import { Context, Schema, Logger, h, $ } from 'koishi'
import { readFileSync, writeFileSync, mkdirSync, unlinkSync } from 'fs'
import { join } from 'path'
import { tmpdir } from 'os'
import axios from 'axios'
import XLSX from 'xlsx'
import puppeteer from 'puppeteer'

// 声明 Leaflet 的全局变量 L，避免 TS 报错
declare const L: any;

export interface MapEntry {
  code: string
  description: string
}

export interface Config {
  commandname: string
  weatherMap: MapEntry[]
  cloudCoverageMap: MapEntry[]
  imageWidth: number
  imageHeight: number
  screenshotquality: number
  pageautoclose: boolean
}

export const name = 'route-search'

export const Config: Schema<Config> = Schema.object({
  commandname: Schema.string().default('航路查询').description('插件指令名'),
  pageautoclose: Schema.boolean().default(true).description('截图后自动关闭页面'),
  imageWidth: Schema.number().default(1280).description('截图宽度'),
  imageHeight: Schema.number().default(720).description('截图高度'),
  screenshotquality: Schema.number().default(80).description('截图质量 (0-100)'),
  weatherMap: Schema.array(Schema.object({
    code: Schema.string().required(),
    description: Schema.string().required(),
  })).role('table').default([/* ... your default weather map here ... */]),
  cloudCoverageMap: Schema.array(Schema.object({
    code: Schema.string().required(),
    description: Schema.string().required(),
  })).role('table').default([/* ... your default cloud coverage map here ... */]),
})

export function apply(ctx: Context, config: Config) {
  const logger = new Logger('route-search')
  const HOST = 'https://api.xflysim.com/pilot'
  const TEMP_DIR = join(ctx.baseDir, 'route-temp')
  const airports: Record<string, { icao: string; iata: string; name: string }> = {}

  const workbook = XLSX.readFile(join(__dirname, 'airports.xlsx'))
const sheet = workbook.Sheets[workbook.SheetNames[0]]
const rows = XLSX.utils.sheet_to_json(sheet) as any[]

for (const row of rows) {
  if (row.CODE_ID) {
    const icao = row.CODE_ID.toString().toUpperCase()
    const iata = row.CODE_IATA?.toString().toUpperCase() || ''
    const name = row.TXT_NAME?.toString() || '未知机场'

    const airport = {
      icao,
      iata,
      name,
    }
    airports[icao] = airport
    if (iata) airports[iata] = airport
    airports[name] = airport
  }
}


  ctx.command(config.commandname || '航路查询 <dep> <arr>', '航路查询')
  .action(async ({ session }, depInput, arrInput) => {
  // 尝试读取机场信息
  let depInfo = airports[depInput.toUpperCase()] || airports[depInput]
  let arrInfo = airports[arrInput.toUpperCase()] || airports[arrInput]

  // 如果没有找到，就直接用输入作为ICAO，拼成一个简单对象
  if (!depInfo) {
    depInfo = { icao: depInput.toUpperCase(), iata: '', name: depInput }
  }
  if (!arrInfo) {
    arrInfo = { icao: arrInput.toUpperCase(), iata: '', name: arrInput }
  }

  const dep = depInfo.icao
  const arr = arrInfo.icao

    // 获取最新cycle
    const cycleRes = await axios.get(`${HOST}/api/realTimeMap/getRouteCycle`).catch(() => null)
    if (!cycleRes?.data?.data) return '获取 CYCLE 数据失败'

    let cycles: string[] = []
    try {
      cycles = JSON.parse(cycleRes.data.data)
    } catch {
      return '解析 CYCLE 数据失败'
    }
    if (!Array.isArray(cycles) || cycles.length === 0) return '获取 CYCLE 数据失败'

    const maxCycle = Math.max(...cycles.map(c => Number(c)))
    const cycle = maxCycle.toString()

    // 请求航路
    const routeRes = await axios.post(`${HOST}/api/realTimeMap/route`, null, {
      params: { dep, arr, cycle },
      timeout: 10000,
    }).catch(() => null)

    if (!routeRes || routeRes.data.code !== 20000) return '获取航路失败'

    // 解析航路数据
    let routeDataRaw: any
    try {
      routeDataRaw = JSON.parse(routeRes.data.data)
    } catch {
      return '解析航路数据失败'
    }

    // 构造waypoints数组
    const waypoints = (routeDataRaw.nodeinformation || []).map((item: any[]) => ({
      name: item[0],
      lat: item[1],
      lon: item[2],
    }))

    // 组合数据
    const routeData = {
      ...routeDataRaw,
      waypoints,
    }

    // 获取天气信息
    const depWeather = await axios.get(`${HOST}/api/realTimeMap/weather/${dep}`).catch(() => null)
    const arrWeather = await axios.get(`${HOST}/api/realTimeMap/weather/${arr}`).catch(() => null)
    const depMetar = {
      raw: depWeather?.data?.data?.metar || '获取失败',
      decoded: depWeather?.data?.data?.metarDecode ? JSON.parse(depWeather.data.data.metarDecode) : null,
    }
    const arrMetar = {
      raw: arrWeather?.data?.data?.metar || '获取失败',
      decoded: arrWeather?.data?.data?.metarDecode ? JSON.parse(arrWeather.data.data.metarDecode) : null,
    }

    const depTafRes = await axios.get(`${HOST}/api/realTimeMap/weatherForecast/${dep}`).catch(() => null)
    const arrTafRes = await axios.get(`${HOST}/api/realTimeMap/weatherForecast/${arr}`).catch(() => null)
    const depTaf = { raw: depTafRes?.data?.data?.taf || '获取失败' }
    const arrTaf = { raw: arrTafRes?.data?.data?.taf || '获取失败' }

    // 生成HTML，示意调用
    const html = renderHtml({
      dep,
      arr,
      cycle,
      depMetar,
      arrMetar,
      depTaf,
      arrTaf,
      routeData,
    })
    mkdirSync(TEMP_DIR, { recursive: true })
    const htmlPath = join(TEMP_DIR, `${dep}-${arr}.html`)
    writeFileSync(htmlPath, html)

    // ✅ 使用 Node 原生 puppeteer
    const browser = await puppeteer.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox'],
    })
    const page = await browser.newPage()
    await page.setViewport({ width: config.imageWidth, height: config.imageHeight })
    await page.goto('file://' + htmlPath, { waitUntil: 'networkidle0' })
    const buffer = await page.screenshot({ type: 'jpeg', quality: config.screenshotquality })

    if (config.pageautoclose) await page.close()
    await browser.close()
    unlinkSync(htmlPath)

    return [
  h.image(buffer, 'image/jpeg'),
  `推荐航路字符串：\n${routeData.route || '无'}`
]

  })

}

// 渲染 HTML 页面内容
function renderHtml(data: any): string {
  const { dep, arr, cycle, depMetar, arrMetar, depTaf, arrTaf, routeData } = data
  const depLat = routeData.waypoints?.[0]?.lat || 0
  const depLon = routeData.waypoints?.[0]?.lon || 0
  const arrLat = routeData.waypoints?.at(-1)?.lat || 0
  const arrLon = routeData.waypoints?.at(-1)?.lon || 0
  const hasWaypoints = Array.isArray(routeData.waypoints) && routeData.waypoints.length > 0
  const showPaginationNote = Array.isArray(routeData.waypoints) && routeData.waypoints.length > 50
  // 将你完整的 HTML 模板粘贴于此，进行字符串插值即可
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>航路查询</title>
  <script src="https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.js"></script>
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.css" />
  <style>
    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
      font-family: Arial, sans-serif;
    }
    
    :root {
      --bg-color: #f0f5ff;
      --card-bg: #ffffff;
      --text-color: #333333;
      --border-color: #e0e0e0;
      --primary: #1a3c6c;
      --secondary: #2196F3;
      --success: #4CAF50;
      --danger: #F44336;
      --warning: #FF9800;
    }
    
    .dark-mode {
      --bg-color: #1e1e2e;
      --card-bg: #2d2d44;
      --text-color: #e0e0e0;
      --border-color: #444466;
      --primary: #4a7bff;
    }
    
    body {
      background: var(--bg-color);
      color: var(--text-color);
      padding: 15px;
      line-height: 1.5;
    }
    
    .container {
      max-width: 1200px;
      margin: 0 auto;
      background: var(--card-bg);
      border-radius: 8px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      overflow: hidden;
    }
    
    .header {
      background: var(--primary);
      color: white;
      padding: 20px;
      text-align: center;
      position: relative;
    }
    
    .title {
      font-size: 24px;
      margin-bottom: 5px;
    }
    
    .subtitle {
      font-size: 16px;
      opacity: 0.9;
      margin-bottom: 10px;
    }
    
    .route-cycle {
      background: rgba(255,255,255,0.2);
      padding: 5px 10px;
      border-radius: 4px;
      display: inline-block;
      font-size: 14px;
    }
    
    .theme-toggle {
      position: absolute;
      top: 20px;
      right: 20px;
      background: rgba(255,255,255,0.2);
      border: none;
      border-radius: 20px;
      padding: 5px 15px;
      color: white;
      cursor: pointer;
      z-index: 1000;
    }
    
    .content-grid {
      display: flex;
      flex-wrap: wrap;
      padding: 15px;
    }
    
    .airport-card {
      flex: 1;
      min-width: 300px;
      padding: 15px;
      margin: 10px;
      border: 1px solid var(--border-color);
      border-radius: 8px;
      background: rgba(233, 244, 255, 0.2);
    }
    
    .card-title {
      font-size: 18px;
      margin-bottom: 15px;
      padding-bottom: 10px;
      border-bottom: 1px solid var(--border-color);
      display: flex;
      align-items: center;
    }
    
    .collapse-btn {
      margin-left: auto;
      cursor: pointer;
      font-size: 16px;
      width: 20px;
      height: 20px;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    
    .info-row {
      display: flex;
      margin-bottom: 8px;
    }
    
    .info-label {
      font-weight: bold;
      min-width: 80px;
    }
    
    .weather-box {
      margin-top: 15px;
      padding: 15px;
      background: rgba(232, 244, 255, 0.3);
      border-radius: 6px;
    }
    
    .weather-title {
      font-weight: bold;
      margin-bottom: 5px;
      color: var(--primary);
    }
    
    .metar, .taf {
      font-family: monospace;
      background: rgba(255,255,255,0.1);
      padding: 10px;
      border-radius: 4px;
      border: 1px solid var(--border-color);
      font-size: 14px;
      margin: 5px 0;
      white-space: pre-wrap;
    }
    
    .weather-details {
      background: rgba(255,255,255,0.1);
      border: 1px solid var(--border-color);
      border-radius: 4px;
      padding: 10px;
      margin-top: 10px;
      font-size: 14px;
    }
    
    .weather-details div {
      margin-bottom: 4px;
    }
    
    .map-container {
      height: 400px;
      margin: 15px;
      border-radius: 8px;
      overflow: hidden;
      border: 1px solid var(--border-color);
      position: relative;
    }
    
    .progress-container {
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      height: 5px;
      background: rgba(0,0,0,0.1);
      z-index: 500;
    }
    
    .progress-bar {
      height: 100%;
      background: var(--primary);
      width: 0%;
      transition: width 0.3s;
    }
    
    .route-summary {
      display: flex;
      justify-content: space-around;
      background: rgba(232, 244, 255, 0.3);
      padding: 15px;
      margin: 15px;
      border-radius: 8px;
      text-align: center;
    }
    
    .summary-item {
      flex: 1;
    }
    
    .summary-value {
      font-size: 24px;
      font-weight: bold;
      color: var(--primary);
      margin-bottom: 5px;
    }
    
    .summary-label {
      font-size: 14px;
      color: var(--text-color);
      opacity: 0.8;
    }
    
    .waypoints-container {
      margin: 15px;
    }
    
    .section-title {
      font-size: 18px;
      margin-bottom: 15px;
      display: flex;
      align-items: center;
      color: var(--primary);
    }
    
    .waypoints-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(140px, 1fr));
      gap: 10px;
    }
    
    .waypoint {
      background: rgba(240, 247, 255, 0.2);
      border: 1px solid var(--border-color);
      border-radius: 6px;
      padding: 10px;
      text-align: center;
      position: relative;
    }
    
    .wpt-index {
      position: absolute;
      top: 5px;
      left: 5px;
      background: var(--primary);
      color: white;
      width: 20px;
      height: 20px;
      border-radius: 50%;
      font-size: 12px;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    
    .wpt-name {
      font-weight: bold;
      margin-bottom: 5px;
      color: var(--primary);
    }
    
    .wpt-info {
      font-size: 12px;
      color: var(--text-color);
      opacity: 0.8;
      margin-bottom: 3px;
    }
    
    .wpt-coord {
      font-size: 11px;
      color: var(--text-color);
      opacity: 0.7;
    }
    
    .pagination-note {
      grid-column: 1 / -1;
      text-align: center;
      padding: 10px;
      font-size: 14px;
      color: var(--text-color);
      opacity: 0.7;
    }
    
    .search-box {
      display: flex;
      margin: 15px;
    }
    
    #waypointSearch {
      flex: 1;
      padding: 8px 12px;
      border: 1px solid var(--border-color);
      border-radius: 4px 0 0 4px;
      background: rgba(255,255,255,0.1);
      color: var(--text-color);
    }
    
    #searchButton {
      padding: 0 15px;
      background: var(--primary);
      color: white;
      border: none;
      border-radius: 0 4px 4px 0;
      cursor: pointer;
    }
    
    .footer {
      text-align: center;
      padding: 15px;
      background: rgba(245, 249, 255, 0.2);
      border-top: 1px solid var(--border-color);
      font-size: 12px;
      color: var(--text-color);
      opacity: 0.7;
    }
    
    .error-box {
      background: rgba(255,235,238,0.5);
      border: 1px solid #ffcdd2;
      border-radius: 4px;
      padding: 10px;
      margin: 10px 0;
      color: #b71c1c;
      font-family: monospace;
      white-space: pre-wrap;
    }
    
    .alt { color: var(--success); }
    .freq { color: var(--secondary); }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1 class="title">航路查询</h1>
      <div class="subtitle">${dep} → ${arr}</div>
      <div class="route-cycle">航行资料周期: ${cycle}</div>
      <button class="theme-toggle" id="themeToggle">🌙 深色模式</button>
    </div>
    
    <div class="content-grid">
      <div class="airport-card">
        <h2 class="card-title">
          起飞机场
          <span class="collapse-btn">−</span>
        </h2>
        <div class="airport-info">
          <div class="info-row">
            <span class="info-label">ICAO:</span>
            <span>${dep}</span>
          </div>
          
          <div class="weather-box">
            <h3 class="weather-title">METAR 实时气象</h3>
            ${(String(depMetar.raw).includes('失败') || String(depMetar.raw).includes('不可用')) ? 
              `<div class="error-box">${depMetar.raw}</div>` : 
              `<div class="metar">${depMetar.raw}</div>`}
            
            ${depMetar.decoded ? `
              <div class="weather-details">
                <div><strong>气象解析:</strong></div>
                <div>时间: ${depMetar.decoded.time || 'N/A'}</div>
                <div>风向: ${depMetar.decoded.wind_dir || 'N/A'}°</div>
                <div>风速: ${depMetar.decoded.wind_speed || 'N/A'} ${depMetar.decoded.wind_unit || ''}</div>
                <div>温度: ${depMetar.decoded.temperature || 'N/A'}°C</div>
                <div>露点: ${depMetar.decoded.dewpoint || 'N/A'}°C</div>
                <div>修正海压: ${depMetar.decoded.qnh || 'N/A'} ${depMetar.decoded.qnh_unit || ''}</div>
                <div>能见度: ${depMetar.decoded.visibility || 'N/A'} ${depMetar.decoded.visibility_unit || ''}</div>
                <div>预报: ${depMetar.decoded.forcast || 'N/A'}</div>
              </div>
            ` : ''}
            
            <h3 class="weather-title">TAF 气象预报</h3>
            ${(String(depTaf.raw).includes('失败') || String(depTaf.raw).includes('不可用')) ? 
              `<div class="error-box">${depTaf.raw}</div>` : 
              `<div class="taf">${depTaf.raw}</div>`}
          </div>
        </div>
      </div>
      
      <div class="airport-card">
        <h2 class="card-title">
          目的地机场
          <span class="collapse-btn">−</span>
        </h2>
        <div class="airport-info">
          <div class="info-row">
            <span class="info-label">ICAO:</span>
            <span>${arr}</span>
          </div>
          
          <div class="weather-box">
            <h3 class="weather-title">METAR 实时气象</h3>
            ${(String(arrMetar.raw).includes('失败') || String(arrMetar.raw).includes('不可用')) ? 
              `<div class="error-box">${arrMetar.raw}</div>` : 
              `<div class="metar">${arrMetar.raw}</div>`}
            
            ${arrMetar.decoded ? `
              <div class="weather-details">
                <div><strong>气象解析:</strong></div>
                <div>时间: ${arrMetar.decoded.time || 'N/A'}</div>
                <div>风向: ${arrMetar.decoded.wind_dir || 'N/A'}°</div>
                <div>风速: ${arrMetar.decoded.wind_speed || 'N/A'} ${arrMetar.decoded.wind_unit || ''}</div>
                <div>温度: ${arrMetar.decoded.temperature || 'N/A'}°C</div>
                <div>露点: ${arrMetar.decoded.dewpoint || 'N/A'}°C</div>
                <div>修正海压: ${arrMetar.decoded.qnh || 'N/A'} ${arrMetar.decoded.qnh_unit || ''}</div>
                <div>能见度: ${arrMetar.decoded.visibility || 'N/A'} ${arrMetar.decoded.visibility_unit || ''}</div>
                <div>预报: ${arrMetar.decoded.forcast || 'N/A'}</div>
              </div>
            ` : ''}
            
            <h3 class="weather-title">TAF 气象预报</h3>
            ${(String(arrTaf.raw).includes('失败') || String(arrTaf.raw).includes('不可用')) ? 
              `<div class="error-box">${arrTaf.raw}</div>` : 
              `<div class="taf">${arrTaf.raw}</div>`}
          </div>
        </div>
      </div>
    </div>
    
    <div class="route-summary">
  <div class="summary-item">
    <div class="summary-value">${routeData.distance || 0}</div>
    <div class="summary-label">总距离 </div>
  </div>
  <div class="summary-item">
    <div class="summary-value">${routeData.flight_time || 'N/A'}</div>
    <div class="summary-label">预计时间</div>
  </div>
  <div class="summary-item">
    <div class="summary-value">${routeData.waypoints?.length || 0}</div>
    <div class="summary-label">航路点数量</div>
  </div>
</div>

<!-- ✅ 新增：推荐航路字符串 -->
<div class="route-string-box" style="margin: 10px 20px; padding: 10px; background: rgba(0,0,0,0.05); border-radius: 8px; font-family: monospace; white-space: pre-wrap;">
  <strong style="display: block; margin-bottom: 5px;">推荐航路字符串：</strong>
  ${routeData.route || '无'}
</div>

<div id="map" class="map-container">
  <div class="progress-container">
    <div class="progress-bar" id="progressBar"></div>
  </div>

  ${!hasWaypoints ? `
  <div class="error-box" style="margin: 15px;">
    警告: 未获取到航路点数据，地图显示为默认位置
  </div>
  ` : ''}

  ${hasWaypoints ? `
  <div class="waypoints-container">
    <h2 class="section-title">航路点信息</h2>
    <div class="waypoints-grid">
      ${routeData.waypoints.slice(0, 50).map((wp, index) => `
        <div class="waypoint">
          <div class="wpt-index">${index + 1}</div>
          <div class="wpt-name">${wp.name}</div>
          <div class="wpt-info">
            ${wp.alt ? `<span class="alt">FL${wp.alt}</span>` : ''}
            ${wp.freq ? `<span class="freq">${wp.freq}MHz</span>` : ''}
          </div>
          <div class="wpt-coord">${wp.lat.toFixed(4)}, ${wp.lon.toFixed(4)}</div>
        </div>
      `).join('')}

      ${showPaginationNote ? `
        <div class="pagination-note">
          显示前 50 个航路点（共 ${routeData.waypoints.length} 个）
        </div>
      ` : ''}
    </div>
  </div>
  ` : ''}
</div>


    
    ${hasWaypoints ? `
    <div class="search-box">
      <input type="text" id="waypointSearch" placeholder="搜索航路点...">
      <button id="searchButton">搜索</button>
    </div>
    ` : ''}
    </div>
    
    <div class="footer">
      <p>Generated by Flight Route Plugin | ${new Date().toLocaleString()}</p>
    </div>
  </div>

  <script>
    // 初始化地图
    const depLat = ${depLat};
    const depLon = ${depLon};
    const arrLat_js = ${arrLat};
    const arrLon_js = ${arrLon};
    const waypoints = ${JSON.stringify(routeData.waypoints || [])};
    
    // 创建地图
    window.map = undefined;
    let map;
    try {
      map = L.map('map').setView([(depLat + arrLat_js) / 2, (depLon + arrLon_js) / 2], 4);
      window.map = map;
      
      // 添加地图图层
      L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
        maxZoom: 18
      }).addTo(map);
      
      // 添加起飞机场标记
      const depMarker = L.marker([depLat, depLon], {
        icon: L.divIcon({
          html: '<div style="background: #4CAF50; color: white; border-radius: 50%; width: 36px; height: 36px; display: flex; align-items: center; justify-content: center; font-weight: bold; border: 2px solid white;">DEP</div>',
          className: '',
          iconSize: [36, 36]
        })
      }).addTo(map).bindPopup('${dep}<br>起飞机场');
      
      // 添加目的地机场标记
      const arrMarker = L.marker([arrLat_js, arrLon_js], {
        icon: L.divIcon({
          html: '<div style="background: #F44336; color: white; border-radius: 50%; width: 36px; height: 36px; display: flex; align-items: center; justify-content: center; font-weight: bold; border: 2px solid white;">ARR</div>',
          className: '',
          iconSize: [36, 36]
        })
      }).addTo(map).bindPopup('${arr}<br>目的地机场');
      
      // 平滑曲线函数（使用贝塞尔曲线）
      function getCurvePoints(start, end, waypoints) {
        const points = [start];
        const numSegments = waypoints.length + 1;
        
        for (let i = 0; i <= numSegments; i++) {
          const t = i / numSegments;
          const lat = start[0] + (end[0] - start[0]) * t;
          const lon = start[1] + (end[1] - start[1]) * t;
          
          // 添加中间控制点偏移
          if (i > 0 && i < numSegments) {
            const wp = waypoints[i - 1];
            points.push([lat + (wp.lat - lat) * 0.7, lon + (wp.lon - lon) * 0.7]);
          }
        }
        
        points.push(end);
        return points;
      }
      
      // 创建航路点标记和连线
      const waypointMarkers = [];
      let routePolyline = null;
      
      // 更新进度条
      function updateProgress(percent) {
        document.getElementById('progressBar').style.width = percent + '%';
      }
      
      // 绘制航路
      function drawRoute() {
        updateProgress(0);
        
        // 清除现有路线
        if (routePolyline) {
          map.removeLayer(routePolyline);
        }
        
        // 清除现有航路点标记
        waypointMarkers.forEach(marker => map.removeLayer(marker));
        waypointMarkers.length = 0;
        
        const startPoint = [depLat, depLon];
        const endPoint = [arrLat_js, arrLon_js];
        let routePoints = [];
        
        if (waypoints.length > 0) {
          // 使用平滑曲线
          routePoints = getCurvePoints(startPoint, endPoint, waypoints);
          
          // 添加航路点标记
          for (let i = 0; i < waypoints.length; i++) {
            updateProgress((i / waypoints.length) * 70);
            
            const wp = waypoints[i];
            const isFirst = i === 0;
            const isLast = i === waypoints.length - 1;
            let color = '#2196F3';
            
            if (isFirst) color = '#4CAF50';
            if (isLast) color = '#F44336';
            
            const marker = L.marker([wp.lat, wp.lon], {
              icon: L.divIcon({
                html: \`<div style="background: \${color}; color: white; border-radius: 50%; width: 24px; height: 24px; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: bold; border: 2px solid white;">\${wp.name}</div>\`,
                className: '',
                iconSize: [24, 24]
              })
            }).addTo(map);
            
            // 创建弹窗内容
            let popupContent = \`<b>\${wp.name}</b><br>\`;
            popupContent += \`坐标: \${wp.lat.toFixed(4)}, \${wp.lon.toFixed(4)}<br>\`;
            if (wp.alt) popupContent += \`高度: FL\${wp.alt}<br>\`;
            if (wp.freq) popupContent += \`频率: \${wp.freq}MHz\`;
            
            marker.bindPopup(popupContent);
            waypointMarkers.push(marker);
          }
          
          // 添加动画路径
          routePolyline = L.polyline(routePoints, {
            color: '#2196F3',
            weight: 3,
            opacity: 0.8,
            dashArray: '10, 10',
            lineCap: 'round'
          }).addTo(map);
          
          // 添加路径动画
          let animatedIndex = 0;
          const animatePath = () => {
            if (animatedIndex <= routePoints.length) {
              const partialPath = routePoints.slice(0, animatedIndex);
              if (routePolyline) {
                map.removeLayer(routePolyline);
              }
              
              routePolyline = L.polyline(partialPath, {
                color: '#2196F3',
                weight: 3,
                opacity: 0.8
              }).addTo(map);
              
              animatedIndex++;
              setTimeout(animatePath, 30);
              updateProgress(70 + (animatedIndex / routePoints.length) * 30);
            }
          };
          
          setTimeout(animatePath, 500);
        } else {
          // 没有航路点时绘制直线
          routePolyline = L.polyline([startPoint, endPoint], {
            color: '#FF9800',
            weight: 2,
            dashArray: '10, 10',
          }).addTo(map).bindPopup('未获取到航路点，显示直线路径');
          updateProgress(100);
        }
      }
      
      // 初始绘制
      drawRoute();
      
      // 添加缩放控件
      L.control.zoom({ position: 'topright' }).addTo(map);
      
      // 调整视图
      function adjustView() {
        try {
          map.invalidateSize();
          if (waypoints.length > 0) {
            const bounds = new L.LatLngBounds();
            bounds.extend([depLat, depLon]);
            bounds.extend([arrLat_js, arrLon_js]);
            map.fitBounds(bounds, { padding: [50, 50] });
          }
          updateProgress(100);
        } catch (e) {
          console.error('地图调整失败:', e);
        }
      }
      
      // 使用防抖优化窗口调整
      let resizeTimer;
      window.addEventListener('resize', () => {
        clearTimeout(resizeTimer);
        resizeTimer = setTimeout(adjustView, 250);
      });
      
      // 初始调整视图
      setTimeout(adjustView, 500);
    } catch (e) {
      console.error('地图初始化失败:', e);
      document.getElementById('map').innerHTML = '<div class="error-box">地图加载失败: ' + e.message + '</div>';
    }
    
    // 主题切换功能
    const themeToggle = document.getElementById('themeToggle');
    themeToggle.addEventListener('click', () => {
      document.body.classList.toggle('dark-mode');
      themeToggle.textContent = document.body.classList.contains('dark-mode') 
        ? '☀️ 浅色模式' 
        : '🌙 深色模式';
    });
    
    // 折叠功能
    document.querySelectorAll('.collapse-btn').forEach(btn => {
      btn.addEventListener('click', function() {
        const card = this.closest('.airport-card');
        const content = card.querySelector('.airport-info');
        const isCollapsed = content.style.display === 'none';
        
        content.style.display = isCollapsed ? 'block' : 'none';
        this.textContent = isCollapsed ? '−' : '+';
      });
    });
    
    // 航路点搜索功能
    if (document.getElementById('waypointSearch')) {
      const searchInput = document.getElementById('waypointSearch');
      const searchButton = document.getElementById('searchButton');
      
      function searchWaypoint() {
        const searchTerm = searchInput.value.trim().toUpperCase();
        if (!searchTerm) return;
        
        // 清除之前的高亮
        waypointMarkers.forEach(marker => {
          const originalHtml = marker.getPopup().getContent().match(/<div style="[^"]*">(.*?)<\/div>/)[1];
          marker.setIcon(L.divIcon({
            html:
              '<div style="background: ' +
              marker.options.icon.options.html.match(/background: ([^;]+)/)[1] +
              '; color: white; border-radius: 50%; width: 24px; height: 24px; display: flex; align-items: center; ' +
              'justify-content: center; font-size: 12px; font-weight: bold; border: 2px solid white;">' +
              originalHtml +
              '</div>',
            className: '',
            iconSize: [24, 24]
          }));
        });
        
        // 查找匹配的航路点
        const found = waypointMarkers.find(function(marker) {
          const markerName = marker.getPopup().getContent().match(/<b>(.*?)<\/b>/)[1];
          return markerName === searchTerm;
        });
        
        if (found) {
          // 高亮显示
          found.setIcon(L.divIcon({
            html: '<div style="background: #FFEB3B; color: #333; border-radius: 50%; width: 28px; height: 28px; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: bold; border: 2px solid white; box-shadow: 0 0 10px yellow;">' + searchTerm + '</div>',
            className: '',
            iconSize: [28, 28]
          }));
          (window.map || map).setView(found.getLatLng(), 8);
          found.openPopup();
          map.setView(found.getLatLng(), 8);
          found.openPopup();
        } else {
          alert('未找到航路点: ' + searchTerm);
        }
      }
      
      searchButton.addEventListener('click', searchWaypoint);
      searchInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') searchWaypoint();
      });
    }
  </script>
</body>
</html>`;
}

async function generateImagePage(html: string, config: Config): Promise<Buffer> {
  const tempDir = join(tmpdir(), 'koishi-route-search')
  mkdirSync(tempDir, { recursive: true })
  const filePath = join(tempDir, `route-${Date.now()}.html`)
  writeFileSync(filePath, html)

  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  })
  const page = await browser.newPage()
  await page.setViewport({
    width: config.imageWidth || 1200,
    height: config.imageHeight || 800,
  })
  await page.goto('file://' + filePath, { waitUntil: 'networkidle0' })

  const screenshotResult = await page.screenshot({ type: 'jpeg', quality: config.screenshotquality || 80 })
  await page.close()
  await browser.close()
  unlinkSync(filePath)

  // Ensure the result is a Node.js Buffer
  const buffer = Buffer.isBuffer(screenshotResult) ? screenshotResult : Buffer.from(screenshotResult as Uint8Array)
  return buffer
}

