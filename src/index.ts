import { Context, Schema, Logger, h, $ } from 'koishi'
import {} from '@shangxueink/koishi-plugin-puppeteer-without-canvas'
import { readFileSync, writeFileSync, mkdirSync, unlinkSync } from 'fs'
import { join } from 'path'
import { tmpdir } from 'os'
import axios from 'axios'
import XLSX from 'xlsx'

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
  mapboxAccessToken: string
}

interface RouteWaypoint {
  name: string
  lat: number
  lon: number
  alt?: number | string
  freq?: number | string
  [key: string]: any
}

interface RouteData {
  route?: string
  distance?: number
  flight_time?: string
  waypoints: RouteWaypoint[]
  SID?: any
  STAR?: any
  [key: string]: any
}

interface MetarDecoded {
  time?: string
  wind_dir?: string
  wind_speed?: string
  wind_unit?: string
  temperature?: string
  dewpoint?: string
  qnh?: string
  qnh_unit?: string
  visibility?: string
  visibility_unit?: string
  forecast?: string
}

interface WeatherRecord {
  raw: string
  decoded: MetarDecoded | null
}

interface TafForecastEntry {
  wx_string?: string
  fcst_time_from?: string
  fcst_time_to?: string
  change_indicator?: string
  sky_condition?: unknown
  wind_dir_degrees?: string
  wind_speed_kt?: string
  visibility_statute_mi?: string
}

interface TafInfo {
  raw: string
  forecastSummary?: string
}

interface RenderHtmlData {
  dep: string
  arr: string
  cycle: string
  depMetar: WeatherRecord
  arrMetar: WeatherRecord
  depTaf: TafInfo
  arrTaf: TafInfo
  routeData: RouteData
}

export const name = 'route-search'
export const inject = ['puppeteer', 'canvas']

export const Config: Schema<Config> = Schema.object({
  commandname: Schema.string().default('航路查询').description('插件指令名'),
  pageautoclose: Schema.boolean().default(true).description('截图后自动关闭页面'),
  imageWidth: Schema.number().default(1280).description('截图基准宽度（将按内容自动扩展）'),
  imageHeight: Schema.number().default(720).description('截图基准高度（完整内容会自动截取）'),
  screenshotquality: Schema.number().default(80).description('截图质量 (0-100)'),
  weatherMap: Schema.array(Schema.object({
    code: Schema.string().required(),
    description: Schema.string().required(),
  })).role('table').default([/* ... your default weather map here ... */]),
  cloudCoverageMap: Schema.array(Schema.object({
    code: Schema.string().required(),
    description: Schema.string().required(),
  })).role('table').default([/* ... your default cloud coverage map here ... */]),
  mapboxAccessToken: Schema.string().role('secret').default('').description('Mapbox access token（留空则不显示地图）'),
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
    const cycleRes = await axios.get(`${HOST}/api/efb/getRouteCycle`).catch(() => null)
    if (!cycleRes?.data?.data) return '获取 CYCLE 数据失败'

    let cycles: string[] = []
    try {
      cycles = Array.isArray(cycleRes.data.data) ? cycleRes.data.data : JSON.parse(cycleRes.data.data)
    } catch {
      return '解析 CYCLE 数据失败'
    }
    if (!Array.isArray(cycles) || cycles.length === 0) return '获取 CYCLE 数据失败'

    const maxCycle = Math.max(...cycles.map(c => Number(c)))
    const cycle = maxCycle.toString()

    // 请求航路
    const routeRes = await axios.post(`${HOST}/api/efb/route`, null, {
      params: { dep, arr, cycle },
      timeout: 10000,
    }).catch(() => null)

    if (!routeRes || routeRes.data.code !== 20000) return '获取航路失败'

    // 解析航路数据
    let routeDataRaw: Record<string, any>
    try {
      routeDataRaw = typeof routeRes.data.data === 'string' ? JSON.parse(routeRes.data.data) : routeRes.data.data
    } catch {
      return '解析航路数据失败'
    }

    // 构造waypoints数组
    const rawNodes = Array.isArray(routeDataRaw.nodeinformation) ? routeDataRaw.nodeinformation : []
    const waypoints: RouteWaypoint[] = rawNodes.map((item: any[]) => {
      const latValue = Number(item[1])
      const lonValue = Number(item[2])
      return {
        name: item[0] || '',
        lat: Number.isFinite(latValue) ? latValue : 0,
        lon: Number.isFinite(lonValue) ? lonValue : 0,
        alt: item[4],
        freq: item[5],
      }
    })

    // nodeinformation 缺失时用 parseRoute 补齐航路点坐标
    if (!waypoints.length && routeDataRaw.route) {
      const routeText = Array.isArray(routeDataRaw.route) ? routeDataRaw.route.join(' ') : String(routeDataRaw.route)
      const parsed = await axios.post(`${HOST}/api/efb/parseRoute`, null, {
        params: { route: routeText, dep, arr, cycle },
        timeout: 10000,
      }).catch(() => null)
      const points = parsed?.data?.code === 20000 ? (parsed.data.data?.waypoints || []) : []
      for (const point of points) {
        const latValue = Number(point.lat)
        const lonValue = Number(point.lon)
        if (Number.isFinite(latValue) && Number.isFinite(lonValue)) {
          waypoints.push({ name: point.name || '', lat: latValue, lon: lonValue })
        }
      }
    }

    // 组合数据
    const routeData: RouteData = {
      ...routeDataRaw,
      waypoints,
    }

    // 获取天气信息
    const depWeather = await axios.get(`${HOST}/api/efb/weather/metar/${dep}`).catch(() => null)
    const arrWeather = await axios.get(`${HOST}/api/efb/weather/metar/${arr}`).catch(() => null)
    const depMetar = parseXfsWeatherResponse(depWeather?.data)
    const arrMetar = parseXfsWeatherResponse(arrWeather?.data)

    const depTafRes = await axios.get(`${HOST}/api/efb/weather/taf/${dep}`).catch(() => null)
    const arrTafRes = await axios.get(`${HOST}/api/efb/weather/taf/${arr}`).catch(() => null)
    const depTaf = parseXfsTafResponse(depTafRes?.data)
    const arrTaf = parseXfsTafResponse(arrTafRes?.data)

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
    }, config.mapboxAccessToken || '')
    mkdirSync(TEMP_DIR, { recursive: true })
    const htmlPath = join(TEMP_DIR, `${dep}-${arr}.html`)
    writeFileSync(htmlPath, html)

    const page = await ctx.puppeteer.page()
    let buffer: Uint8Array
    try {
      await page.setViewport({ width: Math.max(320, config.imageWidth), height: Math.max(320, config.imageHeight) })
      await page.goto('file://' + htmlPath, { waitUntil: 'networkidle0' })
      await page.waitForFunction('window.__mapReady === true', { timeout: 20000 }).catch(() => {})
      const contentWidth = await page.evaluate(() => Math.ceil(Math.max(
        document.documentElement.scrollWidth,
        document.body.scrollWidth,
      )))
      await page.setViewport({
        width: Math.min(Math.max(320, config.imageWidth, contentWidth), 4096),
        height: Math.max(320, config.imageHeight),
      })
      await page.evaluate(() => window.dispatchEvent(new Event('resize')))
      await page.evaluate(async () => {
        await document.fonts?.ready
        await new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve())))
      })
      buffer = await page.screenshot({ type: 'jpeg', quality: config.screenshotquality, fullPage: true })
    } finally {
      await page.close().catch(() => {})
      try { unlinkSync(htmlPath) } catch {}
    }

    return [
  h.image(Buffer.from(buffer), 'image/jpeg'),
  `推荐航路字符串：\n${routeData.route || '无'}`
]

  })

}

// 渲染 HTML 页面内容
function renderHtml(data: RenderHtmlData, mapboxToken = ''): string {
  const { dep, arr, cycle, depMetar, arrMetar, depTaf, arrTaf, routeData } = data
  const waypointList = routeData.waypoints || []
  const depLat = waypointList[0]?.lat || 0
  const depLon = waypointList[0]?.lon || 0
  const arrLat = waypointList.at(-1)?.lat || 0
  const arrLon = waypointList.at(-1)?.lon || 0
  const hasWaypoints = waypointList.length > 0
  const showPaginationNote = waypointList.length > 50
  // 将你完整的 HTML 模板粘贴于此，进行字符串插值即可
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>航路查询</title>
  <link href="https://api.mapbox.com/mapbox-gl-js/v3.12.0/mapbox-gl.css" rel="stylesheet" />
  <script src="https://api.mapbox.com/mapbox-gl-js/v3.12.0/mapbox-gl.js"></script>
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
                <div>预报: ${depMetar.decoded.forecast || 'N/A'}</div>
              </div>
            ` : ''}
            
            <h3 class="weather-title">TAF 气象预报</h3>
            ${(String(depTaf.raw).includes('失败') || String(depTaf.raw).includes('不可用')) ? 
              `<div class="error-box">${depTaf.raw}</div>` : 
              `<div class="taf">${depTaf.raw}</div>`}
            ${depTaf.forecastSummary ? `<div class="weather-details"><div><strong>TAF 分段:</strong></div>${depTaf.forecastSummary}</div>` : ''}
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
                <div>预报: ${arrMetar.decoded.forecast || 'N/A'}</div>
              </div>
            ` : ''}
            
            <h3 class="weather-title">TAF 气象预报</h3>
            ${(String(arrTaf.raw).includes('失败') || String(arrTaf.raw).includes('不可用')) ? 
              `<div class="error-box">${arrTaf.raw}</div>` : 
              `<div class="taf">${arrTaf.raw}</div>`}
            ${arrTaf.forecastSummary ? `<div class="weather-details"><div><strong>TAF 分段:</strong></div>${arrTaf.forecastSummary}</div>` : ''}
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
      <div class="summary-value">${waypointList.length}</div>
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
      ${waypointList.slice(0, 50).map((wp, index) => `
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
          显示前 50 个航路点（共 ${waypointList.length} 个）
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
    window.__mapReady = false;
    (function () {
      const depIcao = ${JSON.stringify(dep)};
      const arrIcao = ${JSON.stringify(arr)};
      const depLat = ${depLat}, depLon = ${depLon};
      const arrLat = ${arrLat}, arrLon = ${arrLon};
      const waypoints = ${JSON.stringify(waypointList)};
      const TOKEN = ${JSON.stringify(mapboxToken || '')};
      const mapEl = document.getElementById('map');

      function finish() {
        if (!window.__mapReady) window.__mapReady = true;
      }
      function setProgress(percent) {
        const bar = document.getElementById('progressBar');
        if (bar) bar.style.width = percent + '%';
      }
      function loadMapbox() {
        if (window.mapboxgl) return Promise.resolve();
        const sources = [
          'https://api.mapbox.com/mapbox-gl-js/v3.12.0/mapbox-gl.js',
          'https://cdn.jsdelivr.net/npm/mapbox-gl@3.12.0/dist/mapbox-gl.js',
          'https://unpkg.com/mapbox-gl@3.12.0/dist/mapbox-gl.js',
        ];
        return new Promise(function (resolve, reject) {
          let index = 0;
          (function next() {
            if (index >= sources.length) return reject(new Error('Mapbox 脚本加载失败'));
            const script = document.createElement('script');
            script.src = sources[index++];
            script.async = true;
            script.onload = function () { window.mapboxgl ? resolve() : next(); };
            script.onerror = next;
            document.head.appendChild(script);
          })();
        });
      }

      loadMapbox().then(function () {
        if (!TOKEN) {
          mapEl.innerHTML = '<div class="error-box">未配置 Mapbox access token</div>';
          setProgress(100);
          finish();
          return;
        }
        mapboxgl.accessToken = TOKEN;
        const map = new mapboxgl.Map({
          container: mapEl,
          style: 'mapbox://styles/mapbox/dark-v11',
          center: [(depLon + arrLon) / 2, (depLat + arrLat) / 2],
          zoom: 4,
          preserveDrawingBuffer: true,
        });
        window.map = map;
        const markerEntries = [];

        function addMarker(name, lat, lon, color) {
          const element = document.createElement('div');
          element.textContent = name;
          element.style.cssText = 'background:' + color + ';color:#fff;border-radius:50%;width:26px;height:26px;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:bold;border:2px solid #fff;';
          const marker = new mapboxgl.Marker({ element: element, anchor: 'center' })
            .setLngLat([lon, lat]).addTo(map);
          markerEntries.push({ name: name, marker: marker });
          return marker;
        }

        map.on('load', function () {
          const coords = waypoints
            .filter(function (point) { return isFinite(point.lat) && isFinite(point.lon) && !(point.lat === 0 && point.lon === 0); })
            .map(function (point) { return [point.lon, point.lat]; });

          if (coords.length >= 2) {
            map.addSource('route', {
              type: 'geojson',
              data: { type: 'Feature', geometry: { type: 'LineString', coordinates: coords } },
            });
            map.addLayer({
              id: 'route-line', type: 'line', source: 'route',
              layout: { 'line-cap': 'round', 'line-join': 'round' },
              paint: { 'line-color': '#2196F3', 'line-width': 3 },
            });
          }

          if (isFinite(depLat) && isFinite(depLon)) addMarker(depIcao, depLat, depLon, '#4CAF50');
          if (isFinite(arrLat) && isFinite(arrLon)) addMarker(arrIcao, arrLat, arrLon, '#F44336');

          const seen = {};
          waypoints.forEach(function (point) {
            if (!isFinite(point.lat) || !isFinite(point.lon)) return;
            if (point.lat === 0 && point.lon === 0) return;
            if (!point.name || point.name === depIcao || point.name === arrIcao) return;
            if (seen[point.name]) return;
            seen[point.name] = true;
            addMarker(point.name, point.lat, point.lon, '#2196F3');
          });

          const bounds = new mapboxgl.LngLatBounds();
          if (isFinite(depLat) && isFinite(depLon)) bounds.extend([depLon, depLat]);
          if (isFinite(arrLat) && isFinite(arrLon)) bounds.extend([arrLon, arrLat]);
          coords.forEach(function (coord) { bounds.extend(coord); });
          try { map.fitBounds(bounds, { padding: 50, maxZoom: 9, duration: 0 }); } catch (e) {}

          setProgress(100);
          map.once('idle', finish);
        });

        map.on('error', function () { setProgress(100); finish(); });

        if (document.getElementById('waypointSearch')) {
          const searchInput = document.getElementById('waypointSearch');
          const searchButton = document.getElementById('searchButton');
          function searchWaypoint() {
            const term = (searchInput.value || '').trim().toUpperCase();
            if (!term) return;
            const found = markerEntries.find(function (entry) { return entry.name === term; });
            if (found) {
              map.flyTo({ center: found.marker.getLngLat(), zoom: 8, duration: 0 });
            } else {
              alert('未找到航路点: ' + term);
            }
          }
          searchButton.addEventListener('click', searchWaypoint);
          searchInput.addEventListener('keypress', function (event) {
            if (event.key === 'Enter') searchWaypoint();
          });
        }
      }).catch(function (error) {
        mapEl.innerHTML = '<div class="error-box">地图加载失败: ' + error.message + '</div>';
        setProgress(100);
        finish();
      });

      // 主题切换功能
      const themeToggle = document.getElementById('themeToggle');
      if (themeToggle) {
        themeToggle.addEventListener('click', function () {
          document.body.classList.toggle('dark-mode');
          themeToggle.textContent = document.body.classList.contains('dark-mode')
            ? '☀️ 浅色模式'
            : '🌙 深色模式';
        });
      }

      // 折叠功能
      document.querySelectorAll('.collapse-btn').forEach(function (button) {
        button.addEventListener('click', function () {
          const card = this.closest('.airport-card');
          const content = card.querySelector('.airport-info');
          const isCollapsed = content.style.display === 'none';
          content.style.display = isCollapsed ? 'block' : 'none';
          this.textContent = isCollapsed ? '−' : '+';
        });
      });

      setTimeout(finish, 15000);
    })();
  </script>
</body>
</html>`;
}

function formatSkyCondition(value?: unknown): string {
  if (!value) return ''
  if (typeof value === 'string') {
    const trimmed = value.trim()
    if (!trimmed) return ''
    try {
      const parsed = JSON.parse(trimmed)
      if (Array.isArray(parsed)) {
        return describeSkyEntries(parsed)
      }
      if (parsed && typeof parsed === 'object') {
        return describeSkyEntries([parsed])
      }
    } catch {
      // ignore
    }
    return trimmed
  }
  if (Array.isArray(value)) return describeSkyEntries(value)
  if (value && typeof value === 'object') return describeSkyEntries([value])
  return ''
}

function describeSkyEntries(entries: unknown[]): string {
  const parts = entries
    .map(entry => {
      if (!entry || typeof entry !== 'object') return ''
      const payload = entry as Record<string, unknown>
      const cover = String(payload.sky_cover || payload.skyCover || payload.cover || '').trim()
      if (!cover) return ''
      const base = String(payload.cloud_base_ft_agl || payload.cloud_base || '').trim()
      return base ? `${cover} ${base}ft` : cover
    })
    .filter(Boolean)
  return Array.from(new Set(parts)).join(', ')
}

function parseForecastEntries(value?: unknown): TafForecastEntry[] {
  if (!value) return []
  if (Array.isArray(value)) return value
  if (typeof value === 'string') {
    try {
      const parsed = JSON.parse(value)
      if (Array.isArray(parsed)) return parsed
    } catch {
      // ignore
    }
  }
  return []
}

function formatForecastEntry(entry: TafForecastEntry): string {
  const indicator = entry.change_indicator ? `${entry.change_indicator} ` : ''
  const from = entry.fcst_time_from
  const to = entry.fcst_time_to
  let timeRange = ''
  if (from && to) {
    timeRange = `${from} ~ ${to}`
  } else if (from) {
    timeRange = from
  } else if (to) {
    timeRange = to
  }

  const details: string[] = []
  if (entry.wx_string) details.push(entry.wx_string)
  const skyDesc = formatSkyCondition(entry.sky_condition)
  if (skyDesc) details.push(skyDesc)
  if (entry.visibility_statute_mi) details.push(`能见度 ${entry.visibility_statute_mi} mi`)

  const windParts: string[] = []
  if (entry.wind_dir_degrees) windParts.push(`${entry.wind_dir_degrees}°`)
  if (entry.wind_speed_kt) windParts.push(`${entry.wind_speed_kt}kt`)
  if (windParts.length) details.push(windParts.join(' '))

  const labelParts = [indicator.trim(), timeRange].filter(Boolean)
  const label = labelParts.join(' ').trim()
  const detailString = details.filter(Boolean).join(' | ')

  if (!label && !detailString) return ''
  if (!label) return detailString
  if (!detailString) return label
  return `${label}：${detailString}`
}

function parseXfsWeatherResponse(response?: any): WeatherRecord {
  if (!response || response.code !== 20000 || !response.data) {
    return { raw: '获取失败', decoded: null }
  }
  const data = response.data
  const raw = data.rawText || data.metar || '暂无数据'
  const skyText = formatSkyCondition(data.skyCondition) || ''
  const forecastText = skyText || (data.wxString || '')
  return {
    raw,
    decoded: {
      time: data.observationTime,
      wind_dir: data.windDirDegrees,
      wind_speed: data.windSpeedKt,
      wind_unit: data.windSpeedKt ? 'kt' : undefined,
      temperature: data.tempC,
      dewpoint: data.dewpointC,
      qnh: data.altimInHg,
      qnh_unit: data.altimInHg ? 'inHg' : undefined,
      visibility: data.visibilityStatuteMi,
      visibility_unit: data.visibilityStatuteMi ? 'SM' : undefined,
      forecast: forecastText ? forecastText.trim() : undefined,
    },
  }
}

function parseXfsTafResponse(response?: any): TafInfo {
  if (!response || response.code !== 20000 || !response.data) {
    return { raw: '获取失败' }
  }
  const data = response.data
  const raw = data.rawText || ''
  const entries = parseForecastEntries(data.forecast)
  const summary = entries
    .map(entry => formatForecastEntry(entry))
    .filter(Boolean)
    .map(line => `<div>${line}</div>`)
    .join('')
  return {
    raw: raw || '暂无数据',
    forecastSummary: summary || undefined,
  }
}

async function generateImagePage(ctx: Context, html: string, config: Config): Promise<Buffer> {
  const tempDir = join(tmpdir(), 'koishi-route-search')
  mkdirSync(tempDir, { recursive: true })
  const filePath = join(tempDir, `route-${Date.now()}.html`)
  writeFileSync(filePath, html)

  const page = await ctx.puppeteer.page()
  let screenshotResult: Uint8Array
  try {
    await page.setViewport({
      width: Math.max(320, config.imageWidth || 1200),
      height: Math.max(320, config.imageHeight || 800),
    })
    await page.goto('file://' + filePath, { waitUntil: 'networkidle0' })

    const contentWidth = await page.evaluate(() => Math.ceil(Math.max(
      document.documentElement.scrollWidth,
      document.body.scrollWidth,
    )))
    await page.setViewport({
      width: Math.min(Math.max(320, config.imageWidth || 1200, contentWidth), 4096),
      height: Math.max(320, config.imageHeight || 800),
    })
    await page.evaluate(() => window.dispatchEvent(new Event('resize')))
    await page.evaluate(async () => {
      await document.fonts?.ready
      await new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve())))
    })
    screenshotResult = await page.screenshot({ type: 'jpeg', quality: config.screenshotquality || 80, fullPage: true })
  } finally {
    await page.close().catch(() => {})
    try { unlinkSync(filePath) } catch {}
  }

  // Ensure the result is a Node.js Buffer
  const buffer = Buffer.isBuffer(screenshotResult) ? screenshotResult : Buffer.from(screenshotResult as Uint8Array)
  return buffer
}
