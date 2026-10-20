const CACHE_MS = 10 * 60 * 1000;
const TIMEOUT_MS = 7000;
const SAMPLE = Object.freeze({
  mode: 'sample', status: '예제 날씨', reason: 'selected-sample', message: '교육용 가상 자료입니다.', cached: false,
  location: { id: 'seoul', name: '서울 중구 참고 좌표', note: '예제A점의 실제 위치가 아닙니다.', lat: 37.5636, lon: 126.9828, timezone: 'Asia/Seoul' },
  source: { name: '교육용 예제 · 실제 관측·예보 아님', url: null },
  fetchedAt: '2026-10-20T03:00:00.000Z', observedAt: '2026-10-20T03:00:00.000Z',
  current: { temp: 17, feelsLike: 16, humidity: 65, wind: 3.2, description: '흐림', conditionId: 804 },
  forecast: [
    { date: '2026-10-20', min: 13, max: 18, pop: 0.2, description: '흐림', conditionId: 804, slots: 4 },
    { date: '2026-10-21', min: 12, max: 17, pop: 0.7, description: '약한 비', conditionId: 500, slots: 8 },
    { date: '2026-10-22', min: 11, max: 18, pop: 0.1, description: '맑음', conditionId: 800, slots: 8 },
    { date: '2026-10-23', min: 12, max: 17, pop: 0.8, description: '비', conditionId: 501, slots: 8 },
    { date: '2026-10-24', min: 13, max: 19, pop: 0.2, description: '구름 조금', conditionId: 801, slots: 8 }
  ]
});

function createWeatherService({ fetchImpl = fetch, now = Date.now, timeoutMs = TIMEOUT_MS } = {}) {
  const cache = new Map();
  const pending = new Map();
  function sample(reason, attemptedAt) {
    const data = SAMPLE;
    return { ...data, mode: 'sample', reason, attemptedAt, cached: false, connection: 'sample-fallback', source: { ...data.source, file: 'data/11_날씨응답_예제.json' } };
  }
  function safeError(code) {
    return ({ MISSING_KEY: '서버에 OPENWEATHER_API_KEY가 설정되지 않았습니다.',
      TIMEOUT: '날씨 조회 제한 시간(7초)을 초과했습니다.',
      INVALID_LOCATION: '지역 이름을 확인해 주세요.',
      UNAUTHORIZED: 'OpenWeather 키 인증에 실패했습니다.',
      RATE_LIMIT: 'OpenWeather 호출 한도를 초과했습니다.',
      PROVIDER_ERROR: 'OpenWeather 응답을 사용할 수 없습니다.',
      INVALID_RESPONSE: 'OpenWeather 현재값 또는 5일 예보가 불완전합니다.',
      NETWORK: 'OpenWeather 네트워크 연결에 실패했습니다.' })[code] || '날씨 조회에 실패했습니다.';
  }
  async function getWeather(query = 'Seoul,KR') {
    const requestedAt = new Date(now()).toISOString();
    const q = String(query).trim();
    const fail = (code) => ({ ...sample(safeError(code), requestedAt), reasonCode: code });
    if (!q || q.length > 100 || /[\x00-\x1f]/.test(q)) return fail('INVALID_LOCATION');
    const key = process.env.OPENWEATHER_API_KEY?.trim();
    if (!key) return fail('MISSING_KEY');
    const cacheKey = q.toLocaleLowerCase();
    const hit = cache.get(cacheKey);
    if (hit && now() - hit.createdAt < CACHE_MS) return { ...hit.data, cached: true, requestedAt };
    if (pending.has(cacheKey)) return pending.get(cacheKey);
    const work = (async () => {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), timeoutMs);
      try {
        async function request(kind) {
          // The credential is read only on the server and never logged or returned.
          const endpoint = new URL(`https://api.openweathermap.org/data/2.5/${kind}`);
          endpoint.search = new URLSearchParams({ q, appid: key, units: 'metric', lang: 'kr' }).toString();
          const response = await fetchImpl(endpoint, { signal: controller.signal, redirect: 'error' });
          if (!response.ok) throw new Error(response.status === 401 ? 'UNAUTHORIZED' : response.status === 404 ? 'INVALID_LOCATION' : response.status === 429 ? 'RATE_LIMIT' : 'PROVIDER_ERROR');
          return response.json();
        }
        const [current, forecast] = await Promise.all([request('weather'), request('forecast')]);
        if (!Number.isFinite(current.main?.temp) || !Number.isFinite(current.dt) || !Array.isArray(forecast.list) || !current.weather?.length) throw new Error('INVALID_RESPONSE');
        const timezoneOffset = forecast.city?.timezone ?? current.timezone ?? 0;
        const dateFor = (seconds) => new Date((seconds + timezoneOffset) * 1000).toISOString().slice(0, 10);
        const today = dateFor(current.dt);
        const days = new Map();
        for (const item of forecast.list) {
          if (!Number.isFinite(item.dt) || !Number.isFinite(item.main?.temp_min) || !Number.isFinite(item.main?.temp_max) || !item.weather?.length) throw new Error('INVALID_RESPONSE');
          const date = dateFor(item.dt);
          if (date <= today) continue;
          const day = days.get(date) || { date, min: Infinity, max: -Infinity, pop: 0, conditionId: item.weather[0].id, description: item.weather[0].description, slots: 0 };
          day.min = Math.min(day.min, item.main.temp_min);
          day.max = Math.max(day.max, item.main.temp_max);
          day.pop = Math.max(day.pop, item.pop || 0);
          day.slots++;
          days.set(date, day);
        }
        const daily = [...days.values()].sort((a, b) => a.date.localeCompare(b.date)).slice(0, 5);
        if (daily.length !== 5) throw new Error('INVALID_RESPONSE');
        const data = { mode: 'real', connection: 'live', cached: false, fetchedAt: new Date(now()).toISOString(), observedAt: new Date(current.dt * 1000).toISOString(),
          location: { name: `${current.name || q}${current.sys?.country ? ' · ' + current.sys.country : ''}`, timezoneOffset, note: '검색 지역의 실제 날씨이며 예제A점 위치를 의미하지 않습니다.' },
          source: { name: 'OpenWeather 현재 날씨 · 5일/3시간 예보', url: 'https://openweathermap.org/', forecastNote: '현지 날짜별 3시간 예보 구간의 최저·최고·최대 강수확률 (마지막 날은 일부 시간대일 수 있음)' },
          current: { temp: current.main.temp, feelsLike: current.main.feels_like, humidity: current.main.humidity, wind: current.wind?.speed, description: current.weather[0].description, conditionId: current.weather[0].id }, forecast: daily };
        if (cache.size >= 100) cache.delete(cache.keys().next().value);
        cache.set(cacheKey, { createdAt: now(), data });
        return data;
      } catch (error) {
        const result = fail(controller.signal.aborted ? 'TIMEOUT' : Object.keys({UNAUTHORIZED:1,INVALID_LOCATION:1,RATE_LIMIT:1,PROVIDER_ERROR:1,INVALID_RESPONSE:1}).includes(error.message) ? error.message : 'NETWORK');
        controller.abort();
        return result;
      } finally { clearTimeout(timer); }
    })();
    pending.set(cacheKey, work);
    try { return await work; } finally { pending.delete(cacheKey); }
  }
  return { getWeather };
}
const service = createWeatherService();
async function handler(req, res) {
  if (req.method !== 'GET') { res.setHeader('Allow', 'GET'); res.statusCode = 405; res.end(); return; }
  try {
    const url = new URL(req.url, 'http://localhost');
    const data = await service.getWeather(url.searchParams.get('q') || 'Seoul,KR');
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Cache-Control', 'no-store');
    res.end(JSON.stringify(data));
  } catch { res.statusCode = 503; res.end(JSON.stringify({ error: '예제 날씨 파일을 읽지 못했습니다.' })); }
}
module.exports = handler;
module.exports.createWeatherService = createWeatherService;
module.exports.CACHE_MS = CACHE_MS;
