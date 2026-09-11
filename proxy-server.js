const express = require('express');
const cors = require('cors');
const axios = require('axios');
const XLSX = require('xlsx');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const { v4: uuidv4 } = require('uuid');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const app = express();
const PORT = process.env.PORT || 5000;
const SERVER_START_TIME = Date.now(); // ✨ 서버 프로세스가 실제로 언제 켜졌는지 — 재시작 여부 진단용

app.use(cors());
app.use(express.json());

// ============================================================
// ✨ 헬스체크 전용 엔드포인트 — UptimeRobot 등 외부 모니터링이 여기로 핑을 보내면 됨.
// 무거운 작업(OneDrive 조회 등) 전혀 없이 즉시 200 OK만 반환한다.
// 루트('/')에 별도 핸들러가 없으면 404가 나서 일부 모니터가 "다운"으로 오탐할 수 있으므로,
// 루트와 /healthz 둘 다 확실히 200을 반환하도록 만들어둔다.
// ============================================================
app.get('/', (req, res) => {
  res.status(200).send('OK');
});
app.get('/healthz', (req, res) => {
  res.status(200).json({ status: 'ok', uptimeSeconds: Math.round((Date.now() - SERVER_START_TIME) / 1000) });
});

// ============================================================
// 환경 설정
// ============================================================
const CONFIG = {
  excelFileName: process.env.EXCEL_FILE_NAME || '재고관리(개발중).xlsx',
  clientId: process.env.CLIENT_ID,
  clientSecret: process.env.CLIENT_SECRET,
  redirectUri: process.env.REDIRECT_URI || 'http://localhost:5000/callback',
  inventorySheet: '공통',                       // 실제 부품 재고 데이터의 유일한 원본 시트
  facilityListSheets: ['충전', '타정', '유틸리티', '제조'],  // 설비명(적용설비) 목록만 있는 시트들 — 카드 UI 생성용
  facilityLogSheetName: '설비이력',              // 설비별 이력 시트
  logSheetName: '사용내역종합',
  teamsWebhookUrl: process.env.TEAMS_WEBHOOK_URL  // Teams Incoming Webhook URL
};

// 환경변수 로딩 상태 로깅
console.log('📋 환경변수 설정 상태:');
console.log(`   Excel File: ${CONFIG.excelFileName ? '✅ 설정됨' : '❌ 미설정'}`);
console.log(`   Client ID: ${CONFIG.clientId ? '✅ 설정됨' : '❌ 미설정'}`);
console.log(`   Gemini Key: ${process.env.GEMINI_API_KEY ? '✅ 설정됨' : '❌ 미설정'}`);
console.log(`   Refresh Token: ${process.env.REFRESH_TOKEN ? '✅ 설정됨' : '❌ 미설정'}`);
console.log(`   Teams Webhook: ${CONFIG.teamsWebhookUrl ? '✅ 설정됨' : '❌ 미설정 (알림 비활성화)'}`);
console.log(`   Company Email Domain: ${process.env.COMPANY_EMAIL_DOMAIN ? '✅ ' + process.env.COMPANY_EMAIL_DOMAIN : '❌ 미설정 (모든 이메일 허용됨 — 주의)'}`);

const TOKEN_FILE = path.join(__dirname, 'onedrive_tokens.json');
const LOG_FILE = path.resolve(__dirname, 'inventory_logs.json');
const USERS_FILE = path.join(__dirname, 'users.json');        // ✨ 이메일 기반 사용자 정보 저장
const SESSION_FILE = path.join(__dirname, 'sessions.json');   // ✨ 로그인 세션 토큰 저장

let memoryLogs = [];
let memoryTokens = null;
let memoryUsers = null;     // { email: { name, createdAt } }
let memorySessions = null;  // { token: { email, name, createdAt } }
const otpStore = new Map(); // 이메일 인증코드 임시 저장 — 수명이 짧아(10분) 파일 저장 불필요

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
const model = genAI.getGenerativeModel({ model: 'gemini-2.5-flash' });

// ============================================================
// Token 관리
// ============================================================
function loadTokens() {
  if (memoryTokens) return memoryTokens;
  try {
    if (fs.existsSync(TOKEN_FILE)) {
      const data = fs.readFileSync(TOKEN_FILE, 'utf8');
      return JSON.parse(data);
    }
  } catch (error) {
    console.error('Token 파일 읽기 실패:', error.message);
  }
  return null;
}

function saveTokens(tokens) {
  memoryTokens = tokens;
  if (!process.env.RENDER) {
    try {
      fs.writeFileSync(TOKEN_FILE, JSON.stringify(tokens, null, 2));
      console.log('✅ Token 파일 저장 완료');
    } catch (error) {
      console.error('❌ Token 파일 저장 실패:', error.message);
    }
  }
}

// ============================================================
// ✨ 사용자 인증 — 회사 이메일 + 일회용 인증코드 방식
// 가입/로그인이 하나의 흐름으로 통합됨: 이메일 입력 → 인증코드 발송(Microsoft Graph 메일) →
// 코드 확인 → 세션 발급. 비밀번호를 저장하지 않으므로 분실/재설정 절차가 필요 없다.
// 지정된 회사 이메일 도메인으로 끝나는 주소만 가입을 허용한다.
// ============================================================
const ALLOWED_EMAIL_DOMAIN = (process.env.COMPANY_EMAIL_DOMAIN || '').toLowerCase().replace(/^@/, '');
const OTP_DURATION_MS = 10 * 60 * 1000;     // 인증코드 유효시간 10분
const OTP_MAX_ATTEMPTS = 5;                  // 코드 입력 최대 시도 횟수
const SESSION_DURATION_MS = 12 * 60 * 60 * 1000; // 세션 유효시간 12시간

function isAllowedCompanyEmail(email) {
  if (!ALLOWED_EMAIL_DOMAIN) return true; // 도메인 제한 미설정 시엔 모두 허용 (환경변수 누락 방지용 안전장치)
  return String(email || '').toLowerCase().endsWith('@' + ALLOWED_EMAIL_DOMAIN);
}

function loadUsers() {
  if (memoryUsers) return memoryUsers;
  try {
    if (fs.existsSync(USERS_FILE)) {
      memoryUsers = JSON.parse(fs.readFileSync(USERS_FILE, 'utf8'));
      return memoryUsers;
    }
  } catch (error) {
    console.error('사용자 파일 읽기 실패:', error.message);
  }
  memoryUsers = {};
  return memoryUsers;
}

function saveUsers(users) {
  memoryUsers = users;
  try {
    fs.writeFileSync(USERS_FILE, JSON.stringify(users, null, 2));
  } catch (error) {
    console.error('❌ 사용자 파일 저장 실패:', error.message);
  }
}

// 6자리 숫자 인증코드 생성
function generateOtpCode() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

// ✨ Microsoft Graph API로 인증코드 메일 발송. OneDrive 연동에 쓰는 것과 동일한 앱 등록의
//    access token을 재사용한다 (Mail.Send 위임 권한이 앱에 추가되어 있어야 함).
async function sendVerificationEmail(toEmail, code) {
  const accessToken = await getValidAccessToken();
  const mailPayload = {
    message: {
      subject: '[스페어파트 재고관리] 로그인 인증코드',
      body: {
        contentType: 'HTML',
        content: `
          <div style="font-family: sans-serif; padding: 16px;">
            <p>안녕하세요,</p>
            <p>스페어파트 재고관리 시스템 로그인 인증코드입니다.</p>
            <p style="font-size: 28px; font-weight: 700; letter-spacing: 4px; color: #2563eb;">${code}</p>
            <p style="color: #6b7280; font-size: 13px;">이 코드는 10분간 유효합니다. 본인이 요청하지 않았다면 이 메일을 무시하세요.</p>
          </div>
        `
      },
      toRecipients: [{ emailAddress: { address: toEmail } }]
    },
    saveToSentItems: false
  };

  await axios.post(
    'https://graph.microsoft.com/v1.0/me/sendMail',
    mailPayload,
    { headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' } }
  );
}

// 세션 토큰 관리 (로그인 성공 시 발급, 이후 모든 요청에 Authorization 헤더로 실어야 함)
function loadSessions() {
  if (memorySessions) return memorySessions;
  try {
    if (fs.existsSync(SESSION_FILE)) {
      memorySessions = JSON.parse(fs.readFileSync(SESSION_FILE, 'utf8'));
      return memorySessions;
    }
  } catch (error) {
    console.error('세션 파일 읽기 실패:', error.message);
  }
  memorySessions = {};
  return memorySessions;
}

function saveSessions(sessions) {
  memorySessions = sessions;
  try {
    fs.writeFileSync(SESSION_FILE, JSON.stringify(sessions, null, 2));
  } catch (error) {
    console.error('❌ 세션 파일 저장 실패:', error.message);
  }
}

function createSession(email, name) {
  const sessions = loadSessions();
  const token = crypto.randomBytes(32).toString('hex');
  sessions[token] = { email, name, createdAt: Date.now() };
  saveSessions(sessions);
  return token;
}

function getSession(token) {
  if (!token) return null;
  const sessions = loadSessions();
  const session = sessions[token];
  if (!session) return null;
  if (Date.now() - session.createdAt > SESSION_DURATION_MS) {
    delete sessions[token];
    saveSessions(sessions);
    return null;
  }
  return session;
}

function destroySession(token) {
  const sessions = loadSessions();
  if (sessions[token]) {
    delete sessions[token];
    saveSessions(sessions);
  }
}

// ============================================================
// ✨ 인증 미들웨어 — 재고 변경(추가/삭제/수정) 등 민감한 API는 반드시 이걸 통과해야 함
// ============================================================
function requireAuth(req, res, next) {
  const authHeader = req.headers['authorization'] || '';
  const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : null;
  const session = getSession(token);
  if (!session) {
    return res.status(401).json({ success: false, message: '로그인이 필요합니다. 다시 로그인해 주세요.' });
  }
  req.authUser = session; // { email, name }
  next();
}

async function refreshAccessToken(refreshToken, clientIdOverride) {
  const clientId = clientIdOverride || CONFIG.clientId;
  try {
    console.log('🔄 Access Token 갱신 중...');
    const params = {
      client_id: clientId,
      refresh_token: refreshToken,
      grant_type: 'refresh_token'
    };
    // client_secret이 있으면 포함
    if (CONFIG.clientSecret) {
      params.client_secret = CONFIG.clientSecret;
    }
    const response = await axios.post(
      'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      new URLSearchParams(params),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    const tokens = {
      access_token: response.data.access_token,
      refresh_token: response.data.refresh_token || refreshToken,
      expires_at: Date.now() + (response.data.expires_in * 1000)
    };
    saveTokens(tokens);
    console.log('✅ Access Token 갱신 성공!');
    return tokens;
  } catch (error) {
    console.error('❌ Token 갱신 실패:', error.response?.data || error.message);
    return null;
  }
}

async function getValidAccessToken() {
  // ✨ 이전에는 REFRESH_TOKEN 환경변수가 있으면 이미 발급받은 access_token이
  //    아직 유효하더라도 매번 MS 서버에 새로 갱신 요청을 보냈다 (모든 요청마다
  //    네트워크 왕복 1회씩 추가되어 응답이 느려지는 원인 중 하나).
  //    이제는 메모리에 유효한(만료 1분 전 이내) 토큰이 있으면 그걸 그대로 재사용하고,
  //    없거나 만료 임박(또는 서버 재배포로 메모리가 초기화된 경우)일 때만 새로 갱신한다.
  //    → 재배포 직후 안전하게 새 토큰을 받아오는 기존 동작은 그대로 유지된다.
  const cachedTokens = memoryTokens || loadTokens();
  if (cachedTokens && cachedTokens.access_token && Date.now() < cachedTokens.expires_at - 60000) {
    return cachedTokens.access_token;
  }

  // 1. 환경변수 REFRESH_TOKEN 최우선 사용 (메모리 토큰이 없거나 만료 임박한 경우에만 도달)
  if (process.env.REFRESH_TOKEN) {
    try {
      console.log('🔑 환경변수 REFRESH_TOKEN으로 갱신 중...');
      const params = {
        client_id: CONFIG.clientId,
        refresh_token: process.env.REFRESH_TOKEN,
        grant_type: 'refresh_token'
      };
      if (CONFIG.clientSecret) {
        params.client_secret = CONFIG.clientSecret;
      }
      const response = await axios.post(
        'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        new URLSearchParams(params),
        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
      );
      const newTokens = {
        access_token: response.data.access_token,
        refresh_token: response.data.refresh_token || process.env.REFRESH_TOKEN,
        expires_at: Date.now() + (response.data.expires_in * 1000)
      };
      saveTokens(newTokens);
      console.log('✅ 환경변수 REFRESH_TOKEN으로 갱신 성공!');
      return newTokens.access_token;
    } catch (err) {
      console.error('❌ 환경변수 토큰 갱신 실패:', err.response?.data || err.message);
      console.log('📁 로컬 저장 토큰으로 전환합니다...');
    }
  }

  // 2. 저장된 토큰 로드
  let tokens = loadTokens();

  // 3. 토큰 없으면 Device Flow
  if (!tokens) {
    console.log('⚠️ 저장된 토큰이 없습니다. Device Flow를 시작합니다.');
    tokens = await getTokenViaDeviceFlow();
    if (!tokens) throw new Error('인증에 실패했습니다.');
    return tokens.access_token;
  }

  // 4. 만료 시 갱신
  if (Date.now() >= tokens.expires_at - 60000) {
    console.log('🔄 토큰 만료됨. 갱신 중...');
    const refreshed = await refreshAccessToken(tokens.refresh_token);
    if (!refreshed) {
      tokens = await getTokenViaDeviceFlow();
      if (!tokens) throw new Error('재인증 실패');
      return tokens.access_token;
    }
    return refreshed.access_token;
  }

  return tokens.access_token;
}

// ============================================================
// 로그 관리
// ============================================================
function loadLogs() {
  if (memoryLogs && memoryLogs.length > 0) return memoryLogs;
  try {
    if (fs.existsSync(LOG_FILE)) {
      const data = fs.readFileSync(LOG_FILE, 'utf8');
      return JSON.parse(data);
    }
  } catch (error) {
    console.error('❌ 로그 읽기 실패:', error.message);
  }
  return [];
}

function saveLogs(logs) {
  memoryLogs = logs;
  try {
    fs.writeFileSync(LOG_FILE, JSON.stringify(logs, null, 2));
  } catch (error) {
    console.error('❌ 로그 저장 실패:', error.message);
  }
}

function addLog(action, item, quantityChange, user = 'System', sharedId = null) {
  const newLog = {
    id: sharedId || uuidv4(),
    timestampKR: getKSTDate(),
    action,
    원본시트: item.원본시트 || '미분류',
    부품종류: item.부품종류,
    모델명: item.모델명,
    적용설비: item.적용설비,
    변경수량: quantityChange,
    변경전수량: item.현재수량 - quantityChange,
    변경후수량: item.현재수량,
    user
  };
  const logs = loadLogs();
  logs.unshift(newLog);
  if (logs.length > 1000) logs.splice(1000);
  saveLogs(logs);
  console.log(`📝 로그: ${action} - ${item.모델명} (${quantityChange > 0 ? '+' : ''}${quantityChange})`);
}

// ============================================================
// OneDrive 엑셀 읽기 (Graph API + OAuth 토큰 방식)
// ============================================================
let cachedData = null;
let lastFetchTime = null;
// ✨ 정상적으로 로드됐던 데이터 개수 중 최댓값을 기억해둔다. updateExcelOnOneDrive()가
//    이 값보다 터무니없이 적은 데이터로 원본 전체를 덮어쓰려 하면 안전하게 막기 위한 기준선.
let lastKnownGoodDataCount = 0;
// ✨ 캐시 유효시간을 60초 → 5분으로 늘림. 실제 재고 변경 시에는 invalidateCache()가
//    즉시 호출되어 캐시가 바로 무효화되므로, 시간을 늘려도 "수정 직후 옛날 값이 보이는"
//    문제는 생기지 않는다. 대신 앱을 열 때마다 캐시가 만료돼 있어 매번 OneDrive 전체를
//    재다운로드하던 상황(3초 vs 30초 편차의 주원인)이 크게 줄어든다.
const CACHE_DURATION = 5 * 60 * 1000;

// 설비이력 메모리 버퍼
let facilityLogs = [];
let lastFacilityLogRowCount = null; // ✨ 시트 행 수 변화 감지용 — 변화 없으면 재파싱 생략

function invalidateCache() {
  cachedData = null;
  lastFetchTime = null;
}

// ============================================================
// 설비명 정리 (공백/줄바꿈만 정리 — 별도 매핑 테이블 사용 안 함)
// ============================================================
// ⚠️ 과거에는 '제조' 시트("원본설비명→표준설비명" 매핑 테이블)를 참조했으나,
//    실제로는 관리되지 않는 사문화된 시트였고, 매핑 실수로 호기 표기(#3 등)가
//    지워지는 등 부작용만 있어 완전히 제거했다.
//    이제 적용설비 원본 문자열의 공백/줄바꿈만 정리해서 그대로 표준설비명으로 사용한다.
// ============================================================
// 전각(全角) 특수문자 → 반각 정규화
// ⚠️ 엑셀에서 한글 입력기를 쓰다 실수로 전각 샵(＃, U+FF03) 등이 섞여 들어가면
//    설비명 매칭(설비 목록 대조)이 실패할 수 있다. 모든 판별 로직에 들어가기 전에
//    이 정규화를 거쳐서, 어떤 문자가 섞여 들어와도 안전하게 반각으로 통일한다.
// ============================================================
function normalizeSpecialChars(str) {
  return String(str || '')
    .replace(/＃/g, '#')   // 전각 샵 → 반각 샵
    .replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)) // 전각 숫자 → 반각 숫자
    .replace(/（/g, '(')
    .replace(/）/g, ')');
}

function normalizeEquipment(originalName) {
  return normalizeSpecialChars(String(originalName || '').replace(/[\r\n]+/g, ' ')).replace(/\s+/g, ' ').trim();
}

// ============================================================
// 설비 목록(카드 UI 생성용) 로드
// ============================================================
// '충전'/'타정' 시트는 이제 부품 데이터를 전혀 담지 않고, 헤더 '적용설비' 하나만 있는
// 단순 설비명 목록이다. 실제 부품 재고는 오직 '공통' 시트에만 존재하며,
// 모든 부품은 출고 시 "어느 설비에 사용했는지" 확인 절차를 거친다 (기존 공통부품 흐름과 동일).
// 이 목록은 카드 UI 구성과, 출고 시 실제사용설비 값 검증(오타 방지) 용도로만 쓰인다.
// ============================================================
let facilityListCache = null; // { 충전: [...], 타정: [...], all: [...] }

async function loadFacilityLists(workbook) {
  const lists = {};
  let all = [];

  CONFIG.facilityListSheets.forEach(sheetName => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      console.warn(`⚠️ 설비 목록 시트 "${sheetName}"을 찾을 수 없습니다.`);
      lists[sheetName] = [];
      return;
    }
    const rows = XLSX.utils.sheet_to_json(sheet);
    const names = rows
      .map(r => normalizeEquipment(r['적용설비']))
      .filter(Boolean);
    lists[sheetName] = names;
    all = all.concat(names);
    console.log(`🏭 "${sheetName}" 설비 목록 로드: ${names.length}개`);
  });

  facilityListCache = { ...lists, all: [...new Set(all)] };
  return facilityListCache;
}

// ============================================================
// 설비이력 관리
// ============================================================
function addFacilityLog(action, item, quantityChange, user, sharedId = null) {
  const stdEquipment = item.표준설비명 || item.적용설비;
  const entry = {
    id: sharedId || uuidv4(),
    timestampKR: getKSTDate(),
    action,
    원본시트: item.원본시트 || '',
    표준설비명: stdEquipment,
    원본설비명: item.적용설비,
    부품종류: item.부품종류,
    모델명: item.모델명,
    변경수량: quantityChange,
    변경전수량: item.현재수량 - quantityChange,
    변경후수량: item.현재수량,
    isCommonPart: item.isCommonPart || false,
    user
  };
  facilityLogs.unshift(entry);
  if (facilityLogs.length > 5000) facilityLogs.splice(5000);
  console.log(`🏭 설비이력: [${stdEquipment}] ${action} - ${item.모델명} (${quantityChange > 0 ? '+' : ''}${quantityChange})`);
}

async function saveFacilityLogsToOneDrive(workbook) {
  // 설비이력 시트에 현재까지의 facilityLogs를 저장 (updateExcelOnOneDrive 내부에서 호출)
  if (facilityLogs.length === 0) return workbook;
  const rows = [...facilityLogs].reverse(); // 오래된 순서로 저장
  const ws = XLSX.utils.json_to_sheet(rows);
  if (workbook.Sheets[CONFIG.facilityLogSheetName]) {
    workbook.Sheets[CONFIG.facilityLogSheetName] = ws;
  } else {
    XLSX.utils.book_append_sheet(workbook, ws, CONFIG.facilityLogSheetName);
  }
  return workbook;
}

async function fetchExcelFromOneDrive() {
  const now = Date.now();
  if (cachedData && lastFetchTime && (now - lastFetchTime) < CACHE_DURATION) {
    console.log(`📦 캐시된 통합 데이터 사용 (캐시 나이: ${Math.round((now - lastFetchTime) / 1000)}초)`);
    return cachedData;
  }

  const fetchStartTime = Date.now();
  console.log('🐢 캐시 미스 — OneDrive 재다운로드 시작');

  try {
    const accessToken = await getValidAccessToken();
    console.log(`📥 OneDrive에서 "${CONFIG.excelFileName}" 다운로드 중...`);

    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/me/drive/root:/${CONFIG.excelFileName}:/content`,
      {
        headers: { 'Authorization': `Bearer ${accessToken}` },
        responseType: 'arraybuffer'
      }
    );

    const workbook = XLSX.read(Buffer.from(response.data), { type: 'buffer' });
    let allMappedData = [];

    // 설비 목록(카드 UI용) 먼저 로드 — 충전/타정 시트는 이제 적용설비 목록만 담고 있음
    const facilityLists = await loadFacilityLists(workbook);

    // 공통 시트 하나만 순회 — 실제 부품 재고의 유일한 원본
    const worksheet = workbook.Sheets[CONFIG.inventorySheet];
    if (!worksheet) {
      console.warn(`⚠️ 시트 "${CONFIG.inventorySheet}"을 찾을 수 없습니다`);
    } else {
      const jsonData = XLSX.utils.sheet_to_json(worksheet);
      console.log(`✅ "${CONFIG.inventorySheet}" 시트: ${jsonData.length}개 항목`);

      allMappedData = jsonData.map((row, index) => {
        const rowKeys = Object.keys(row);
        const foundKey = rowKeys.find(key => key.trim() === '보관장소');
        const rawEquip = row['적용설비'] || '';
        const stdEquip = normalizeEquipment(rawEquip);

        return {
          id: `${CONFIG.inventorySheet}_${index + 1}`,
          원본시트: CONFIG.inventorySheet,       // 이제 모든 부품이 '공통' 소속
          대분류: row['대분류'] || '미분류',
          부품종류: row['부품종류'] || '',
          모델명: row['모델명'] || '',
          적용설비: row['적용설비'] || '',        // 엑셀 원본 그대로 (참고/필터용)
          표준설비명: stdEquip,
          isCommonPart: true,                    // 모든 부품이 실사용 설비 확인 절차를 거침
          // ✨ 후보설비목록은 모든 부품이 동일한 전체 설비 목록을 참조하므로,
          //    항목마다(1592개) 중복 포함하지 않는다. 프론트는 전역 facilityLists.all을 대신 사용한다.
          //    (이전에는 부품마다 전체 설비 배열을 통째로 복제해 JSON 응답 크기가 크게 부풀어 있었음)
          현재수량: Number(row['현재수량']) || 0,
          최소보유수량: Number(row['최소보유수량']) || 0,
          최종수정시각: row['최종수정시각'] || '',
          작업자: row['작업자'] || '',
          용도: row['용도'] || '',
          보관장소: foundKey ? row[foundKey] : '위치 미지정'
        };
      });
    }

    // 설비 목록이 비어있으면 확인 절차 자체가 불가능하므로 경고
    if ((facilityListCache?.all || []).length === 0) {
      console.warn(`⚠️ 설비 목록(충전/타정 시트)이 비어있습니다 — 부품 사용 시 설비 선택지가 제공되지 않습니다.`);
    }

    // 로그 시트 로드 (사용내역종합) — memoryLogs가 이미 있으면 덮어쓰지 않음
    const logWorksheet = workbook.Sheets[CONFIG.logSheetName];
    if (logWorksheet && memoryLogs.length === 0) {
      const logJson = XLSX.utils.sheet_to_json(logWorksheet);
      // 오래된 순 저장 → 최신순으로 reverse, 상한 없이 전체 보관
      memoryLogs = logJson.reverse();
      console.log(`📜 로그 시트 로드 완료: ${memoryLogs.length}건`);
    }

    // 설비이력 시트 로드 — 엑셀 이력과 메모리 이력을 병합하되, 매번 전체를 다시 파싱하지 않는다.
    // ✨ 이전에는 캐시가 만료될 때마다(60초마다 한 번씩) 이 시트를 통째로 sheet_to_json()으로
    //    다시 변환했는데, 이력이 쌓일수록(최대 5000건) 이 비용이 계속 커져 응답이 느려지는
    //    원인 중 하나였다. 시트의 실제 행 개수가 이전과 같으면(=엑셀에서 직접 수정되지 않았으면)
    //    파싱을 건너뛰고 메모리 상태를 그대로 사용한다.
    const facilityLogWs = workbook.Sheets[CONFIG.facilityLogSheetName];
    if (facilityLogWs) {
      const sheetRange = XLSX.utils.decode_range(facilityLogWs['!ref'] || 'A1:A1');
      const currentRowCount = sheetRange.e.r; // 헤더 제외 데이터 행 수(0-index 기준 근사치)

      if (currentRowCount !== lastFacilityLogRowCount) {
        const rows = XLSX.utils.sheet_to_json(facilityLogWs);
        const excelLogs = rows.reverse(); // 최신순
        const memoryIds = new Set(facilityLogs.map(l => l.id));
        const newFromExcel = excelLogs.filter(l => l.id && !memoryIds.has(l.id));
        if (newFromExcel.length > 0) {
          facilityLogs = [...facilityLogs, ...newFromExcel];
          facilityLogs.sort((a, b) => new Date(b.timestampKR || 0) - new Date(a.timestampKR || 0));
          facilityLogs = facilityLogs.slice(0, 5000);
          console.log(`🏭 설비이력 병합 완료: 총 ${facilityLogs.length}건 (엑셀에서 ${newFromExcel.length}건 추가)`);
        } else if (facilityLogs.length === 0) {
          facilityLogs = excelLogs.slice(0, 5000);
          console.log(`🏭 설비이력 초기 로드: ${facilityLogs.length}건`);
        }
        lastFacilityLogRowCount = currentRowCount;
      } else {
        console.log('🏭 설비이력 시트 변경 없음 — 재파싱 생략');
      }
    }

    cachedData = allMappedData;
    lastFetchTime = now;
    // ✨ 정상적으로 부품 데이터가 로드됐을 때만(0건이 아닐 때만) 기준선을 갱신.
    //    실제로 재고를 전부 비우는 정당한 상황은 없다고 가정 — 있다면 이 기준선 자체를
    //    수동으로 초기화해야 함을 인지할 수 있도록 로그를 남긴다.
    if (allMappedData.length > 0) {
      lastKnownGoodDataCount = Math.max(lastKnownGoodDataCount, allMappedData.length);
    }
    console.log(`✅ 데이터 로드 완료: 총 ${allMappedData.length}건 (${Date.now() - fetchStartTime}ms 소요)`);
    return allMappedData;

  } catch (error) {
    console.error('❌ OneDrive 읽기 실패:', error.response?.data || error.message);
    // ✨ 매우 중요: 절대로 빈 배열을 반환하면 안 된다.
    //    과거에 여기서 조용히 []를 반환했던 것이, 호출부(add-part, AI챗봇 재고수정 등)가
    //    이를 "실제로 재고가 텅 빈 상태"로 착각하게 만들어 updateExcelOnOneDrive()가
    //    OneDrive 원본 파일 전체를 거의 빈 데이터로 덮어써버리는 심각한 데이터 유실 사고로
    //    이어진 적이 있다. 실패는 반드시 예외로 전파해서, 호출부가 명확히 실패를 인지하고
    //    쓰기 작업을 절대 진행하지 않도록 한다.
    throw error;
  }
}

async function updateExcelOnOneDrive(data, retries = 3) {
  // ✨ 최종 안전장치: 이 함수는 워크북을 처음부터 다시 만들어 OneDrive 원본 파일 전체를
  //    통째로 덮어쓰는 매우 위험한 작업이다. 만약 상위 호출부에서 데이터를 잘못 불러왔거나
  //    (예: fetchExcelFromOneDrive 실패를 놓친 경우) 비정상적으로 줄어든 배열을 넘기면,
  //    실제 재고 데이터가 통째로 날아가는 사고로 이어진다(과거 실제 발생 이력 있음).
  //    정상적으로 로드됐던 최대 건수 대비 절반 미만으로 줄어들면 명백히 비정상으로 보고
  //    쓰기 자체를 거부한다. (부품을 대량으로 일부러 삭제하는 정상적인 작업이라면
  //    이 기준선이 걸릴 수 있으니, 그런 작업 이후엔 서버 재시작 등으로 기준선이 자연히 갱신됨)
  if (lastKnownGoodDataCount > 0 && Array.isArray(data) && data.length < lastKnownGoodDataCount * 0.5) {
    console.error(`🚨 위험 감지: 저장하려는 데이터(${data.length}건)가 기존 정상 데이터(최대 ${lastKnownGoodDataCount}건)의 절반 미만입니다. 데이터 유실 방지를 위해 저장을 거부합니다.`);
    return false;
  }

  // ✨ 엑셀 셀 하나에 32,767자를 넘는 텍스트가 들어가면 xlsx 저장 자체가 실패한다.
  //    (실제 사례: '_x000D_' 같은 깨진 캐리지리턴 이스케이프 문자열이 반복 삽입되어
  //     특정 셀이 수만~10만자 이상으로 부풀어 있던 경우가 있었음)
  //    저장 직전에 모든 셀 값을 검사해서, 한도를 넘으면 안전하게 잘라내고 서버 로그에
  //    남긴다 — 데이터 유실 위험이 있으므로 절단된 셀은 반드시 로그로 알 수 있게 한다.
  const EXCEL_CELL_MAX_LEN = 32767;

  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const accessToken = await getValidAccessToken();
      const workbook = XLSX.utils.book_new();
      const truncatedCells = [];
      const sanitizeCellValue = (value, rowLabel, fieldName) => {
        if (typeof value !== 'string') return value;
        if (value.length <= EXCEL_CELL_MAX_LEN) return value;
        truncatedCells.push({ row: rowLabel, field: fieldName, originalLength: value.length });
        return value.slice(0, EXCEL_CELL_MAX_LEN);
      };

      // 공통 시트 저장 — 이제 부품 재고의 유일한 원본
      const excelRows = data.map(item => {
        const rowLabel = item.모델명 || item.id || '(모델명 없음)';
        return {
          '대분류': sanitizeCellValue(item.대분류 || '미분류', rowLabel, '대분류'),
          '부품종류': sanitizeCellValue(item.부품종류 || '', rowLabel, '부품종류'),
          '모델명': sanitizeCellValue(item.모델명 || '', rowLabel, '모델명'),
          '적용설비': sanitizeCellValue(item.적용설비 || '', rowLabel, '적용설비'),
          '현재수량': Number(item.현재수량) || 0,
          '최소보유수량': Number(item.최소보유수량) || 0,
          '최종수정시각': item.최종수정시각 || '',
          '작업자': sanitizeCellValue(item.작업자 || '', rowLabel, '작업자'),
          '용도': sanitizeCellValue(item.용도 || '', rowLabel, '용도'),
          '보관장소': sanitizeCellValue(item.보관장소 || '위치 미지정', rowLabel, '보관장소')
        };
      });

      if (truncatedCells.length > 0) {
        console.warn(`⚠️ 셀 길이 초과로 잘라낸 항목 ${truncatedCells.length}건:`);
        truncatedCells.forEach(c => {
          console.warn(`   - [${c.row}] ${c.field} 필드: ${c.originalLength}자 → ${EXCEL_CELL_MAX_LEN}자로 절단`);
        });
      }

      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(excelRows), CONFIG.inventorySheet);

      // 충전/타정 설비명 목록 시트 복원 — 앱에서 직접 수정하지 않는 참조 목록이므로
      // 로드 시점에 캐시해둔 목록을 그대로 다시 써서 유실 방지 (헤더는 '적용설비' 단일열)
      CONFIG.facilityListSheets.forEach(sheetName => {
        const names = facilityListCache?.[sheetName] || [];
        const rows = names.map(name => ({ '적용설비': name }));
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), sheetName);
      });

      // 로그 시트 저장
      const logRows = [...memoryLogs].reverse();
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(logRows), CONFIG.logSheetName);

      // 설비이력 시트 저장
      if (facilityLogs.length > 0) {
        const facilityRows = [...facilityLogs].reverse();
        const facilityWs = XLSX.utils.json_to_sheet(facilityRows);
        XLSX.utils.book_append_sheet(workbook, facilityWs, CONFIG.facilityLogSheetName);
      }

      const excelBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

      await axios.put(
        `https://graph.microsoft.com/v1.0/me/drive/root:/${CONFIG.excelFileName}:/content`,
        excelBuffer,
        {
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          }
        }
      );

      console.log(`✅ OneDrive 업데이트 완료! (${CONFIG.excelFileName})`);
      invalidateCache();
      return true;

    } catch (error) {
      console.error(`❌ OneDrive 쓰기 실패 (${attempt}/${retries}): ${error.message}`);
      if (attempt < retries) {
        await new Promise(resolve => setTimeout(resolve, attempt * 2000));
        continue;
      }
      return false;
    }
  }
  return false;
}

const getKSTDate = () => {
  const curr = new Date();
  const utc = curr.getTime() + (curr.getTimezoneOffset() * 60 * 1000);
  const KR_TIME_DIFF = 9 * 60 * 60 * 1000;
  const kstDate = new Date(utc + KR_TIME_DIFF);
  return kstDate.toLocaleString('ko-KR');
};

// ============================================================
// Teams 재고 부족 알림
// ============================================================

// 중복 알림 방지 — 동일 항목은 1시간에 1번만 알림
const alertCooldown = new Map();
const ALERT_COOLDOWN_MS = 60 * 60 * 1000;

async function sendTeamsAlert(lowStockItems) {
  if (!CONFIG.teamsWebhookUrl) {
    console.log('⚠️ TEAMS_WEBHOOK_URL 미설정 — 알림 스킵');
    return;
  }

  // 최소보유수량 > 0 이고 쿨다운 지난 항목만 필터
  const now = Date.now();
  const filtered = lowStockItems.filter(item => {
    if (item.최소보유수량 <= 0) return false;
    const lastAlerted = alertCooldown.get(item.id) || 0;
    return now - lastAlerted >= ALERT_COOLDOWN_MS;
  });

  if (filtered.length === 0) {
    console.log('ℹ️ Teams 알림 대상 없음 (쿨다운 또는 조건 미충족)');
    return;
  }

  // ✨ 카드 페이로드 크기 제한: 항목 수가 너무 많으면(예: 초기 도입 시 수백~천 건대) 카드가
  //    Teams/Power Automate 웹훅의 페이로드 크기 제한을 초과해 전송 자체가 실패할 수 있다.
  //    긴급도(재고 0 → 부족량 큰 순)로 정렬해 상위 N건만 카드에 담고, 나머지는 건수만 요약해서 알려준다.
  const MAX_ITEMS_IN_CARD = 30;
  const sortedFiltered = [...filtered].sort((a, b) => {
    if (a.현재수량 === 0 && b.현재수량 !== 0) return -1;
    if (a.현재수량 !== 0 && b.현재수량 === 0) return 1;
    const shortageA = a.최소보유수량 - a.현재수량;
    const shortageB = b.최소보유수량 - b.현재수량;
    return shortageB - shortageA;
  });
  const toSend = sortedFiltered.slice(0, MAX_ITEMS_IN_CARD);
  const remainingCount = sortedFiltered.length - toSend.length;

  const critical = filtered.filter(i => i.현재수량 === 0);
  const warning  = filtered.filter(i => i.현재수량 > 0);

  const titleText = critical.length > 0 ? '🚨 긴급 재고 부족 알림' : '⚠️ 재고 부족 알림';
  const summaryParts = [];
  if (critical.length > 0) summaryParts.push(`🔴 재고 0: **${critical.length}건**`);
  if (warning.length  > 0) summaryParts.push(`🟡 부족 경고: **${warning.length}건**`);

  // 품목별 카드형 블록 생성 (설비명이 길어도 깔끔하게 표시) — 상위 MAX_ITEMS_IN_CARD건만
  const itemBlocks = toSend.map(item => {
    const isCritical = item.현재수량 === 0;
    const badge = isCritical ? '🔴 재고 없음' : '🟡 부족 경고';
    const shortage = item.최소보유수량 - item.현재수량;
    const facilityName = String(item.표준설비명 || item.적용설비 || '-').replace(/[\r\n]+/g, ' ').trim();
    return {
      type: 'Container',
      style: isCritical ? 'attention' : 'warning',
      spacing: 'Small',
      items: [
        {
          type: 'ColumnSet',
          columns: [
            {
              type: 'Column',
              width: 'stretch',
              items: [
                {
                  type: 'TextBlock',
                  text: `**${item.모델명 || '-'}**  ${badge}`,
                  wrap: true,
                  size: 'Default',
                  weight: 'Bolder',
                  color: isCritical ? 'Attention' : 'Warning'
                },
                {
                  type: 'TextBlock',
                  text: `📦 ${item.부품종류 || '-'}　|　🏭 ${facilityName}　|　📋 ${item.원본시트 || '-'}시트`,
                  wrap: true,
                  size: 'Small',
                  isSubtle: true,
                  spacing: 'None'
                }
              ]
            },
            {
              type: 'Column',
              width: 'auto',
              items: [
                {
                  type: 'TextBlock',
                  text: `현재 **${item.현재수량}**개`,
                  wrap: false,
                  size: 'Small',
                  color: isCritical ? 'Attention' : 'Warning',
                  weight: 'Bolder',
                  horizontalAlignment: 'Right'
                },
                {
                  type: 'TextBlock',
                  text: `최소 ${item.최소보유수량}개 / **${shortage}개 부족**`,
                  wrap: false,
                  size: 'Small',
                  isSubtle: true,
                  horizontalAlignment: 'Right',
                  spacing: 'None'
                }
              ]
            }
          ]
        }
      ]
    };
  });

  const card = {
    type: 'message',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: {
        type: 'AdaptiveCard',
        version: '1.4',
        body: [
          {
            type: 'Container',
            style: 'emphasis',
            items: [
              { type: 'TextBlock', text: titleText, weight: 'Bolder', size: 'Large', color: 'Attention', wrap: true },
              { type: 'TextBlock', text: `🕐 ${getKSTDate()}　　${summaryParts.join('　|　')}`, size: 'Small', isSubtle: true, spacing: 'None', wrap: true }
            ]
          },
          { type: 'TextBlock', text: '─────────────────────', size: 'Small', isSubtle: true, spacing: 'Small' },
          ...itemBlocks,
          ...(remainingCount > 0
            ? [{ type: 'TextBlock', text: `➕ 그 외 ${remainingCount}건 더 있음 (앱에서 전체 확인)`, size: 'Small', weight: 'Bolder', color: 'Attention', wrap: true, spacing: 'Medium' }]
            : []),
          { type: 'TextBlock', text: '※ 최소보유수량 0 설정 항목은 알림 제외', size: 'Small', isSubtle: true, wrap: true, spacing: 'Medium' }
        ]
      }
    }]
  };

  try {
    await axios.post(CONFIG.teamsWebhookUrl, card);
    // ✨ 쿨다운은 전송이 실제로 성공했을 때만 설정한다. (이전에는 시도 전에 미리 설정해서,
    //    전송이 실패해도 "이미 보냄" 상태가 되어 다음 1시간 동안 재시도가 조용히 막혔던 버그가 있었음)
    filtered.forEach(item => alertCooldown.set(item.id, now));
    console.log(`✅ Teams 알림 전송 완료 — 카드에 ${toSend.length}건 표시 (전체 대상 ${filtered.length}건, 긴급 ${critical.length}, 경고 ${warning.length})`);
  } catch (err) {
    console.error('❌ Teams 알림 전송 실패 — 쿨다운 설정 안 함(다음 체크 때 재시도됨):', err.response?.status, err.response?.data || err.message);
  }
}

// 재고 수정 후 저재고 체크 & 알림 트리거 (non-blocking)
function checkAndNotifyLowStock(data) {
  const lowStock = data.filter(d => d.최소보유수량 > 0 && d.현재수량 < d.최소보유수량);
  if (lowStock.length > 0) {
    console.log(`📊 저재고 감지: ${lowStock.length}건 → Teams 알림 시도`);
    sendTeamsAlert(lowStock).catch(err => console.error('Teams 알림 오류:', err.message));
  }
}

// ============================================================
// API Routes
// ============================================================
app.get('/api/inventory', async (req, res) => {
  try {
    const data = await fetchExcelFromOneDrive();
    // 설비 목록(충전/타정 시트 기반) — 프론트엔드가 카드 목록/설비 선택지를 구성할 때 사용
    res.json({ success: true, data, facilityLists: facilityListCache || { 충전: [], 타정: [], all: [] } });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.get('/api/inventory/categories', async (req, res) => {
  try {
    const data = await fetchExcelFromOneDrive();
    const categories = {};
    data.forEach(item => {
      const mainCat = item.대분류 || '미분류';
      if (!categories[mainCat]) {
        categories[mainCat] = { name: mainCat, totalCount: 0, itemCount: 0, lowStockCount: 0, items: [] };
      }
      categories[mainCat].items.push(item);
      categories[mainCat].totalCount += item.현재수량;
      categories[mainCat].itemCount += 1;
      if (item.최소보유수량 > 0 && item.현재수량 < item.최소보유수량) {
        categories[mainCat].lowStockCount += 1;
      }
    });
    res.json({ success: true, data: Object.values(categories) });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.get('/api/inventory/category/:categoryName', async (req, res) => {
  try {
    const data = await fetchExcelFromOneDrive();
    const filtered = data.filter(item => item.대분류 === req.params.categoryName);
    res.json({ success: true, data: filtered });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.get('/api/inventory/summary', async (req, res) => {
  try {
    const data = await fetchExcelFromOneDrive();
    const summary = {
      totalItems: data.length,
      totalQuantity: data.reduce((sum, d) => sum + d.현재수량, 0),
      lowStockItems: data.filter(d => d.최소보유수량 > 0 && d.현재수량 < d.최소보유수량),
      lowStockCount: data.filter(d => d.최소보유수량 > 0 && d.현재수량 < d.최소보유수량).length,
      categoryBreakdown: {}
    };
    data.forEach(item => {
      if (!summary.categoryBreakdown[item.부품종류]) {
        summary.categoryBreakdown[item.부품종류] = { total: 0, count: 0, lowStock: 0 };
      }
      summary.categoryBreakdown[item.부품종류].total += item.현재수량;
      summary.categoryBreakdown[item.부품종류].count += 1;
      if (item.현재수량 < item.최소보유수량) summary.categoryBreakdown[item.부품종류].lowStock += 1;
    });
    res.json({ success: true, data: summary });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.post('/api/inventory/update', requireAuth, async (req, res) => {
  try {
    const { id, 현재수량, action, user } = req.body;
    const data = await fetchExcelFromOneDrive();
    const item = data.find(d => d.id == id);
    if (!item) return res.status(404).json({ success: false, message: '항목을 찾을 수 없습니다.' });

    const oldQuantity = item.현재수량;
    item.현재수량 = 현재수량;
    item.최종수정시각 = getKSTDate();

    const success = await updateExcelOnOneDrive(data);
    if (success) {
      try {
        const sharedLogId = uuidv4(); // 두 로그 저장소(사용내역종합/설비이력)를 같은 id로 연결 — 되돌리기 시 함께 찾기 위함
        addLog(action || '수정', item, 현재수량 - oldQuantity, user || 'Manual', sharedLogId);
        addFacilityLog(action || '수정', item, 현재수량 - oldQuantity, user || 'Manual', sharedLogId);
      } catch (logErr) {
        console.error('로그 기록 중 오류(무시됨):', logErr.message);
      }
      checkAndNotifyLowStock(data); // Teams 저재고 알림
      return res.status(200).json({ success: true, message: '업데이트 완료', data: item });
    } else {
      return res.status(500).json({ success: false, message: 'OneDrive 업데이트 실패' });
    }
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.post('/api/inventory/manual-update', requireAuth, async (req, res) => {
  try {
    const { id, 현재수량, action, user } = req.body;
    const data = await fetchExcelFromOneDrive();
    const item = data.find(d => d.id == id);

    if (!item) {
      console.error(`❌ 항목 찾기 실패: 요청된 ID=${id}, 데이터 첫항목 ID=${data[0]?.id}`);
      return res.status(404).json({ success: false, message: '항목을 찾을 수 없습니다.' });
    }

    // 설비 확인이 필요한 항목(isCommonPart)은 반드시 common-update로만 처리
    if (item.isCommonPart && Number(현재수량) < Number(item.현재수량)) {
      return res.status(400).json({ success: false, message: '이 부품은 실사용 설비 확인이 필요합니다. common-update를 사용해 주세요.' });
    }

    const oldQuantity = item.현재수량;
    const qtyDelta = Number(현재수량) - Number(oldQuantity);
    item.현재수량 = 현재수량;
    item.최종수정시각 = getKSTDate();
    item.작업자 = user || 'Manual';

    const success = await updateExcelOnOneDrive(data);
    if (success) {
      try {
        const sharedLogId = uuidv4();
        addLog(action || '수정', item, qtyDelta, user || 'Manual', sharedLogId);
        addFacilityLog(action || '수정', item, qtyDelta, user || 'Manual', sharedLogId);
      } catch (logErr) {
        console.error('📝 로그 기록 오류(무시됨):', logErr.message);
      }
      checkAndNotifyLowStock(data); // Teams 저재고 알림
      return res.status(200).json({ success: true, message: '업데이트 완료', data: item });
    } else {
      return res.status(500).json({ success: false, message: 'OneDrive 업데이트 실패' });
    }
  } catch (error) {
    console.error('❌ manual-update 서버 에러:', error.message);
    return res.status(500).json({ success: false, message: error.message });
  }
});

// ============================================================
// ✨ 부품 신규 등록 / 삭제 — 수동 전용 API
// ⚠️ 절대 AI 챗봇(/api/ai/chat)에서 호출하거나 파싱하지 않는다.
//    프론트엔드의 전용 관리 화면(부품추가/부품삭제 버튼)에서만 직접 호출되어야 한다.
// ============================================================
app.post('/api/inventory/add-part', requireAuth, async (req, res) => {
  try {
    const { 대분류, 부품종류, 모델명, 적용설비, 현재수량, 최소보유수량, 용도, 보관장소, user } = req.body;

    // 필수값 검증 — 모델명, 부품종류는 반드시 있어야 부품 식별이 가능
    if (!모델명 || !String(모델명).trim()) {
      return res.status(400).json({ success: false, message: '모델명은 필수 입력입니다.' });
    }
    if (!부품종류 || !String(부품종류).trim()) {
      return res.status(400).json({ success: false, message: '부품종류는 필수 입력입니다.' });
    }

    invalidateCache();
    const data = await fetchExcelFromOneDrive();

    // 동일 모델명 중복 등록 방지
    const dup = data.find(d => String(d.모델명).trim() === String(모델명).trim());
    if (dup) {
      return res.status(409).json({ success: false, message: `이미 등록된 모델명입니다 (기존 재고: ${dup.현재수량}개).` });
    }

    const newItem = {
      id: `${CONFIG.inventorySheet}_${data.length + 1}`, // fetchExcelFromOneDrive가 다음 로드 시 인덱스 기준으로 재계산함
      원본시트: CONFIG.inventorySheet,
      대분류: 대분류 || '미분류',
      부품종류: String(부품종류).trim(),
      모델명: String(모델명).trim(),
      적용설비: 적용설비 || '',
      표준설비명: normalizeEquipment(적용설비 || ''),
      isCommonPart: true,
      현재수량: Number(현재수량) || 0,
      최소보유수량: Number(최소보유수량) || 0,
      최종수정시각: getKSTDate(), // 등록 시각 = 최초 최종수정시각
      작업자: user || 'Manual',
      용도: 용도 || '',
      보관장소: 보관장소 || '위치 미지정'
    };

    data.push(newItem);
    const success = await updateExcelOnOneDrive(data);

    if (success) {
      try {
        addLog('신규등록', newItem, newItem.현재수량, user || 'Manual');
      } catch (logErr) {
        console.error('📝 로그 기록 오류(무시됨):', logErr.message);
      }
      checkAndNotifyLowStock(data);
      return res.status(200).json({ success: true, message: '부품이 등록되었습니다.', data: newItem });
    } else {
      return res.status(500).json({ success: false, message: 'OneDrive 업데이트 실패' });
    }
  } catch (error) {
    console.error('❌ add-part 서버 에러:', error.message);
    return res.status(500).json({ success: false, message: error.message });
  }
});

app.post('/api/inventory/delete-part', requireAuth, async (req, res) => {
  try {
    const { id, user } = req.body;
    if (!id) {
      return res.status(400).json({ success: false, message: '삭제할 부품 id가 필요합니다.' });
    }

    invalidateCache();
    const data = await fetchExcelFromOneDrive();
    const idx = data.findIndex(d => d.id == id);

    if (idx === -1) {
      return res.status(404).json({ success: false, message: '해당 부품을 찾을 수 없습니다.' });
    }

    const deletedItem = data[idx];
    data.splice(idx, 1);

    const success = await updateExcelOnOneDrive(data);
    if (success) {
      try {
        // 변경수량은 삭제 시점 재고를 음수로 기록해 "얼마가 빠졌는지"를 이력에 남김
        addLog('삭제', { ...deletedItem, 현재수량: 0 }, -deletedItem.현재수량, user || 'Manual');
      } catch (logErr) {
        console.error('📝 로그 기록 오류(무시됨):', logErr.message);
      }
      console.log(`🗑️ 부품 삭제 — ${deletedItem.모델명} (${deletedItem.부품종류}), 삭제 전 재고 ${deletedItem.현재수량}개, 처리자: ${user || 'Manual'}`);
      return res.status(200).json({ success: true, message: '부품이 삭제되었습니다.', data: deletedItem });
    } else {
      return res.status(500).json({ success: false, message: 'OneDrive 업데이트 실패' });
    }
  } catch (error) {
    console.error('❌ delete-part 서버 에러:', error.message);
    return res.status(500).json({ success: false, message: error.message });
  }
});

// ============================================================
// ✨ Teams 웹훅 진단용 테스트 엔드포인트
// 저재고 조건과 무관하게 즉시 테스트 메시지를 보내보고, 성공/실패 및 실제 에러 내용을
// 그대로 응답으로 반환한다. 브라우저에서 이 URL을 열어보면 웹훅이 살아있는지 바로 확인 가능.
// (Office 365 Connectors 방식 웹훅은 Microsoft가 단계적으로 폐지 중이라, URL 자체가
//  이미 만료됐을 가능성을 직접 확인하기 위한 용도)
// ============================================================
app.get('/api/test-teams-alert', async (req, res) => {
  if (!CONFIG.teamsWebhookUrl) {
    return res.status(400).json({ success: false, message: 'TEAMS_WEBHOOK_URL 환경변수가 설정되지 않았습니다.' });
  }

  const testCard = {
    type: 'message',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: {
        type: 'AdaptiveCard',
        version: '1.4',
        body: [
          { type: 'TextBlock', text: '✅ 웹훅 테스트 메시지', weight: 'Bolder', size: 'Large', wrap: true },
          { type: 'TextBlock', text: `이 메시지가 보인다면 웹훅이 정상 작동 중입니다. (${getKSTDate()})`, wrap: true, isSubtle: true }
        ]
      }
    }]
  };

  try {
    const response = await axios.post(CONFIG.teamsWebhookUrl, testCard);
    console.log('✅ Teams 테스트 메시지 전송 성공');
    res.json({ success: true, message: '테스트 메시지 전송 성공 — Teams 채널을 확인하세요.', status: response.status });
  } catch (err) {
    // 웹훅이 폐지되어 만료됐다면 보통 404, 410(Gone), 400 등의 상태 코드나
    // "webhook URL is not valid" 류의 메시지가 그대로 담겨온다.
    console.error('❌ Teams 테스트 메시지 전송 실패:', err.response?.status, err.response?.data || err.message);
    res.status(500).json({
      success: false,
      message: '테스트 메시지 전송 실패',
      httpStatus: err.response?.status || null,
      errorDetail: err.response?.data || err.message
    });
  }
});

// ============================================================
// ✨ 이메일 인증 로그인 API
// 1) /request-code : 회사 이메일 도메인 검증 후 6자리 코드를 생성해 메일 발송
// 2) /verify-code  : 코드 확인. 최초 로그인이면 name 필수(신규 가입), 이후엔 이름 그대로 재사용
// 3) /logout       : 세션 폐기
// ============================================================
app.post('/api/auth/request-code', async (req, res) => {
  try {
    const { email } = req.body;
    if (!email || !String(email).includes('@')) {
      return res.status(400).json({ success: false, message: '올바른 이메일 주소를 입력해 주세요.' });
    }
    const normalizedEmail = String(email).trim().toLowerCase();

    if (!isAllowedCompanyEmail(normalizedEmail)) {
      return res.status(403).json({
        success: false,
        message: ALLOWED_EMAIL_DOMAIN
          ? `@${ALLOWED_EMAIL_DOMAIN} 이메일만 가입할 수 있습니다.`
          : '허용되지 않은 이메일입니다.'
      });
    }

    const code = generateOtpCode();
    otpStore.set(normalizedEmail, { code, expiresAt: Date.now() + OTP_DURATION_MS, attempts: 0 });

    await sendVerificationEmail(normalizedEmail, code);

    const users = loadUsers();
    const isNewUser = !users[normalizedEmail];

    res.json({ success: true, message: '인증코드를 전송했습니다. 이메일을 확인해 주세요.', isNewUser });
  } catch (error) {
    console.error('❌ 인증코드 발송 실패:', error.response?.data || error.message);
    res.status(500).json({ success: false, message: '인증코드 발송에 실패했습니다. 잠시 후 다시 시도해 주세요.' });
  }
});

app.post('/api/auth/verify-code', (req, res) => {
  try {
    const { email, code, name } = req.body;
    if (!email || !code) {
      return res.status(400).json({ success: false, message: '이메일과 인증코드를 입력해 주세요.' });
    }
    const normalizedEmail = String(email).trim().toLowerCase();
    const entry = otpStore.get(normalizedEmail);

    if (!entry) {
      return res.status(400).json({ success: false, message: '인증코드를 먼저 요청해 주세요.' });
    }
    if (Date.now() > entry.expiresAt) {
      otpStore.delete(normalizedEmail);
      return res.status(400).json({ success: false, message: '인증코드가 만료됐습니다. 다시 요청해 주세요.' });
    }
    if (entry.attempts >= OTP_MAX_ATTEMPTS) {
      otpStore.delete(normalizedEmail);
      return res.status(429).json({ success: false, message: '시도 횟수를 초과했습니다. 인증코드를 다시 요청해 주세요.' });
    }
    if (String(code).trim() !== entry.code) {
      entry.attempts += 1;
      return res.status(400).json({ success: false, message: `인증코드가 일치하지 않습니다. (${OTP_MAX_ATTEMPTS - entry.attempts}회 남음)` });
    }

    // 코드 일치 — 성공 처리, 재사용 방지를 위해 즉시 폐기
    otpStore.delete(normalizedEmail);

    const users = loadUsers();
    let userRecord = users[normalizedEmail];

    if (!userRecord) {
      // 최초 로그인 = 신규 가입. 표시 이름이 필요하다.
      if (!name || !String(name).trim()) {
        return res.status(400).json({ success: false, message: '최초 로그인입니다. 이름을 입력해 주세요.', requiresName: true });
      }
      userRecord = { name: String(name).trim(), createdAt: getKSTDate() };
      users[normalizedEmail] = userRecord;
      saveUsers(users);
      console.log(`✅ 신규 가입: ${normalizedEmail} (${userRecord.name})`);
    }

    const token = createSession(normalizedEmail, userRecord.name);
    res.json({ success: true, token, email: normalizedEmail, name: userRecord.name });
  } catch (error) {
    console.error('❌ 인증코드 확인 실패:', error.message);
    res.status(500).json({ success: false, message: error.message });
  }
});

app.post('/api/auth/logout', (req, res) => {
  const authHeader = req.headers['authorization'] || '';
  const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : null;
  if (token) destroySession(token);
  res.json({ success: true });
});


// 실제로 최소보유수량 이하인 부품이 몇 개인지, 그리고 각 부품이 쿨다운(1시간)에
// 걸려있는지를 직접 보여준다. "웹훅은 되는데 알림이 안 온다"의 원인을
// (1) 조건 미충족 (2) 쿨다운 (3) 다른 문제 중 어디인지 좁히기 위한 용도.
// ============================================================
// ============================================================
// ✨ 서버 상태 진단 — 로딩이 느릴 때 이 URL을 열어보면, 서버가 방금 막 재시작된 건지
// (콜드스타트 직후라 느린 게 당연함) 아니면 오래 켜져 있었는데도 느린 건지(다른 원인)
// 바로 구분할 수 있다.
// ============================================================
app.get('/api/server-status', (req, res) => {
  const now = Date.now();
  const uptimeMs = now - SERVER_START_TIME;
  const cacheAgeMs = lastFetchTime ? now - lastFetchTime : null;

  // ✨ 로컬 파일 영속성 확인용 — 재배포 전후로 이 값들을 비교해보면,
  //    Northflank가 Render처럼 파일시스템을 재배포마다 초기화하는지 바로 알 수 있다.
  //    (파일이 존재하고 mtime이 예전 배포 시점 그대로라면 = 영속적, 방금 시각이라면 = 초기화됨)
  const checkFile = (filePath) => {
    try {
      if (fs.existsSync(filePath)) {
        const stat = fs.statSync(filePath);
        return { exists: true, modifiedAt: stat.mtime.toISOString(), sizeBytes: stat.size };
      }
      return { exists: false };
    } catch (e) {
      return { exists: false, error: e.message };
    }
  };

  res.json({
    success: true,
    serverStartedAt: new Date(SERVER_START_TIME).toISOString(),
    uptimeMinutes: Math.round(uptimeMs / 60000 * 10) / 10,
    uptimeSeconds: Math.round(uptimeMs / 1000),
    // uptimeSeconds가 작다면(예: 몇십 초 이내) 방금 서버가 재시작됐다는 뜻 —
    // 재배포였는지, 아니면 슬립 후 깨어난 것인지는 플랫폼 대시보드의 Events/Deploys 탭에서 확인 가능
    cacheStatus: cachedData ? '있음' : '없음',
    cacheAgeSeconds: cacheAgeMs !== null ? Math.round(cacheAgeMs / 1000) : null,
    cacheDurationSeconds: CACHE_DURATION / 1000,
    localFilePersistence: {
      onedrive_tokens: checkFile(TOKEN_FILE),
      inventory_logs: checkFile(LOG_FILE),
      user_auth: checkFile(USER_AUTH_FILE),
      sessions: checkFile(SESSION_FILE)
    }
  });
});

app.get('/api/debug-low-stock', async (req, res) => {
  try {
    const data = await fetchExcelFromOneDrive();
    const lowStock = data.filter(d => d.최소보유수량 > 0 && d.현재수량 < d.최소보유수량);

    const now = Date.now();
    const detail = lowStock.map(item => {
      const lastAlerted = alertCooldown.get(item.id) || 0;
      const msSinceLastAlert = lastAlerted ? now - lastAlerted : null;
      const onCooldown = lastAlerted ? (now - lastAlerted < ALERT_COOLDOWN_MS) : false;
      return {
        id: item.id,
        모델명: item.모델명,
        현재수량: item.현재수량,
        최소보유수량: item.최소보유수량,
        onCooldown,
        minutesUntilCooldownEnds: onCooldown ? Math.ceil((ALERT_COOLDOWN_MS - msSinceLastAlert) / 60000) : 0
      };
    });

    res.json({
      success: true,
      totalLowStockCount: lowStock.length,
      willActuallyAlertCount: detail.filter(d => !d.onCooldown).length,
      teamsWebhookConfigured: !!CONFIG.teamsWebhookUrl,
      cooldownDurationMinutes: ALERT_COOLDOWN_MS / 60000,
      items: detail
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.get('/api/inventory/alerts', async (req, res) => {
  try {
    const data = await fetchExcelFromOneDrive();
    const alerts = data
      .filter(item => item.최소보유수량 > 0 && item.현재수량 < item.최소보유수량)
      .map(item => ({
        id: item.id,
        부품종류: item.부품종류,
        모델명: item.모델명,
        적용설비: item.적용설비,
        현재수량: item.현재수량,
        최소보유수량: item.최소보유수량,
        부족수량: item.최소보유수량 - item.현재수량,
        긴급도: item.현재수량 === 0 ? 'critical' : 'warning'
      }))
      .sort((a, b) => {
        if (a.긴급도 === 'critical' && b.긴급도 !== 'critical') return -1;
        if (a.긴급도 !== 'critical' && b.긴급도 === 'critical') return 1;
        return b.부족수량 - a.부족수량;
      });
    res.json({ success: true, data: alerts, count: alerts.length });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.get('/api/inventory/logs', (req, res) => {
  try {
    let logs = loadLogs();
    const limit    = parseInt(req.query.limit)  || 100;
    const offset   = parseInt(req.query.offset) || 0;
    const facility = req.query.facility ? String(req.query.facility) : null;
    const partType = req.query.partType ? String(req.query.partType) : null;
    // ✨ action 필터: '입고'/'출고' 등으로 지정하면 그 종류만 조회.
    //    입고 로그가 압도적으로 많이 쌓이는 경우, 최근 N건을 통째로 가져오면 출고 로그가
    //    전부 밀려나 "내역 없음"으로 보이는 문제가 있었다. 프론트가 입고/출고를 각각
    //    독립적으로 조회할 수 있도록 지원한다.
    const action   = req.query.action ? String(req.query.action) : null;
    // ✨ excludeActions: 콤마로 구분된 액션 목록을 제외 — '기타' 컬럼(수량변경/신규등록/삭제 등)이
    //    입고/출고에 밀려 안 보이는 일이 없도록 독립적으로 조회할 때 사용
    const excludeActions = req.query.excludeActions
      ? String(req.query.excludeActions).split(',').map(s => s.trim()).filter(Boolean)
      : null;
    // ✨ 통합 검색어: 설비명/모델명/부품종류 어디에든 포함되면 매칭 (공백/하이픈/언더스코어 무시)
    const normalize = (s) => String(s || '').toLowerCase().replace(/[\s\-_]+/g, '');
    const q = req.query.q ? normalize(req.query.q) : null;

    if (facility) {
      logs = logs.filter(l =>
        String(l.적용설비 || '').includes(facility) ||
        String(l.표준설비명 || '').includes(facility)
      );
    }
    if (partType) {
      logs = logs.filter(l => String(l.부품종류 || '') === partType);
    }
    if (action) {
      logs = logs.filter(l => String(l.action || '') === action);
    }
    if (excludeActions) {
      logs = logs.filter(l => !excludeActions.includes(String(l.action || '')));
    }
    if (q) {
      logs = logs.filter(l =>
        normalize(l.적용설비).includes(q) ||
        normalize(l.표준설비명).includes(q) ||
        normalize(l.모델명).includes(q) ||
        normalize(l.부품종류).includes(q)
      );
    }

    const total = logs.length;
    res.json({ success: true, data: logs.slice(offset, offset + limit), total });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

// 설비이력 전체 조회 (설비명 필터 가능)
app.get('/api/inventory/facility-logs', (req, res) => {
  try {
    const limit = parseInt(req.query.limit) || 200;
    const facility = req.query.facility ? String(req.query.facility) : null;
    const isCommon = req.query.isCommon === 'true'; // 공통탭 여부
    let logs = facilityLogs;
    if (facility) {
      logs = logs.filter(l => {
        // 일반 설비: 표준설비명 또는 원본설비명 매칭
        if (l.표준설비명 === facility || l.원본설비명 === facility) return true;
        // 공통부품 출고 이력: 원본시트가 '공통'이고 실제사용설비(표준설비명에 저장)가 매칭
        if (isCommon && l.isCommonPart && l.표준설비명 === facility) return true;
        // 어떤 설비든 공통부품 이력을 원본시트 기반으로 조회 (공통탭 대시보드용)
        if (isCommon && l.원본시트 === '공통') return true;
        return false;
      });
    }
    res.json({ success: true, data: logs.slice(0, limit), total: logs.length });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

// 설비별 출고 요약 (대시보드용)
app.get('/api/inventory/facility-summary', async (req, res) => {
  try {
    // 설비별 총 출고 건수/수량 집계
    const summary = {};
    facilityLogs.forEach(log => {
      const facility = log.표준설비명 || log.원본설비명 || '미분류';
      if (!summary[facility]) {
        summary[facility] = { 표준설비명: facility, 출고건수: 0, 출고수량: 0, 입고건수: 0, 입고수량: 0, 최근이력: null };
      }
      const qty = Math.abs(Number(log.변경수량) || 0);
      if (log.action === '출고' || (log.변경수량 < 0)) {
        summary[facility].출고건수 += 1;
        summary[facility].출고수량 += qty;
      } else {
        summary[facility].입고건수 += 1;
        summary[facility].입고수량 += qty;
      }
      if (!summary[facility].최근이력) {
        summary[facility].최근이력 = log.timestampKR;
      }
    });
    res.json({ success: true, data: Object.values(summary) });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

// ============================================================
// ✨ 전체 사용내역 요약 — 메인화면 "전체 사용내역 요약" 카드용
// 설비이력(facilityLogs) 전체를 기준으로 월별/연도별 집계, 설비별/부품별 랭킹을 계산해 반환
// ============================================================
function parseKSTDateServer(raw) {
  if (!raw) return null;
  const m = String(raw).match(/(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})/);
  if (!m) return null;
  return new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
}

app.get('/api/inventory/usage-summary', (req, res) => {
  try {
    // 출고 이력만 대상으로 함 ("사용"의 의미이므로 입고는 제외)
    // ✨ '쓰레기통' 설비는 실제 소모성 부품 사용이 아니라 폐기물 처리 성격이라 집계에서 제외
    const EXCLUDED_FACILITIES = ['쓰레기통'];
    const outLogs = facilityLogs.filter(l =>
      (l.변경수량 < 0 || l.action === '출고') &&
      !EXCLUDED_FACILITIES.includes(l.표준설비명 || l.원본설비명)
    );

    // ---- 월별 집계 ----
    const monthlyMap = {}; // 'YYYY-MM' -> { total, byFacility: {}, byPart: {} }
    // ---- 연도별 집계 ----
    const yearlyMap = {}; // 'YYYY' -> { total, byFacility: {}, byPart: {} }
    // ---- 설비별 전체 랭킹 ----
    const facilityTotals = {}; // 표준설비명 -> { facility, total, count }
    // ---- 부품별 전체 랭킹 ----
    const partTotals = {}; // 모델명 -> { model, partType, total, count }
    // ---- 설비별 부품 사용 내역 (설비 랭킹 클릭 시 "이 설비에서 어떤 부품이 쓰였는지" 표시용) ----
    const facilityPartBreakdown = {}; // 표준설비명 -> { 모델명: qty }

    outLogs.forEach(log => {
      const d = parseKSTDateServer(log.timestampKR);
      if (!d) return;

      const qty = Math.abs(Number(log.변경수량) || 0);
      const facility = log.표준설비명 || log.원본설비명 || '미분류';
      const model = log.모델명 || '미상';
      const partType = log.부품종류 || '';

      const monthKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
      const yearKey = `${d.getFullYear()}`;

      if (!monthlyMap[monthKey]) monthlyMap[monthKey] = { total: 0, byFacility: {}, byPart: {}, byFacilityPart: {} };
      monthlyMap[monthKey].total += qty;
      monthlyMap[monthKey].byFacility[facility] = (monthlyMap[monthKey].byFacility[facility] || 0) + qty;
      monthlyMap[monthKey].byPart[model] = (monthlyMap[monthKey].byPart[model] || 0) + qty;
      if (!monthlyMap[monthKey].byFacilityPart[facility]) monthlyMap[monthKey].byFacilityPart[facility] = {};
      monthlyMap[monthKey].byFacilityPart[facility][model] =
        (monthlyMap[monthKey].byFacilityPart[facility][model] || 0) + qty;

      if (!yearlyMap[yearKey]) yearlyMap[yearKey] = { total: 0, byFacility: {}, byPart: {}, byFacilityPart: {} };
      yearlyMap[yearKey].total += qty;
      yearlyMap[yearKey].byFacility[facility] = (yearlyMap[yearKey].byFacility[facility] || 0) + qty;
      yearlyMap[yearKey].byPart[model] = (yearlyMap[yearKey].byPart[model] || 0) + qty;
      if (!yearlyMap[yearKey].byFacilityPart[facility]) yearlyMap[yearKey].byFacilityPart[facility] = {};
      yearlyMap[yearKey].byFacilityPart[facility][model] =
        (yearlyMap[yearKey].byFacilityPart[facility][model] || 0) + qty;

      if (!facilityTotals[facility]) facilityTotals[facility] = { facility, total: 0, count: 0 };
      facilityTotals[facility].total += qty;
      facilityTotals[facility].count += 1;

      if (!partTotals[model]) partTotals[model] = { model, partType, total: 0, count: 0 };
      partTotals[model].total += qty;
      partTotals[model].count += 1;

      if (!facilityPartBreakdown[facility]) facilityPartBreakdown[facility] = {};
      facilityPartBreakdown[facility][model] = (facilityPartBreakdown[facility][model] || 0) + qty;
    });

    // ✨ 월별 배열 범위: 가장 오래된 이력 시점(또는 이번 달 12개월 전 중 더 이른 쪽)부터
    //    이번 달 + 12개월 뒤까지 전부 생성한다. 과거 이력도 확인할 수 있어야 하므로
    //    이번 달만 남기고 자르지 않는다 — 대신 프론트에서 "이번 달이 기본으로 보이는 위치"를
    //    currentMonthIndex 기준으로 계산해 보여준다.
    const now = new Date();
    let earliestMonth = null;
    Object.keys(monthlyMap).sort().forEach(key => {
      if (!earliestMonth) earliestMonth = key;
    });
    const fallbackStart = new Date(now.getFullYear(), now.getMonth() - 12, 1);
    let startDate;
    if (earliestMonth) {
      const [ey, em] = earliestMonth.split('-').map(Number);
      const earliestAsDate = new Date(ey, em - 1, 1);
      startDate = earliestAsDate < fallbackStart ? earliestAsDate : fallbackStart;
    } else {
      startDate = new Date(now.getFullYear(), now.getMonth(), 1); // 이력이 전혀 없으면 이번 달부터 시작
    }
    const endDate = new Date(now.getFullYear(), now.getMonth() + 12, 1); // 이번 달부터 12개월 뒤까지 여유분

    const monthly = [];
    let cursor = new Date(startDate);
    while (cursor <= endDate) {
      const key = `${cursor.getFullYear()}-${String(cursor.getMonth() + 1).padStart(2, '0')}`;
      const label = `${cursor.getFullYear()}.${cursor.getMonth() + 1}`;
      const entry = monthlyMap[key] || { total: 0, byFacility: {}, byPart: {}, byFacilityPart: {} };
      monthly.push({
        key, label, total: entry.total,
        byFacility: entry.byFacility, byPart: entry.byPart, byFacilityPart: entry.byFacilityPart
      });
      cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
    }

    // 이번 달이 실제로 몇 번째 인덱스인지 계산해서 내려준다 (프론트가 이 인덱스를 기준으로
    // 6개월 창의 가운데 근처에 이번 달이 오도록 배치함)
    const currentMonthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    let currentMonthIndex = monthly.findIndex(m => m.key === currentMonthKey);
    if (currentMonthIndex === -1) currentMonthIndex = monthly.length - 1;

    // 연도별은 있는 것만 오름차순
    const yearly = Object.keys(yearlyMap).sort().map(key => ({
      key, label: `${key}년`, total: yearlyMap[key].total,
      byFacility: yearlyMap[key].byFacility, byPart: yearlyMap[key].byPart,
      byFacilityPart: yearlyMap[key].byFacilityPart
    }));

    const facilityRanking = Object.values(facilityTotals).sort((a, b) => b.total - a.total);
    const partRanking = Object.values(partTotals).sort((a, b) => b.total - a.total);

    res.json({
      success: true,
      data: {
        monthly,
        currentMonthIndex,
        yearly,
        facilityRanking,
        partRanking,
        facilityPartBreakdown,
        totalOutCount: outLogs.length,
        totalOutQty: outLogs.reduce((s, l) => s + Math.abs(Number(l.변경수량) || 0), 0)
      }
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.get('/api/inventory/search', async (req, res) => {
  try {
    // ✨ 검색어 정규화: 공백/하이픈/언더스코어를 무시하고 비교 (예: "PA-12" = "PA12" = "PA 12")
    const normalize = (s) => String(s || '').toLowerCase().replace(/[\s\-_]+/g, '');
    const query = normalize(req.query.q);
    const data = await fetchExcelFromOneDrive();

    if (!Array.isArray(data)) {
      return res.json({ success: true, data: [] });
    }

    const filtered = data.filter(item => {
      if (!item) return false;
      const model = normalize(item.모델명);
      const type = normalize(item.부품종류);
      const facility = normalize(item.적용설비);
      const mainCat = normalize(item.대분류);
      return model.includes(query) || type.includes(query) || facility.includes(query) || mainCat.includes(query);
    });

    res.json({ success: true, data: filtered });
  } catch (error) {
    console.error('❌ 검색 API 내부 에러:', error.stack);
    res.status(500).json({ success: false, message: '서버 내부 오류' });
  }
});

// ============================================================
// 방안C: 공통부품 출고 — 실제 사용 설비 포함 처리
// POST /api/inventory/common-update
// body: { id, 현재수량, action, user, 실제사용설비 }
// ============================================================
app.post('/api/inventory/common-update', requireAuth, async (req, res) => {
  try {
    const { id, 현재수량, action, user, 실제사용설비 } = req.body;

    if (!실제사용설비) {
      return res.status(400).json({ success: false, message: '공통부품 출고 시 실제사용설비는 필수입니다.' });
    }

    const data = await fetchExcelFromOneDrive();
    const item = data.find(d => d.id == id);
    if (!item) return res.status(404).json({ success: false, message: '항목을 찾을 수 없습니다.' });

    // ✨ 설비 유효성 검증: 부품마다 후보설비목록을 중복 저장하지 않게 된 이후로는
    //    전역 facilityListCache.all(전체 설비 목록)을 직접 참조해 오타/임의 입력을 막는다.
    const norm실제사용설비 = normalizeEquipment(실제사용설비);
    const validFacilities = facilityListCache?.all || [];
    if (validFacilities.length > 0 && !validFacilities.includes(norm실제사용설비)) {
      return res.status(400).json({ success: false, message: `"${실제사용설비}"는 이 부품의 사용 가능 설비 목록에 없습니다.` });
    }

    const oldQuantity = item.현재수량;
    item.현재수량 = 현재수량;
    item.최종수정시각 = getKSTDate();
    item.작업자 = user || 'Manual';

    const success = await updateExcelOnOneDrive(data);
    if (success) {
      // 로그에는 정규화된 실제 설비명 기록 (엑셀 원본 표기 흔들림 방지)
      const logItem = { ...item, 적용설비: norm실제사용설비, 표준설비명: norm실제사용설비, isCommonPart: true };
      const sharedLogId = uuidv4();
      addLog(action || '출고', logItem, 현재수량 - oldQuantity, user || 'Manual', sharedLogId);
      addFacilityLog(action || '출고', logItem, 현재수량 - oldQuantity, user || 'Manual', sharedLogId);
      checkAndNotifyLowStock(data);
      console.log(`🏭 공통부품 출고 기록 — ${item.모델명} → 실제설비: ${norm실제사용설비}`);
      return res.status(200).json({ success: true, message: '공통부품 출고 완료', data: item });
    } else {
      return res.status(500).json({ success: false, message: 'OneDrive 업데이트 실패' });
    }
  } catch (error) {
    console.error('❌ common-update 에러:', error.message);
    return res.status(500).json({ success: false, message: error.message });
  }
});

// ============================================================
// 이력 되돌리기(롤백) — 실수로 출고/입고 처리한 것을 취소
// POST /api/inventory/rollback-log
// body: { logId, user }
// 안전장치:
//  1) 해당 부품(모델명+부품종류) 기준으로 "가장 최근" 이력만 되돌릴 수 있음
//     (중간 이력을 되돌리면 그 이후 이력들과 수량이 어긋나기 때문)
//  2) 현재 재고 수량이 이 로그의 변경후수량과 정확히 일치해야만 진행
//     (그 사이에 다른 경로로 재고가 바뀌었다면 안전하게 거부)
//  3) 이미 되돌린 이력은 다시 되돌릴 수 없음
// ============================================================
// ── 롤백 핵심 로직 (수동 화면의 '되돌리기' 버튼 / AI챗봇 양쪽에서 공유) ──
// data: fetchExcelFromOneDrive()로 가져온 재고 배열. 성공 시 이 배열을 직접
//       수정(mutate)하고 facilityLogs/memoryLogs에서도 이력을 제거한다.
//       호출부는 성공 시 이어서 updateExcelOnOneDrive(data)만 호출하면 된다.
//       ⚠️ 이 함수는 절대 반대 방향의 입고/출고(addLog/addFacilityLog)를 새로
//          만들지 않는다 — 반드시 "원본 이력 자체를 삭제"하는 방식으로만 되돌린다.
//          (반대 거래를 새로 쌓으면 재고 수량은 맞아떨어져 보여도, 잘못 기록된
//          원본 사용이력이 그대로 남아 사용내역이 롤백되지 않는 문제가 생긴다.)
function rollbackLogEntry(log, data, user) {
  // 같은 부품(모델명+부품종류)의 이력 중 가장 최근 것인지 확인
  const relatedLogs = facilityLogs.filter(l => l.모델명 === log.모델명 && l.부품종류 === log.부품종류);
  const latestLogForItem = relatedLogs[0]; // facilityLogs는 항상 최신순 정렬됨
  if (!latestLogForItem || latestLogForItem.id !== log.id) {
    return { success: false, message: '이후에 재고 변동이 있어 이 이력은 되돌릴 수 없습니다. (가장 최근 이력만 취소 가능)' };
  }

  const item = data.find(d => d.모델명 === log.모델명 && d.부품종류 === log.부품종류);
  if (!item) {
    return { success: false, message: '해당 부품을 현재 재고에서 찾을 수 없습니다.' };
  }

  // 현재 재고가 이 로그가 남겼던 결과값과 일치하는지 확인
  if (Number(item.현재수량) !== Number(log.변경후수량)) {
    return { success: false, message: '현재 재고 수량이 이력과 일치하지 않아 되돌릴 수 없습니다.' };
  }

  const restoredQty = Number(log.변경전수량);
  const oldQuantity = item.현재수량;
  item.현재수량 = restoredQty;
  item.최종수정시각 = getKSTDate();
  item.작업자 = user || 'Manual';

  // ── 이력 자체를 완전히 제거 (되돌린 건은 사용내역에 아예 남지 않도록) ──
  // 설비이력(facilityLogs)에서 제거
  const facilityIdx = facilityLogs.findIndex(l => l.id === log.id);
  if (facilityIdx !== -1) facilityLogs.splice(facilityIdx, 1);

  // 사용내역종합(memoryLogs, addLog로 쌓이는 로그)에서도 같은 id로 저장된 짝을 제거
  // (같은 이벤트를 남기는 addLog/addFacilityLog가 sharedLogId로 연결되어 있음)
  const generalLogs = loadLogs();
  const filteredGeneralLogs = generalLogs.filter(l => l.id !== log.id);
  if (filteredGeneralLogs.length !== generalLogs.length) {
    saveLogs(filteredGeneralLogs);
  }

  return { success: true, item, oldQuantity, restoredQty };
}

app.post('/api/inventory/rollback-log', requireAuth, async (req, res) => {
  try {
    const { logId, user } = req.body;
    if (!logId) {
      return res.status(400).json({ success: false, message: 'logId가 필요합니다.' });
    }

    const log = facilityLogs.find(l => l.id === logId);
    if (!log) {
      return res.status(404).json({ success: false, message: '해당 이력을 찾을 수 없습니다.' });
    }

    const data = await fetchExcelFromOneDrive();
    const result = rollbackLogEntry(log, data, user);
    if (!result.success) {
      return res.status(400).json(result);
    }

    // 재고 + (이력이 제거된) 로그 시트들을 함께 저장
    const success = await updateExcelOnOneDrive(data);
    if (!success) {
      return res.status(500).json({ success: false, message: 'OneDrive 업데이트 실패' });
    }

    checkAndNotifyLowStock(data);
    console.log(`↩️ 이력 되돌리기 — ${result.item.모델명} (${log.action}) 취소 및 이력 삭제, 재고 ${result.oldQuantity} → ${result.restoredQty}`);
    return res.status(200).json({ success: true, message: '되돌리기 완료', data: result.item });
  } catch (error) {
    console.error('❌ rollback-log 에러:', error.message);
    return res.status(500).json({ success: false, message: error.message });
  }
});

app.post('/api/ai/chat', requireAuth, async (req, res) => {
  try {
    const { message, conversationHistory } = req.body;
    // ✨ 클라이언트가 보낸 user 값을 그대로 신뢰하지 않는다 (위조 가능).
    //    로그인 세션에서 검증된 이름을 사용한다.
    const user = req.authUser.name;

    // ✨ 이전에는 매 메시지마다 invalidateCache()로 캐시를 강제 무효화해
    //    엑셀 전체를 재다운로드했다 (챗봇 응답이 느려지는 주된 원인).
    //    이제는 일반 조회 API와 동일하게 60초 캐시(CACHE_DURATION)를 그대로 활용한다.
    //    → 최대 60초 정도의 재고 반영 지연이 생길 수 있지만, 매 대화마다
    //      OneDrive 풀 다운로드가 발생하던 문제는 사라진다.
    let inventoryData = await fetchExcelFromOneDrive();

    // 전체 특정 설비 목록(예시/일반 안내용)
    const realFacilities = (facilityListCache?.all || []).slice().sort();

    // ============================================================
    // ✨ 토큰 절약: 매 메시지마다 전체 재고를 통째로 보내지 않는다.
    //    재고 품목 수가 많으면(FULL_LIST_THRESHOLD 초과) 이번 메시지+최근 대화와
    //    관련 있어 보이는 부품과 재고부족 품목만 골라서 system instruction에 넣는다.
    //    → Gemini free tier의 "분당 입력 토큰" 한도(250,000) 초과로 인한 429 방지.
    //    실제 입출고 처리(findCandidates 등)는 아래에서 여전히 전체 inventoryData를
    //    사용하므로 기능에는 영향 없다 — 여기서 줄이는 건 어디까지나 "프롬프트에 적는 목록"뿐.
    // ============================================================
    const normalizeForSearch = (s) => String(s || '').toLowerCase().replace(/[\s\-_]+/g, '');

    const BROAD_QUERY_KEYWORDS = ['전체', '전부', '모두', '모든', '다 보여', '리스트', '목록 보여', '재고 현황', '재고현황', '부족한 부품', '부족 목록'];
    const isBroadQuery = BROAD_QUERY_KEYWORDS.some(k => String(message || '').includes(k));

    // 이번 대화에서 참고할 최근 맥락(직전 몇 턴 + 이번 메시지) — "그거 얼마나 있어?" 같은 대명사 참조에도 어느 정도 대응
    const recentContextRaw = [...(conversationHistory || []).slice(-4).map(m => m.text), message].join(' ');
    const contextNoSpace = normalizeForSearch(recentContextRaw);
    const contextWords = recentContextRaw
      .split(/[\s,.\/!?()~\-·]+/)
      .map(w => w.trim())
      .filter(w => w.length >= 2);

    const FULL_LIST_THRESHOLD = 150; // 이 개수 이하면 굳이 필터링 안 해도 토큰 부담이 크지 않음

    let itemsForPrompt;
    let isFilteredSubset = false;

    if (isBroadQuery || inventoryData.length <= FULL_LIST_THRESHOLD) {
      itemsForPrompt = inventoryData;
    } else {
      const matched = inventoryData.filter(item => {
        const model = normalizeForSearch(item.모델명);
        const type = normalizeForSearch(item.부품종류);
        const cat = normalizeForSearch(item.대분류);
        if (model && contextNoSpace.includes(model)) return true; // 모델명을 그대로 붙여 쓴 경우
        return [type, cat, model].some(f => f && f.length >= 2 && contextWords.some(w => f.includes(w)));
      });

      // 재고 부족 품목은 질문과 무관해도 중요한 경고 정보이므로 항상 포함
      const lowStock = inventoryData.filter(item => item.최소보유수량 > 0 && item.현재수량 < item.최소보유수량);

      const combinedMap = new Map();
      [...matched, ...lowStock].forEach(item => combinedMap.set(item.id, item));
      itemsForPrompt = Array.from(combinedMap.values()).slice(0, 120); // 안전 상한
      isFilteredSubset = true;
    }

    // ── 모든 부품이 공통 시트 소속이며, 전체 설비 목록 중에서 실사용 설비를 확인해야 하는 구조 ──
    const inventoryTable = itemsForPrompt.map(item => {
      const stockStatus = item.최소보유수량 > 0 && item.현재수량 < item.최소보유수량 ? '⚠️부족' : '정상';
      return `모델명:${item.모델명} | 부품종류:${item.부품종류} | 현재수량:${item.현재수량} | 최소보유:${item.최소보유수량} | 재고:${stockStatus}`;
    }).join('\n');

    const listNote = isFilteredSubset
      ? `⚠️ 아래 [최신 재고 현황]은 전체 재고가 아니라, 이번 질문과 관련 있어 보이는 부품 + 재고부족 품목만 골라서 보여준 "부분 목록"이다 (전체 ${inventoryData.length}건 중 ${itemsForPrompt.length}건 표시). 이 목록에 없다고 해서 "등록된 부품이 아니다"라고 단정짓지 말 것 — 대신 정확한 모델명이나 부품종류를 다시 물어볼 것. (사용자가 "전체 목록 보여줘"처럼 다시 물으면 전체를 보여줄 수 있다고 안내할 것)`
      : `아래 [최신 재고 현황]은 전체 재고 목록이다. 이 목록에 없는 부품은 "등록된 부품이 아닙니다"라고 명확히 답할 것.`;

    // ── 수정3: system_instruction 파라미터로 분리 ──
    const systemInstruction = `당신은 스마트 재고 관리 AI 어시스턴트입니다.
반드시 아래 [최신 재고 현황]만을 근거로 답변하세요.
${listNote}

[최신 재고 현황]
${inventoryTable}

[전체 설비 목록]
${realFacilities.join(', ')}

[설비 목록에 대한 설명]
- 위 [전체 설비 목록]은 원본 엑셀 "충전"/"타정" 시트의 "적용설비" 헤더 열에 등록된 설비명을
  그대로 가져온 것이다. 이 목록에 있는 표기만 유효한 설비명이며, 목록에 없는 이름은 절대
  존재하지 않는 것으로 간주한다.
- 같은 설비가 1공장/2공장 양쪽에 있는 경우, 설비명 뒤에 "(1공장)", "(2공장)"처럼 공장 표기가
  붙어 구분된다. 예: "유성충전기 (1공장)", "유성충전기 (2공장)"은 이름은 비슷하지만 서로 다른
  별개의 설비이므로 절대 혼동하거나 하나로 합쳐 판단하지 말 것.

[절대 준수 규칙]
1. 마크다운 코드블록(\`\`\`json) 절대 금지. 반드시 ~~~ 기호만 사용.
2. 한 번에 50개 이상 변동 요청 시 두 번 재확인할 것.
3. 모델명 매칭 시 공백·대소문자 차이는 무시하고 찾을 것.
4. 재고 현황에 없는 부품에 대한 판단은 위 [최신 재고 현황] 바로 위의 안내(전체 목록인지 부분 목록인지)를 따를 것.

[설비 확인 규칙 — 예외 없음]
5. 모든 부품은 여러 설비가 공용으로 쓰므로, 출고(사용) 요청 시 사용자가 이번 메시지 또는
   직전 대화에서 이미 구체적인 설비(호기, 공장 포함)를 명시하지 않은 이상 절대로
   INVENTORY_UPDATE 명령을 생성하지 말고, 반드시 먼저 "어느 설비에 사용하셨나요?"라고
   물어볼 것. [전체 설비 목록] 중 관련성 높아 보이는 몇 개를 예시로 제시해도 좋다.
6. 반대로, 사용자가 이번 메시지나 직전 대화에서 이미 설비명을 구체적으로 말했다면
   (예: "1호기에 썼어요", "유성충전기 2공장 것") 절대 다시 "어느 설비에 사용하셨나요?"라고
   되묻지 말 것. 이미 답변받은 내용을 또 물어보는 것은 규칙 위반이다. 이 경우 바로 6-1로 진행.
   6-1. 사용자가 말한 설비명이 [전체 설비 목록]에 있는 정확한 표기와 일치하는지 확인한다.
        - 일치하면(공백/대소문자 차이는 무시) 그 정확한 표기 그대로 실제사용설비 필드에
          넣어 INVENTORY_UPDATE 명령을 생성한다.
        - 일치하지 않으면, 절대로 목록에 없는 이름을 임의로 추측해서 채우거나 비슷한 이름으로
          지레짐작하지 말 것. 대신 "말씀하신 설비명은 목록에서 찾을 수 없습니다. 정확한
          설비명을 다시 한 번 말씀해 주시겠어요?"라고 답하고, 표기가 헷갈릴 만한 후보
          (예: 1공장/2공장 버전이 둘 다 있는 경우 그 두 가지)를 함께 보여줄 것.
7. 입고(재고 보충)는 설비 확인 없이 바로 처리 가능하다.
8. "그냥 출고", "확인 생략" 요청에도 설비 확인 절대 생략 금지.

[되돌리기(롤백) 요청 처리 규칙 — 예외 없음]
9. 사용자가 "롤백해줘", "취소해줘", "되돌려줘", "방금 처리한 거 잘못했어" 등으로 직전에
   처리한 입고/출고를 취소하고 싶어하는 경우, 절대로 반대 방향의 입고/출고를
   INVENTORY_UPDATE로 새로 만들어서 수량만 맞추려고 하지 말 것. 그렇게 하면 재고 수량은
   맞아 보여도 잘못 기록된 원본 사용이력이 삭제되지 않고 그대로 남는다. 되돌리기는 반드시
   아래 10번 형식의 전용 명령(INVENTORY_ROLLBACK)만 사용해서 "원본 이력 자체를 삭제"하는
   방식으로 처리한다.
10. 사용자가 되돌리려는 부품의 모델명(그리고 가능하면 부품종류)이 이번 메시지나 직전
    대화에서 명확히 특정되지 않았다면, 절대로 짐작해서 명령을 만들지 말고 먼저 "어떤
    부품의 어떤 처리를 되돌리시겠어요? (모델명을 말씀해 주세요)"라고 되물을 것. 직전
    대화에서 방금 자신이 처리한 항목이 명확하다면 그것을 그대로 사용해도 된다.
11. 롤백은 "해당 부품의 가장 최근 이력 1건"만 대상이 된다는 점, 그리고 이후 다른 변동이
    있었다면 되돌릴 수 없다는 점을 서버가 다시 한번 검증한다. 조건이 맞지 않으면 서버가
    거부 사유를 알려주므로, 그 사유를 사용자에게 그대로 전달하면 된다.

[응답 형식 — 입고]
설명 후 마지막에:
~~~INVENTORY_UPDATE
{"action": "입고", "items": [{"모델명": "정확한모델명", "수량": 1}]}
~~~

[응답 형식 — 출고(설비 확인 완료 후)]
~~~INVENTORY_UPDATE
{"action": "출고", "items": [{"모델명": "정확한모델명", "수량": 1, "실제사용설비": "확인된정확한설비명"}]}
~~~

[응답 형식 — 되돌리기(롤백)]
설명 후 마지막에 (입고/출고 형식과 절대 함께 쓰지 말 것 — 둘 중 하나만):
~~~INVENTORY_ROLLBACK
{"모델명": "정확한모델명", "부품종류": "정확한부품종류(모르면 생략 가능)"}
~~~`;

    // 대화 이력만 contents에, systemInstruction은 별도 파라미터로
    const contents = [];
    if (conversationHistory?.length > 0) {
      conversationHistory.forEach(msg => {
        contents.push({ role: msg.role === 'model' ? 'model' : 'user', parts: [{ text: msg.text }] });
      });
    }
    contents.push({ role: 'user', parts: [{ text: message }] });

    let result;
    try {
      result = await model.generateContent({
        contents,
        systemInstruction: { parts: [{ text: systemInstruction }] }
      });
    } catch (apiError) {
      // Gemini API 호출 자체가 실패한 경우 (네트워크, 429/쿼터, 500 등)
      console.error('❌ Gemini API 호출 실패:', apiError.message);
      console.error('❌ Gemini API 에러 상세:', JSON.stringify(apiError?.response?.data || apiError?.errorDetails || apiError, null, 2).slice(0, 2000));
      return res.status(200).json({
        success: true,
        message: '⚠️ AI 응답을 받아오는 중 문제가 발생했습니다. 잠시 후 다시 시도해 주세요.',
        inventoryUpdated: false,
        updateResult: null,
        timestamp: new Date().toISOString()
      });
    }

    // 응답이 안전 필터(SAFETY/RECITATION 등)에 의해 차단됐는지 먼저 확인
    const candidate = result?.response?.candidates?.[0];
    const finishReason = candidate?.finishReason;
    if (!candidate || (finishReason && finishReason !== 'STOP' && finishReason !== 'MAX_TOKENS')) {
      console.error(`🚫 Gemini 응답 차단됨 — finishReason: ${finishReason}`);
      console.error('🚫 promptFeedback:', JSON.stringify(result?.response?.promptFeedback || {}, null, 2));
      return res.status(200).json({
        success: true,
        message: '⚠️ AI가 이번 요청에는 답변을 생성하지 못했습니다. 문장을 조금 다르게 바꿔서(예: 부품명과 설비명을 한 문장에 같이) 다시 말씀해 주시겠어요?',
        inventoryUpdated: false,
        updateResult: null,
        timestamp: new Date().toISOString()
      });
    }

    let responseText = result.response.text();
    let inventoryUpdated = false;
    let updateResult = null;

    if (responseText.includes('~~~INVENTORY_ROLLBACK')) {
      // ── AI챗봇을 통한 되돌리기(롤백) ──
      // ⚠️ 여기서도 반대 방향 거래를 새로 쌓지 않고, 수동 화면과 동일한
      //    rollbackLogEntry()를 그대로 재사용해 "원본 이력 삭제" 방식으로 처리한다.
      try {
        const parts = responseText.split('~~~INVENTORY_ROLLBACK');
        let jsonPart = parts[1].split('~~~')[0].trim();
        jsonPart = jsonPart.replace(/```json|```/g, '');
        const { 모델명, 부품종류 } = JSON.parse(jsonPart);

        const normModel = String(모델명 || '').replace(/\s+/g, '').toLowerCase();
        const normType = String(부품종류 || '').replace(/\s+/g, '').toLowerCase();

        // facilityLogs는 항상 최신순 정렬 → find로 찾으면 자동으로 "가장 최근 이력"이 잡힌다.
        const targetLog = facilityLogs.find(l => {
          const matchModel = String(l.모델명 || '').replace(/\s+/g, '').toLowerCase() === normModel;
          if (!matchModel) return false;
          if (!부품종류) return true;
          return String(l.부품종류 || '').replace(/\s+/g, '').toLowerCase() === normType;
        });

        if (!targetLog) {
          responseText = `⚠️ "${모델명}"에 해당하는, 되돌릴 수 있는 이력을 찾을 수 없습니다. 정확한 모델명을 다시 확인해 주시겠어요?`;
        } else {
          const result2 = rollbackLogEntry(targetLog, inventoryData, user || 'AI 어시스턴트');
          if (!result2.success) {
            responseText = `⚠️ ${result2.message}`;
          } else {
            const success = await updateExcelOnOneDrive(inventoryData);
            if (success) {
              inventoryUpdated = true;
              updateResult = { success: true, action: 'ROLLBACK', item: result2.item };
              checkAndNotifyLowStock(inventoryData);
              responseText = `↩️ "${result2.item.모델명}" 이력(${targetLog.action})을 되돌리고 사용내역에서 삭제했습니다. (재고 ${result2.oldQuantity}개 → ${result2.restoredQty}개)`;
            } else {
              responseText = '⚠️ 되돌리기 처리 중 OneDrive 업데이트에 실패했습니다.';
            }
          }
        }
      } catch (error) {
        console.error('❌ AI 롤백 명령 처리 오류:', error.message);
        responseText = '⚠️ 되돌리기 요청을 처리하는 중 오류가 발생했습니다.';
      }
    } else if (responseText.includes('~~~INVENTORY_UPDATE')) {
      try {
        const parts = responseText.split('~~~INVENTORY_UPDATE');
        let jsonPart = parts[1].split('~~~')[0].trim();
        jsonPart = jsonPart.replace(/```json|```/g, '');

        const updateData = JSON.parse(jsonPart);
        const { action, items } = updateData;

        // 모델명(공백/대소문자 무시) 기준으로 후보 항목 찾는 헬퍼
        const findCandidates = (모델명) => {
          const normModel = String(모델명 || '').replace(/\s+/g, '').toLowerCase();
          return inventoryData.filter(d => String(d.모델명 || '').replace(/\s+/g, '').toLowerCase() === normModel);
        };

        // ── 백엔드 안전망: 출고(사용)는 항상 실제사용설비가 필수. 입고는 면제 ──
        {
          const missingFacilityItems = items.filter(item => {
            if (action !== '출고') return false; // 입고는 설비 확인 불필요
            return !(item.실제사용설비 && item.실제사용설비.trim());
          });

          if (missingFacilityItems.length > 0) {
            const modelNames = missingFacilityItems.map(i => i.모델명).join(', ');
            const clarifyMsg = `⚠️ 어느 설비에서 사용하신 건지 확인이 필요합니다. (${modelNames})\n정확한 설비명을 말씀해 주세요.`;
            console.log(`🚫 설비 미확인 차단: ${modelNames}`);
            return res.json({ success: true, message: clarifyMsg, inventoryUpdated: false, updateResult: null, timestamp: new Date().toISOString() });
          }
        }

        // 출고 시 실제사용설비가 전체 설비 목록(충전+타정)에 있는 정확한 표기인지 검증
        // ⚠️ 수정: Gemini가 사용자 발화를 그대로 옮겨 적어 공백/전각문자 등이 섞여 들어올 수 있으므로,
        //    facilityListCache 생성 때와 동일한 normalizeEquipment()로 정리한 뒤 비교한다.
        //    (이전엔 trim()만 해서 비교 → 정상 설비명도 "등록된 설비명이 아닙니다"로 계속 반려되는
        //     무한 재확인 루프의 원인이 되었음)
        if (action === '출고') {
          items.forEach(item => {
            if (item.실제사용설비) item.실제사용설비 = normalizeEquipment(item.실제사용설비);
          });
          const allUnitNames = new Set(facilityListCache?.all || []);
          const invalidFacilityItems = items.filter(item => {
            const facRaw = String(item.실제사용설비 || '').trim();
            return facRaw && !allUnitNames.has(facRaw);
          });
          if (invalidFacilityItems.length > 0) {
            const lines = invalidFacilityItems.map(i => `· ${i.모델명} → "${i.실제사용설비}"는 등록된 설비명이 아닙니다.`);
            const clarifyMsg = `⚠️ 설비명을 정확히 확인해 주세요.\n${lines.join('\n')}`;
            console.log(`🚫 설비명 불일치 차단: ${invalidFacilityItems.map(i => i.실제사용설비).join(', ')}`);
            return res.json({ success: true, message: clarifyMsg, inventoryUpdated: false, updateResult: null, timestamp: new Date().toISOString() });
          }
        }

        let updatedCount = 0;
        for (const item of items) {
          const candidates = findCandidates(item.모델명);

          let targetItem = null;
          if (candidates.length === 1) {
            targetItem = candidates[0];
          } else if (candidates.length > 1) {
            // 모델명이 여러 행에 걸쳐 등록된 경우, 어느 행이든 재고 관리 대상은 동일하므로 첫 번째 행 사용
            // (모델명 자체가 유일 키가 아니라 "적용설비" 원본 텍스트만 다른 레거시 중복 행 대비)
            targetItem = candidates[0];
          }

          if (targetItem) {
            updatedCount++;
            const changeQty = Number(item.수량) || 0;
            const finalChange = action === '출고' ? -changeQty : changeQty;
            targetItem.현재수량 = action === '출고' ? Math.max(0, targetItem.현재수량 - changeQty) : targetItem.현재수량 + changeQty;
            targetItem.최종수정시각 = getKSTDate();
            targetItem.작업자 = user || 'AI 어시스턴트';

            // 출고는 실제사용설비를 이력에 확정 기록. 입고는 설비 구분 없이 기록.
            const logItem = { ...targetItem };
            if (action === '출고' && item.실제사용설비) {
              logItem.적용설비 = item.실제사용설비;
              logItem.표준설비명 = item.실제사용설비;
              console.log(`🏭 실제 사용 설비: ${item.실제사용설비}`);
            }

            const sharedLogId = uuidv4(); // 사용내역종합/설비이력 두 저장소를 같은 id로 연결 (되돌리기용)
            addLog(action, logItem, finalChange, user || 'AI 어시스턴트', sharedLogId);
            addFacilityLog(action, logItem, finalChange, user || 'AI 어시스턴트', sharedLogId);
          }
        }

        // ✨ 실제로 변경된 항목이 하나도 없으면 굳이 원본 파일을 다시 쓰지 않는다.
        //    (이전에는 무조건 updateExcelOnOneDrive를 호출했는데, 만약 inventoryData가
        //     어떤 이유로 비정상적이었다면 이 지점에서 불필요하게 위험한 쓰기가 발생할 수 있었다)
        if (updatedCount === 0) {
          responseText = `⚠️ 요청하신 부품을 재고에서 찾지 못해 반영되지 않았습니다. 모델명을 다시 확인해 주세요.`;
        } else {
          const success = await updateExcelOnOneDrive(inventoryData);
          if (success) {
            inventoryUpdated = true;
            updateResult = { success: true, action, items };
            checkAndNotifyLowStock(inventoryData); // Teams 저재고 알림
          }
        }
      } catch (error) {
        console.error('❌ AI 명령 처리 오류:', error.message);
      }
    }

    res.json({ success: true, message: responseText, inventoryUpdated, updateResult, timestamp: new Date().toISOString() });
  } catch (error) {
    console.error('❌ AI 채팅 에러:', error.message);
    res.status(500).json({ success: false, message: 'AI 응답 오류' });
  }
});

// ============================================================
// Device Code Flow
// ============================================================
async function getTokenViaDeviceFlow() {
  try {
    console.log('\n📱 Device Code Flow 시작...\n');
    const deviceCodeResponse = await axios.post(
      'https://login.microsoftonline.com/common/oauth2/v2.0/devicecode',
      new URLSearchParams({
        client_id: CONFIG.clientId,
        scope: 'Files.ReadWrite Files.ReadWrite.All offline_access'
      }),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    const { user_code, device_code, verification_uri, expires_in, interval } = deviceCodeResponse.data;
    console.log('====================================================');
    console.log(`1. 브라우저에서 접속: ${verification_uri}`);
    console.log(`2. 코드 입력: ${user_code}`);
    console.log(`3. Microsoft 계정으로 로그인`);
    console.log('====================================================\n대기 중');

    const pollInterval = (interval || 5) * 1000;
    const maxAttempts = Math.floor(expires_in / (interval || 5));

    for (let i = 0; i < maxAttempts; i++) {
      await new Promise(resolve => setTimeout(resolve, pollInterval));
      try {
        const tokenResponse = await axios.post(
          'https://login.microsoftonline.com/common/oauth2/v2.0/token',
          new URLSearchParams({
            client_id: CONFIG.clientId,
            grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
            device_code: device_code
          }),
          { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
        );

        const tokens = {
          access_token: tokenResponse.data.access_token,
          refresh_token: tokenResponse.data.refresh_token,
          expires_at: Date.now() + (tokenResponse.data.expires_in * 1000)
        };
        saveTokens(tokens);
        console.log('\n✅ 인증 성공!');
        return tokens;
      } catch (error) {
        if (error.response?.data?.error === 'authorization_pending') {
          process.stdout.write('.');
        } else {
          throw error;
        }
      }
    }
    return null;
  } catch (error) {
    console.error('❌ Device Flow 실패:', error.response?.data || error.message);
    return null;
  }
}

// ============================================================
// ✨ 원격 재인증(Device Flow) — node 명령어를 로컬에서 실행할 수 없는 환경에서도
// 브라우저 URL만으로 새 refresh_token(Mail.Send 권한 포함)을 재발급받을 수 있게
// 두 단계로 분리했다. 서버(이미 배포되어 Node로 돌아가는 쪽)가 폴링을 대신 수행하므로
// 회사 PC의 node 실행 제한과 무관하다.
// ⚠️ 이 엔드포인트들은 민감한 인증 절차를 다루므로, 재발급 작업이 끝나면 코드에서
//    제거하거나 접근을 제한하는 것을 권장한다.
// ============================================================
let deviceAuthState = null; // { status, userCode, verificationUri, expiresAt, resultToken, errorMessage }

app.get('/api/admin/device-auth/start', async (req, res) => {
  try {
    const deviceCodeResponse = await axios.post(
      'https://login.microsoftonline.com/common/oauth2/v2.0/devicecode',
      new URLSearchParams({
        client_id: CONFIG.clientId,
        // ✨ Mail.Send 추가 — 이메일 인증코드 발송 기능에 필요
        scope: 'Files.ReadWrite Files.ReadWrite.All Mail.Send offline_access'
      }),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    const { user_code, device_code, verification_uri, expires_in, interval } = deviceCodeResponse.data;
    deviceAuthState = {
      status: 'pending',
      userCode: user_code,
      verificationUri: verification_uri,
      expiresAt: Date.now() + expires_in * 1000,
      resultToken: null,
      errorMessage: null
    };

    // 백그라운드에서 폴링 — 이 HTTP 요청과 무관하게 서버 프로세스 안에서 계속 진행됨
    const pollMs = (interval || 5) * 1000;
    const pollTimer = setInterval(async () => {
      if (!deviceAuthState || deviceAuthState.status !== 'pending') {
        clearInterval(pollTimer);
        return;
      }
      if (Date.now() > deviceAuthState.expiresAt) {
        deviceAuthState.status = 'expired';
        clearInterval(pollTimer);
        return;
      }
      try {
        const tokenResponse = await axios.post(
          'https://login.microsoftonline.com/common/oauth2/v2.0/token',
          new URLSearchParams({
            client_id: CONFIG.clientId,
            grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
            device_code: device_code
          }),
          { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
        );
        const tokens = {
          access_token: tokenResponse.data.access_token,
          refresh_token: tokenResponse.data.refresh_token,
          expires_at: Date.now() + tokenResponse.data.expires_in * 1000
        };
        saveTokens(tokens);
        deviceAuthState.status = 'success';
        deviceAuthState.resultToken = tokens.refresh_token;
        console.log('✅ 원격 재인증 성공 — 새 refresh_token 발급됨 (환경변수에 반영 필요)');
        clearInterval(pollTimer);
      } catch (error) {
        if (error.response?.data?.error !== 'authorization_pending') {
          deviceAuthState.status = 'error';
          deviceAuthState.errorMessage = error.response?.data?.error_description || error.message;
          clearInterval(pollTimer);
        }
        // authorization_pending이면 계속 대기 (다음 인터벌에 재시도)
      }
    }, pollMs);

    res.json({
      success: true,
      message: '아래 주소로 접속해서 코드를 입력하세요. 완료 후 /api/admin/device-auth/status 에서 결과를 확인할 수 있습니다.',
      verification_uri,
      user_code,
      expires_in_seconds: expires_in
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.response?.data?.error_description || error.message });
  }
});

app.get('/api/admin/device-auth/status', (req, res) => {
  if (!deviceAuthState) {
    return res.json({ success: false, message: '진행 중인 인증이 없습니다. /start를 먼저 호출하세요.' });
  }
  res.json({ success: true, ...deviceAuthState });
});

app.listen(PORT, () => {
  console.log(`\n🚀 백엔드 서버 실행 중: http://localhost:${PORT}`);
  console.log(`📁 OneDrive 파일: ${CONFIG.excelFileName}`);

  if (process.env.REFRESH_TOKEN) {
    console.log('✅ REFRESH_TOKEN 환경변수 감지됨 - OneDrive 연동 준비 완료');
  } else {
    console.log('⚠️ REFRESH_TOKEN 없음 - 로컬에서 get-token.js를 먼저 실행하세요');
  }

  // ✨ 캐시 예열(warm-up) — 캐시가 만료되기 전에 백그라운드에서 미리 갱신해둔다.
  //    이렇게 하면 사용자가 접속했을 때 "하필 캐시가 막 만료된 시점이라 30초씩 걸리는"
  //    상황이 거의 사라진다. (서버 자체가 잠들지 않는 건 별도 외부 핑으로 이미 해결했으므로,
  //    서버가 살아있는 동안 이 인터벌이 캐시를 계속 따뜻하게 유지해준다.)
  //    CACHE_DURATION(5분)보다 짧은 4분 주기로 돌려서, 캐시가 실제로 만료되기 전에 항상 갱신됨.
  setInterval(async () => {
    try {
      const before = Date.now();
      await fetchExcelFromOneDrive();
      console.log(`🔄 캐시 예열 완료 (${Date.now() - before}ms)`);
    } catch (err) {
      console.error('⚠️ 캐시 예열 실패 (다음 실제 요청 때 재시도됨):', err.message);
    }
  }, 4 * 60 * 1000);

  if (CONFIG.teamsWebhookUrl) {
    console.log('✅ Teams Webhook 설정됨 - 재고 부족 알림 활성화');

    // 매일 오전 9시 (KST) 정기 재고 부족 알림
    // alertCooldown을 초기화하여 정기 체크는 항상 발송되도록 함
    setInterval(async () => {
      const kstHour = new Date(Date.now() + 9 * 60 * 60 * 1000).getUTCHours();
      const kstMin  = new Date(Date.now() + 9 * 60 * 60 * 1000).getUTCMinutes();
      if (kstHour === 9 && kstMin < 5) {
        console.log('⏰ 오전 9시 정기 재고 체크 실행');
        alertCooldown.clear(); // 정기 체크는 쿨다운 무시하고 전체 발송
        try {
          const data = await fetchExcelFromOneDrive();
          const lowStock = data.filter(d => d.최소보유수량 > 0 && d.현재수량 < d.최소보유수량);
          if (lowStock.length > 0) {
            await sendTeamsAlert(lowStock);
          } else {
            console.log('✅ 정기 체크: 저재고 항목 없음');
          }
        } catch (err) {
          console.error('❌ 정기 재고 체크 오류:', err.message);
        }
      }
    }, 5 * 60 * 1000); // 5분마다 시각 확인

  } else {
    console.log('⚠️ TEAMS_WEBHOOK_URL 미설정 - Teams 알림 비활성화');
  }
});
