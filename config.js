// ★ Apps Script 웹 앱 URL — 배포 후 나오는 주소(…/exec)를 붙여넣으세요
const GAS_URL = 'https://script.google.com/macros/s/AKfycbzpZCUULOY0IBtBvshfsgNmJZLZhX1ai5Tnee9Mod5uiKi0ATkoumUYnR4jP2oQJLpHPA/exec';

// 사이트 상단에 표시할 학교 이름
const SCHOOL_NAME = '영도초등학교 생활안전부';

// 파일 1개 최대 크기(MB). Apps Script 한계 때문에 25MB 이하를 권장합니다.
const MAX_FILE_MB = 25;

// 같은 요청을 다시 보내도 결과가 같은 것만 자동 재시도 (업로드·저장은 중복될 수 있어 제외)
const RETRY_ACTIONS = ['list', 'check', 'adminLogin', 'adminList', 'adminDetail', 'getRoster', 'getFile', 'zip'];

async function api(action, payload = {}) {
  if (!GAS_URL) throw new Error('config.js에 GAS_URL이 설정되지 않았습니다.');
  const tries = RETRY_ACTIONS.includes(action) ? 3 : 1;
  let lastError;
  for (let i = 0; i < tries; i++) {
    if (i) await new Promise(r => setTimeout(r, 1500 * i));
    try {
      // Content-Type을 지정하지 않아야(text/plain) Apps Script가 CORS 오류 없이 받습니다.
      const res = await fetch(GAS_URL, { method: 'POST', body: JSON.stringify({ action, ...payload }) });
      const text = await res.text();
      let data;
      try { data = JSON.parse(text); }
      catch {
        // 구글이 JSON 대신 오류 페이지(HTML)를 보낸 경우 — 페이지 제목을 같이 보여줌
        const title = (text.match(/<title>([^<]*)<\/title>/i) || [])[1] || `HTTP ${res.status}`;
        throw new Error(`구글 서버가 일시적으로 응답하지 않습니다 (${title.trim()}). 잠시 후 다시 시도해주세요.`);
      }
      // 서버가 알려준 오류(비밀번호 틀림 등)는 다시 시도하지 않고 바로 돌려줌
      if (!data.ok) return Promise.reject(new Error(data.error || '요청에 실패했습니다.'));
      return data;
    } catch (e) {
      // fetch 자체가 실패하면 TypeError (브라우저마다 문구가 다름)
      lastError = e instanceof TypeError ? new Error('인터넷 연결 또는 학교 네트워크에서 구글 서버에 접속하지 못했습니다.') : e;
    }
  }
  throw lastError;
}

function esc(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

/** 'yyyy-MM-dd' → 'D-3' / 'D-Day' / '마감일 지남' */
function dday(deadline) {
  if (!deadline) return '';
  const end = new Date(deadline + 'T00:00:00');
  if (isNaN(end)) return '';
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const diff = Math.round((end - today) / 86400000);
  if (diff > 0) return 'D-' + diff;
  if (diff === 0) return 'D-Day';
  return '마감일 지남';
}

function formatDate(deadline) {
  if (!deadline) return '';
  const d = new Date(deadline + 'T00:00:00');
  if (isNaN(d)) return deadline;
  return `${d.getMonth() + 1}월 ${d.getDate()}일(${'일월화수목금토'[d.getDay()]})`;
}
