// ★ Apps Script 웹 앱 URL — 배포 후 나오는 주소(…/exec)를 붙여넣으세요
const GAS_URL = 'https://script.google.com/macros/s/AKfycbzpZCUULOY0IBtBvshfsgNmJZLZhX1ai5Tnee9Mod5uiKi0ATkoumUYnR4jP2oQJLpHPA/exec';

// 사이트 상단에 표시할 학교 이름
const SCHOOL_NAME = '영도초등학교 생활안전부';

// 파일 1개 최대 크기(MB). Apps Script 한계 때문에 25MB 이하를 권장합니다.
const MAX_FILE_MB = 25;

async function api(action, payload = {}) {
  if (!GAS_URL) throw new Error('config.js에 GAS_URL이 설정되지 않았습니다.');
  // Content-Type을 지정하지 않아야(text/plain) Apps Script가 CORS 오류 없이 받습니다.
  const res = await fetch(GAS_URL, { method: 'POST', body: JSON.stringify({ action, ...payload }) });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || '요청에 실패했습니다.');
  return data;
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
