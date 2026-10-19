// PDF 합치기 — pdf-lib(https://pdf-lib.js.org)를 필요할 때만 불러옵니다.
const PDF_LIB_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js';

const A4 = [595.28, 841.89];

function loadPdfLib() {
  if (window.PDFLib) return Promise.resolve(window.PDFLib);
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = PDF_LIB_URL;
    s.onload = () => resolve(window.PDFLib);
    s.onerror = () => reject(new Error('PDF 도구를 불러오지 못했습니다. 인터넷 연결을 확인해주세요.'));
    document.head.appendChild(s);
  });
}

/** 합칠 수 있는 파일 종류: 'pdf' | 'jpg' | 'png' | null */
function mergeKind(name) {
  const ext = (String(name).match(/\.([^.]+)$/) || ['', ''])[1].toLowerCase();
  if (ext === 'pdf') return 'pdf';
  if (ext === 'jpg' || ext === 'jpeg') return 'jpg';
  if (ext === 'png') return 'png';
  return null;
}

/**
 * items 순서대로 하나의 PDF로 합칩니다. 사진은 A4 한 장에 맞춰 넣습니다.
 * items: [{ name, kind, bytes: Uint8Array }]
 * 반환: { bytes, pages, failed: [이름], warned: [이름] }
 */
async function mergeToPdf(PDFLib, items) {
  const out = await PDFLib.PDFDocument.create();
  const failed = [];
  const warned = [];

  for (const it of items) {
    try {
      if (it.kind === 'pdf') {
        const src = await PDFLib.PDFDocument.load(it.bytes, { ignoreEncryption: true });
        if (src.isEncrypted) warned.push(it.name);
        const pages = await out.copyPages(src, src.getPageIndices());
        pages.forEach(p => out.addPage(p));
      } else {
        let img;
        try { img = it.kind === 'png' ? await out.embedPng(it.bytes) : await out.embedJpg(it.bytes); }
        catch (e) { img = it.kind === 'png' ? await out.embedJpg(it.bytes) : await out.embedPng(it.bytes); } // 확장자와 실제 형식이 다른 경우
        const size = img.width > img.height ? [A4[1], A4[0]] : A4; // 가로 사진은 가로 페이지
        const margin = 28;
        const scale = Math.min((size[0] - margin * 2) / img.width, (size[1] - margin * 2) / img.height);
        const w = img.width * scale;
        const h = img.height * scale;
        out.addPage(size).drawImage(img, { x: (size[0] - w) / 2, y: (size[1] - h) / 2, width: w, height: h });
      }
    } catch (e) {
      failed.push(it.name);
    }
  }

  return { bytes: await out.save(), pages: out.getPageCount(), failed, warned };
}

function base64ToBytes(b64) {
  return Uint8Array.from(atob(b64), ch => ch.charCodeAt(0));
}
