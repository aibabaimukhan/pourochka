// Вставьте сюда URL веб-приложения Apps Script (Deploy → Web app → URL, заканчивается на /exec)
window.API_URL = 'https://script.google.com/macros/s/AKfycbyV63LdU16_pAz3ANBPAk7uaMRkVjgG-vc_Lz7gxbbHuSv6bu-muHinJ7EJnO2wnjmW/exec';
window.api = async function (payload) {
  const r = await fetch(window.API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(payload) });
  return r.json();
};
