// All of the Node.js APIs are available in the preload process.
// It has the same sandbox as a Chrome extension.
//
const { ipcRenderer } = require('electron');

window.addEventListener('DOMContentLoaded', () => {
  onClick('button-file', () => ipcRenderer.send('file'));
  onClick('button-dir', () => ipcRenderer.send('dir'));
  onClick('button-run', () => {
    const year = Number.parseInt(document.getElementById('input-year').value);
    const to = Number.parseInt(document.getElementById('input-to').value);
    const file = document.getElementById('p-file').innerText;
    const dir = document.getElementById('p-dir').innerText;
    if (!isNaN(year) && !isNaN(to) && file !== '선택된 파일 없음' && dir !== '선택된 폴더 없음') {
      document.getElementById('p-complete').style = 'display: none;'
      ipcRenderer.send('run', { year, to, file, dir });
    }
  });
});

onMessage('file', arg => document.getElementById('p-file').innerText = arg);
onMessage('dir', arg => document.getElementById('p-dir').innerText = arg);
onMessage('run', () => document.getElementById('p-complete').style = 'display: inline-block;');

function onClick(b, cb) {
  document.getElementById(b).addEventListener('click', cb);
}

function onMessage(m, cb) {
  ipcRenderer.on(m, (event, arg) => cb(arg));
}
