// Modules to control application life and create native browser window
const {app, BrowserWindow, ipcMain, dialog} = require('electron');
const path = require('path');
const xlsx = require('xlsx');
const xlsxs = require('xlsx-style');
const moment = require('moment');
const _ = require('lodash');

function createWindow () {
  const mainWindow = new BrowserWindow({
    width: 600,
    height: 300,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
    }
  });
  mainWindow.loadFile('index.html');
  // mainWindow.webContents.openDevTools();
};

app.whenReady().then(() => {
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
});

app.on('window-all-closed', () => app.quit());

ipcMain.on('file', event => {
  selectFile('openFile', file => {
    if (file !== undefined)
      event.reply('file', file);
  });
});

ipcMain.on('dir', event => {
  selectFile('openDirectory', file => {
    if (file !== undefined)
      event.reply('dir', file);
  });
});

ipcMain.on('run', (event, arg) => {
  const { file, dir, year, to } = arg;
  run(file, dir, year, to);
  event.reply('run');
});

function run(file, dir, year, capacity) {
  clearErrors();

  const students = getRows(file).map(o => new Student(o));
  const [healthy, _unhealthy] = _.partition(students, 'healthy');
  const [_medical, _normal] = _.partition(healthy, 'medical');

  var unhealthy = _unhealthy;
  var normal = _normal;
  var medical = _medical;
  var rank = 1;
  var arr = [];
  while (true) {
    const [picked, _normal, _medical] = pick(normal, medical, year, capacity);
    assertInternal(picked.length + _normal.length + _medical.length === normal.length + medical.length, 'wrong sum');
    picked.map((s, i) => s.setRank(rank + i));

    const [pu1, ru1] = _.partition(unhealthy, s => s.semester(year) >= 3);
    for (const s of pu1) s.pick(moment([year, 2, 1]), '보충역');
    const [pu2, ru] = _.partition(ru1, s => s.semester(year) === 2);
    for (const s of pu2) s.pick(moment([year, 8, 1]), '보충역');

    arr = _.concat(arr, pu1, picked, pu2);

    if (picked.length < capacity) {
      assertInternal(_normal.length + _medical.length === 0, 'should be empty');
      break;
    }

    year++;
    unhealthy = ru;
    normal = _normal;
    medical = _medical;
    rank += capacity;
  }

  const header = [
    '학번',
    '신체등급',
    '생년월일',
    '박사과정진입일',
    '의무사관후보생',
    '편입예상시점',
    '편입시점학기',
    '편입시점나이',
    '우선순위',
    '비고'
  ];
  const body = arr.map(s => s.toArray());
  const ws = arrayToSheet(_.concat([header], body));
  const wb = sheetsToBook([['대상자', ws]]);
  xlsxs.writeFile(wb, path.join(dir, `편입예상시점.xlsx`), { bookType: 'xlsx' });

  showErrors();
}

function pick(normal, medical, year, capacity) {
  const date = () => moment([year, 2, 1]);

  const [sMedical, jMedical] = _.partition(medical, s => s.semester(year) >= 3);
  for (const s of sMedical) {
    const semester = s.semester(year);
    assert(semester === 3, `${s.id}: 의무사관후보생 ${semester}학기`);
    s.pick(date(), '의무사관후보생');
  }

  const [old, young] = _.partition(normal, s => s.age(year) >= 29);
  for (const s of old) {
    const age = s.age(year);
    assert(age === 29, `${s.id}: ${age}세`);
    s.pick(date(), '29세');
  }

  const [senior, _junior] = _.partition(young, s => s.semester(year) >= 6);
  for (const s of senior) {
    const semester = s.semester(year);
    assert(semester <= 7, `${s.id}: ${semester}학기`);
    s.pick(date(), '7학기');
  }

  let remaining = capacity - sMedical.length - old.length - senior.length;
  assert(remaining >= 0, `${year} 정원 부족: ${capacity} < ${sMedical.length} + ${old.length} + ${senior.length}`);
  if (remaining < 0) remaining = 0;

  const junior = _.sortBy(_junior, ['start', 'birth']);
  const pickedJunior = _.take(junior, remaining);
  for (const s of pickedJunior) s.pick(date(), '');
  const remainingJunior = _.drop(junior, remaining);

  return [_.concat(sMedical, old, senior, pickedJunior), remainingJunior, jMedical];
}

function checkTime(t) {
  const h = t.hour();
  const m = t.minute();
  const s = t.second();
  assertInternal(h === 23, `wrong hour ${h}`);
  assertInternal(m === 59, `wrong minute ${m}`);
  assertInternal(s === 8, `wrong second ${s}`);
}

function makeStart(id, t) {
  let m = t.month() + 1;
  switch (m) {
    case 2:
    case 3:
      m = 2;
      break;
    case 8:
    case 9:
      m = 8;
      break;
    default:
      assert(false, `${id}: ${m}월 진입`);
      if (m <= 6)
        m = 2;
      else
        m = 8;
  }
  return moment([t.year(), m, 1]);
}

class Student {
  constructor(obj) {
    this.id = obj['학번'];
    this.healthy = obj['신체등급'] === '현역';
    this.medical = obj['의무사관후보생'] === "O";

    this.birth = moment(obj['생년월일']);
    checkTime(this.birth);
    this.birth.add(52, 's');

    this._start = moment(obj['박사과정진입일']);
    checkTime(this._start);
    this._start.add(52, 's');
    this.start = makeStart(this.id, this._start);
  }

  age(year) {
    return year - this.birth.year();
  }

  semester(year) {
    return (year - this.start.year()) * 2 + (this.start.month() === 8 ? 0 : 1);
  }

  pick(date, reason) {
    this.milStart = date;
    this.reason = reason;
    this.rank = 0;
    this.milSemester = (date.year() - this.start.year()) * 2 + (date.month() - this.start.month()) / 6 + 1;
    this.milAge = date.year() - this.birth.year();
  }

  setRank(rank) {
    this.rank = rank;
  }

  toArray() {
    return [
      this.id, 
      this.healthy ? '현역' : '보충역',
      this.birth.add(9, 'h').toDate(),
      this._start.add(9, 'h').toDate(),
      this.medical ? 'O' : '',
      this.milStart.add(9, 'h').toDate(),
      this.milSemester,
      this.milAge,
      this.rank,
      this.reason
    ];
  }
}

function assertInternal(b, msg) {
  if (!b) throw new Error(msg);
}

var errors = [];

function assert(b, msg) {
  if (!b) errors.push(msg);
}

function clearErrors() {
  errors = [];
}

function showErrors() {
  if (errors.length > 0) {
    errors.sort();
    dialog.showErrorBox("오류", errors.join('\n'));
  }
}

function selectFile(prop, cb) {
  dialog.showOpenDialog({ properties: [prop] })
  .then(x => cb(x.filePaths[0]))
  .catch(e => console.log(e));
}

function getRows(fn) {
  const wb = xlsx.readFile(fn, { cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return xlsx.utils.sheet_to_json(ws);
}

function widthOf(v) {
  const str = typeof(v) === 'object' ? moment(v).format('YYYY.M.D') : v.toString();
  let w = 0;
  for (const i in str) {
    let code = str.charCodeAt(i);
    w += (0xac00 <= code && code <= 0xd7af) ? 2 : 1;
  }
  return w;
}

function arrayToSheet(arr) {
  const ws = {};
  const R = arr.length;
  const C = arr.map(a => a.length).reduce((a, b) => (a > b) ? a : b, 0);
  const range = { s: { c: 0, r: 0 }, e: { c: C - 1, r: R - 1 } };
  const widths = Array(C);
  widths.fill(10);

  for (let r = 0; r < R; r++) {
    for (let c = 0; c < arr[r].length; c++) {
      const cell = { v: arr[r][c] };
      cell.t =
        (typeof(cell.v) === 'number') ? 'n' :
        (typeof(cell.v) === 'string') ? 's' :
        'd';
      if (r === 0)
        cell.s = { font: { bold: true } };
      else
        cell.s = {};
      if (r === 0) {
        cell.s.border = {
          top: { style: "thin", color: { rgb: "FF000000" } },
          bottom: { style: "thin", color: { rgb: "FF000000" } },
          left: { style: "thin", color: { rgb: "FF000000" } },
          right: { style: "thin", color: { rgb: "FF000000" } }
        };
      }
      const w = widthOf(cell.v);
      if (w > widths[c]) widths[c] = w;
      const ref = xlsxs.utils.encode_cell({ r, c });
      ws[ref] = cell;
    }
  }
  ws['!ref'] = xlsxs.utils.encode_range(range);
  ws['!cols'] = widths.map(w => { return { wch: w }; });
  return ws;
}

function sheetsToBook(sheets) {
  const wb = { SheetNames: [], Sheets: {} };
  for (const [name, sheet] of sheets) {
    wb.SheetNames.push(name);
    wb.Sheets[name] = sheet;
  }
  return wb;
}

// run('/Users/medowhill/Downloads/mil.xlsx', '/Users/medowhill/Downloads');
