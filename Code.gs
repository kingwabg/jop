// ================================================================
//  운영일지 관리 - Google Apps Script (Code.gs)
//  [아키텍처]
//  - 운영일지 : 화면 표시용 시트 1개 (항상 고정)
//  - __데이터__: 숨김 데이터 저장 시트 (일지마다 한 행)
//  - 사이드바 목록 클릭 → 해당 데이터를 운영일지 시트에 로드
// ================================================================

var DISP  = '운영일지';   // 표시 시트 이름 (고정)
var STORE = '__데이터__'; // 데이터 저장 시트
var TMPL  = '__템플릿__'; // 템플릿 시트 (사용자 편집용)

// ── 메뉴 ──
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 운영일지')
    .addItem('사이드바 열기', 'openSidebar')
    .addSeparator()
    .addItem('📝 템플릿 편집하기', 'editTemplateMenu')
    .addItem('새 일지 만들기 (템플릿 기반)', 'createNewJournal')
    .addItem('현재 내용 저장', 'saveCurrentDataMenu')
    .addItem('인쇄 설정 적용', 'applyPrintSettingsMenu')
    .addToUi();
}

function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('📋 운영일지').setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ══════════════════════════════════════════════════════════════
//  내부 헬퍼
// ══════════════════════════════════════════════════════════════

function getDispSheet(ss) {
  return ss.getSheetByName(DISP) || ss.insertSheet(DISP);
}

function getStoreSheet(ss) {
  var s = ss.getSheetByName(STORE);
  if (!s) {
    s = ss.insertSheet(STORE);
    s.hideSheet();
    s.appendRow(['title','date','manager','record_type','op_time','saved_at','data_json']);
    s.getRange('A:A').setNumberFormat('@'); // A열(제목)을 강제로 '텍스트' 형식으로 지정
    s.setFrozenRows(1);
  } else {
    // 이미 시트가 있는 경우에도 A열 형식을 텍스트로 보장 (혹시 모를 변환 방지)
    s.getRange('A:A').setNumberFormat('@');
  }
  return s;
}

function findStoreRow(store, title) {
  var data = store.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(title)) return i + 1;
  }
  return -1;
}

function buildList(store) {
  var rows = store.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var list = [];
  for (var i = 1; i < rows.length; i++) {
    var title = rows[i][0], date = rows[i][1], manager = rows[i][2], savedAt = rows[i][5];
    if (!title) continue;
    list.push({
      name:    String(title),
      date:    toDateStr(date),
      savedAt: savedAt ? Utilities.formatDate(new Date(savedAt),'Asia/Seoul','MM/dd HH:mm') : '',
      manager: String(manager || '')
    });
  }
  list.sort(function(a,b){ return b.date.localeCompare(a.date); });
  return list;
}

/** 날짜 값을 'yyyy-MM-dd' 문자열로 변환 */
function toDateStr(val) {
  if (!val) return '';
  try {
    var d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return String(val).substring(0,10);
    return Utilities.formatDate(d,'Asia/Seoul','yyyy-MM-dd');
  } catch(e) { return String(val).substring(0,10); }
}

// ── 데이터 셀만 초기화 (템플릿 구조/서식 유지) ──
function clearDataCells(sheet) {
  sheet.getRangeList([
    'B3','F3','J3',
    'C7:H7','C8:H8',
    'B11','D11','F11',
    'B14','D14','F14','H14',
    'A18:G18',
    'B22:F26','G22:G26',
    'A31',
    'A35:A40','C35:C40','I35:I40','M35:M40',
    'A44:A48','D44:D48','K44:K48'
  ]).clearContent();
}

// ── 데이터 → 시트 셀 적용 ──
function applyData(sheet, data, recordLabel) {
  recordLabel = recordLabel || '아동';
  sheet.getRange('A1').setValue('운영일지(' + recordLabel + ')');
  try { sheet.getRange('A28').setValue('운영일지(' + recordLabel + ')  ─  2 페이지'); } catch(e){}

  sheet.getRange('B3').setValue(data.date    || '');
  sheet.getRange('F3').setValue(data.opTime  || '09:00~18:00');
  sheet.getRange('J3').setValue(data.manager || '');

  // 아동 현황 배치 쓰기
  sheet.getRange('C7:H7').setValues([ fillArr(data.male,   6) ]);
  sheet.getRange('C8:H8').setValues([ fillArr(data.female, 6) ]);

  // 급식
  var m = data.meals || [0,0,0];
  sheet.getRange('B11').setValue(m[0]||0);
  sheet.getRange('D11').setValue(m[1]||0);
  sheet.getRange('F11').setValue(m[2]||0);

  // 교사
  var st = data.staff || [0,0,0,0];
  sheet.getRange('B14').setValue(st[0]||0);
  sheet.getRange('D14').setValue(st[1]||0);
  sheet.getRange('F14').setValue(st[2]||0);
  sheet.getRange('H14').setValue(st[3]||0);

  // 출석
  sheet.getRange('A18:G18').setValues([ fillArr(data.att, 7) ]);

  // 프로그램 (배치)
  var progs = data.progs || [];
  var bfArr = [], gArr = [];
  for (var i = 0; i < 5; i++) {
    var p = progs[i] || [];
    bfArr.push([p[0]||'', p[1]||'', p[2]||'', p[3]||'', p[4]||'']);
    gArr.push([p[5]||'']);
  }
  sheet.getRange('B22:F26').setValues(bfArr);
  sheet.getRange('G22:G26').setValues(gArr);

  // 메모
  sheet.getRange('A31').setValue(data.memo || '');

  // 업무일지 (배치)
  var biz = data.biz || [];
  var bizA=[], bizC=[], bizI=[], bizM=[];
  for (var j = 0; j < 6; j++) {
    var b = biz[j] || [];
    bizA.push([b[0]||'']); bizC.push([b[1]||'']);
    bizI.push([b[2]||'']); bizM.push([b[3]||'']);
  }
  sheet.getRange('A35:A40').setValues(bizA);
  sheet.getRange('C35:C40').setValues(bizC);
  sheet.getRange('I35:I40').setValues(bizI);
  sheet.getRange('M35:M40').setValues(bizM);

  // 안전점검 (배치)
  var safety = data.safety || [];
  var saA=[], saD=[], saK=[];
  for (var k = 0; k < 5; k++) {
    var s = safety[k] || [];
    saA.push([s[0]||'']); saD.push([s[1]||'']); saK.push([s[2]||'']);
  }
  sheet.getRange('A44:A48').setValues(saA);
  sheet.getRange('D44:D48').setValues(saD);
  sheet.getRange('K44:K48').setValues(saK);

  SpreadsheetApp.flush();
}

// ── 시트 셀 → 데이터 객체 추출 ──
function extractData(sheet) {
  var info   = sheet.getRange('B3:N3').getValues()[0];
  var male   = sheet.getRange('C7:H7').getValues()[0];
  var female = sheet.getRange('C8:H8').getValues()[0];
  var meal   = sheet.getRange('B11:F11').getValues()[0];
  var tchr   = sheet.getRange('B14:H14').getValues()[0];
  var att    = sheet.getRange('A18:G18').getValues()[0];
  var progsR = sheet.getRange('B22:N26').getValues();
  var memo   = sheet.getRange('A31').getValue();
  var bizR   = sheet.getRange('A35:N40').getValues();
  var sfR    = sheet.getRange('A44:N48').getValues();

  return {
    date:   info[0], opTime: info[4], manager: info[8],
    male:   male.map(function(v){return v||0;}),
    female: female.map(function(v){return v||0;}),
    meals:  [meal[0]||0, meal[2]||0, meal[4]||0],
    staff:  [tchr[0]||0, tchr[2]||0, tchr[4]||0, tchr[6]||0],
    att:    att.map(function(v){return v||0;}),
    progs:  progsR.map(function(r){ return [r[0]||'',r[1]||'',r[2]||'',r[3]||'',r[4]||'',r[5]||'']; }),
    memo:   memo || '',
    biz:    bizR.map(function(r){ return [r[0]||'',r[2]||'',r[8]||'',r[12]||'']; }),
    safety: sfR.map(function(r){ return [r[0]||'',r[3]||'',r[10]||'']; })
  };
}

function fillArr(arr, len) {
  var a = arr || [];
  var result = [];
  for (var i = 0; i < len; i++) result.push(a[i]||0);
  return result;
}

function blankData(date, opTime, manager) {
  return {
    date: date||'', opTime: opTime||'09:00~18:00', manager: manager||'',
    male:[0,0,0,0,0,0], female:[0,0,0,0,0,0],
    meals:[0,0,0], staff:[0,0,0,0], att:[0,0,0,0,0,0,0],
    progs:[[],[],[],[],[]], memo:'', biz:[[],[],[],[],[],[]], safety:[[],[],[],[],[]]
  };
}

// ══════════════════════════════════════════════════════════════
//  사이드바 ↔ 시트 통신 (alert 사용 금지)
// ══════════════════════════════════════════════════════════════

/** 목록 반환 */
function getJournalList() {
  try {
    return buildList(getStoreSheet(SpreadsheetApp.getActiveSpreadsheet()));
  } catch(e) { return []; }
}

/**
 * 새 일지 생성 + 데이터 저장소 등록 + 목록 반환 (1회 호출)
 */
function createAndRegister(formData) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var store = getStoreSheet(ss);
    var disp  = getDispSheet(ss);

    var title  = (formData && formData.title)  || makeDefaultTitle();
    var date   = (formData && (formData.opdate || formData.date)) || '';
    var mgr    = (formData && formData.manager) || '';
    var recType= (formData && formData.record)  || '운영 일지(아동)';
    var recLbl = recType.replace(/운영 일지\(|\)/g,'') || '아동';

    // 템플릿 구조가 없으면 빌드, 있으면 데이터만 초기화
    var hasTemplate = disp.getLastRow() >= 40 &&
      disp.getRange('A1').getValue().toString().indexOf('운영일지') >= 0;
    if (!hasTemplate) {
      buildTemplate(disp, date, mgr, recLbl);
    } else {
      clearDataCells(disp);
    }

    var data = blankData(date, '09:00~18:00', mgr);
    applyData(disp, data, recLbl);
    applyPrintSettings(disp, {top:1.5,bottom:1.5,left:1.5,right:1.5});
    ss.setActiveSheet(disp);

    // 데이터 저장소 등록
    var now  = new Date();
    var json = JSON.stringify(data);
    var foundRow = findStoreRow(store, title);
    if (foundRow > 0) {
      store.getRange(foundRow,1,1,7).setValues([[title,toDateStr(date),mgr,recType,'09:00~18:00',now,json]]);
    } else {
      store.appendRow([title,toDateStr(date),mgr,recType,'09:00~18:00',now,json]);
    }
    // 현재 활성 일지 타이틀을 문서 속성에 저장 (사이드바 재오픈 시 복구용)
    PropertiesService.getDocumentProperties().setProperty('activeTitle', title);

    return { ok:true, name:title, list:buildList(store) };
  } catch(e) { return { ok:false, error:e.message }; }
}

/**
 * 목록에서 일지 클릭 → 운영일지 시트에 데이터 로드
 */
function loadEntry(title) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var store = getStoreSheet(ss);
    var disp  = getDispSheet(ss);

    var foundRow = findStoreRow(store, title);
    if (foundRow < 0) return { ok:false, error:'항목 없음: '+title };

    var row     = store.getRange(foundRow,1,1,7).getValues()[0];
    var recType = String(row[3]||'운영 일지(아동)');
    var recLbl  = recType.replace(/운영 일지\(|\)/g,'') || '아동';
    var data;
    try { data = JSON.parse(String(row[6]||'{}')); } catch(e) { data = {}; }
    if (!data.date) data = blankData(String(row[1]||''), String(row[4]||''), String(row[2]||''));

    var hasTemplate = disp.getLastRow() >= 40 &&
      disp.getRange('A1').getValue().toString().indexOf('운영일지') >= 0;
    if (!hasTemplate) {
      buildTemplate(disp, data.date||'', data.manager||'', recLbl);
    } else {
      clearDataCells(disp);
    }
    applyData(disp, data, recLbl);
    ss.setActiveSheet(disp);
    // 현재 활성 일지 저장 (사이드바 재오픈 시 복구용)
    PropertiesService.getDocumentProperties().setProperty('activeTitle', title);
    return { ok:true };
  } catch(e) { return { ok:false, error:e.message }; }
}

/**
 * 현재 운영일지 시트 내용 → 데이터 저장소 갱신 + 목록 반환
 * title이 null이면 PropertiesService에서 복구 시도
 */
function saveCurrentData(title) {
  try {
    // title 복구: 사이드바가 재오픈된 경우 PropertiesService에서 가져옴
    if (!title) {
      title = PropertiesService.getDocumentProperties().getProperty('activeTitle');
    }
    if (!title) return { ok:false, error:'저장할 일지를 찾을 수 없어요.\n사이드바에서 일지를 먼저 선택하거나 만들어 주세요.' };

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var store = getStoreSheet(ss);
    var disp  = getDispSheet(ss);

    // 저장할 때도 제목 형식을 한글 일자로 강제 변환 (구글 시트의 자동 날짜 변환 방지)
    var formattedTitle = makeTitleFromDate(title);
    
    var data   = extractData(disp);
    var now    = new Date();
    var json   = JSON.stringify(data);
    var foundRow = findStoreRow(store, formattedTitle);
    var recType  = foundRow > 0 ? String(store.getRange(foundRow,4).getValue()) : '운영 일지(아동)';

    if (foundRow > 0) {
      store.getRange(foundRow,1,1,7).setValues([[formattedTitle,toDateStr(data.date),data.manager,recType,data.opTime,now,json]]);
    } else {
      store.appendRow([formattedTitle,toDateStr(data.date),data.manager,recType,data.opTime,now,json]);
    }
    SpreadsheetApp.flush(); 
    var list = buildList(store);
    PropertiesService.getDocumentProperties().setProperty('activeTitle', formattedTitle);
    return { ok:true, name:formattedTitle, list:list };
  } catch(e) { return { ok:false, error:e.message }; }
}

/**
 * 사이드바 로드 시 이전에 활성화된 일지 타이틀 반환 (active 복구용)
 */
function getActiveTitle() {
  try {
    return PropertiesService.getDocumentProperties().getProperty('activeTitle') || null;
  } catch(e) { return null; }
}

// ══════════════════════════════════════════════════════════════
//  템플릿 시트 관리
// ══════════════════════════════════════════════════════════════

/**
 * 템플릿 시트를 가져오거나 새로 생성 (구조 빌드)
 */
function getTmplSheet(ss) {
  var s = ss.getSheetByName(TMPL);
  if (!s) {
    s = ss.insertSheet(TMPL);
    // 빈 템플릿 구조 만들기 (기본값으로)
    buildTemplate(s, '', '', '아동');
    applyPrintSettings(s, {top:1.5,bottom:1.5,left:1.5,right:1.5});
  }
  return s;
}

/**
 * 사이드바에서 호출: 템플릿 시트를 활성화 시트로 전환
 */
function editTemplate() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tmpl = getTmplSheet(ss);
    ss.setActiveSheet(tmpl);
    return { ok:true };
  } catch(e) { return { ok:false, error:e.message }; }
}

/**
 * 메뉴에서 직접 호출: 템플릿 편집
 */
function editTemplateMenu() {
  var r = editTemplate();
  if (!r.ok) SpreadsheetApp.getUi().alert('오류: ' + r.error);
}

/**
 * 템플릿 시트 존재 여부 및 기본 정보 반환 (사이드바용)
 */
function getTemplateInfo() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s = ss.getSheetByName(TMPL);
    if (!s) return { exists: false };
    // 템플릿 시트의 기본 정보 읽기
    var date = s.getRange('B3').getDisplayValue();
    var manager = s.getRange('J3').getDisplayValue();
    var opTime = s.getRange('F3').getDisplayValue();
    return { exists: true, date: date, manager: manager, opTime: opTime };
  } catch(e) { return { exists: false }; }
}

/**
 * 템플릿 시트 기반으로 새 일지를 생성+등록
 * ★ Sheet.copyTo(ss) 방식: 기존 운영일지 삭제 후 템플릿을 통째로 복사
 *    → 병합·행높이·열너비·서식·값 모두 100% 그대로 반영됨
 * formData: { title, date, manager, record, opdate }
 */
function createFromTemplate(formData) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var store = getStoreSheet(ss);
    var tmpl  = getTmplSheet(ss);  // 템플릿 시트 확보 (없으면 buildTemplate으로 생성)

    var rawDate = (formData && (formData.opdate || formData.date)) || '';
    var title   = makeTitleFromDate(rawDate);
    var date    = rawDate;
    var mgr     = (formData && formData.manager)  || '';
    var recType = (formData && formData.record)   || '운영 일지(아동)';
    var recLbl  = recType.replace(/운영 일지\(|\)/g,'') || '아동';

    // ── 핵심: 기존 운영일지 시트를 삭제하고 템플릿을 통째로 복사 ──
    // 이 방법만이 병합 충돌 없이 서식/병합/행높이/열너비를 100% 복사함

    // 1) 기존 운영일지 시트가 있으면 삭제 (삭제 전 다른 시트를 먼저 활성화)
    var existingDisp = ss.getSheetByName(DISP);
    if (existingDisp) {
      var allSheets = ss.getSheets();
      for (var si = 0; si < allSheets.length; si++) {
        if (allSheets[si].getName() !== DISP) {
          ss.setActiveSheet(allSheets[si]);
          break;
        }
      }
      ss.deleteSheet(existingDisp);
    }

    // 2) 템플릿을 같은 스프레드시트에 복사 → 새 운영일지 시트
    var disp = tmpl.copyTo(ss);
    disp.setName(DISP);
    disp.showSheet();  // 혹시 숨겨진 경우 대비

    // 3) 날짜·담당자만 폼 값으로 덮어쓰기
    if (date) {
      disp.getRange('B3').setNumberFormat('@');  // 날짜 자동변환 방지
      disp.getRange('B3').setValue(date);
    }
    if (mgr) disp.getRange('J3').setValue(mgr);

    // 4) 기록물 종류 레이블 업데이트 (헤더)
    try {
      var a1 = disp.getRange('A1').getValue().toString();
      if (a1.indexOf('운영일지') >= 0 || a1 === '') {
        disp.getRange('A1').setValue('운영일지(' + recLbl + ')');
      }
      var a28 = disp.getRange('A28').getValue().toString();
      if (a28.indexOf('운영일지') >= 0 || a28 === '') {
        disp.getRange('A28').setValue('운영일지(' + recLbl + ')  ─  2 페이지');
      }
    } catch(e) {}

    // 5) 인쇄 설정 (실패해도 생성에 영향 없음)
    try { applyPrintSettings(disp, {top:1.5, bottom:1.5, left:1.5, right:1.5}); } catch(e) {}

    SpreadsheetApp.flush();
    ss.setActiveSheet(disp);

    // 6) 데이터 저장소 등록
    var data = extractData(disp);
    if (date) data.date    = date;
    if (mgr)  data.manager = mgr;
    if (!data.opTime) data.opTime = '09:00~18:00';

    var now      = new Date();
    var json     = JSON.stringify(data);
    var foundRow = findStoreRow(store, title);
    if (foundRow > 0) {
      store.getRange(foundRow,1,1,7).setValues(
        [[title, toDateStr(data.date), data.manager, recType, data.opTime, now, json]]
      );
    } else {
      store.appendRow([title, toDateStr(data.date), data.manager, recType, data.opTime, now, json]);
    }
    PropertiesService.getDocumentProperties().setProperty('activeTitle', title);

    return { ok:true, name:title, list:buildList(store) };
  } catch(e) {
    return { ok:false, error:'[createFromTemplate] ' + e.message + ' (line:' + e.lineNumber + ')' };
  }
}

/**
 * 항목 삭제 + 목록 반환
 */
function deleteEntry(title) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var store = getStoreSheet(ss);
    var foundRow = findStoreRow(store, title);
    if (foundRow < 0) return { ok:false, error:'항목 없음' };
    store.deleteRow(foundRow);
    return { ok:true, list:buildList(store) };
  } catch(e) { return { ok:false, error:e.message }; }
}

function makeDefaultTitle() {
  return makeTitleFromDate(new Date());
}

/** 날짜 값(문자열/객체)을 'M월 D일 요일' 형식으로 변환 */
function makeTitleFromDate(val) {
  if (!val) return makeTitleFromDate(new Date());
  try {
    var d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return String(val); // 변환 불가시 그대로
    
    var m = d.getMonth() + 1;
    var day = d.getDate();
    var days = ['일','월','화','수','목','금','토'];
    var wd = days[d.getDay()];
    
    return m + '월 ' + day + '일 ' + wd;
  } catch(e) { return String(val); }
}

// ══════════════════════════════════════════════════════════════
//  메뉴에서 직접 호출
// ══════════════════════════════════════════════════════════════
function createNewJournal() {
  var r = createFromTemplate(null);
  SpreadsheetApp.getUi().alert(r.ok
    ? '"'+r.name+'" 일지 생성 완료! (템플릿 기반)\n사이드바에서 데이터 입력 후 저장하세요.'
    : '오류: '+r.error);
}
function saveCurrentDataMenu() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var disp = getDispSheet(ss);
  // 제목을 A1 셀 값에서 추출하거나 오늘 날짜 사용
  var title = disp.getRange('B3').getDisplayValue() || makeDefaultTitle();
  var r = saveCurrentData(title);
  SpreadsheetApp.getUi().alert(r.ok ? '저장 완료!' : '오류: '+r.error);
}
function applyPrintSettingsMenu() {
  applyPrintSettings(getDispSheet(SpreadsheetApp.getActiveSpreadsheet()),
                     {top:1.5,bottom:1.5,left:1.5,right:1.5});
  SpreadsheetApp.getUi().alert('A4 2페이지 인쇄 설정 적용 완료!');
}

// ── 사이드바에서 호출 ──
function applyPrintSettingsSidebar(margins) {
  try {
    applyPrintSettings(getDispSheet(SpreadsheetApp.getActiveSpreadsheet()), margins);
    return { ok:true, name:DISP };
  } catch(e) { return { ok:false, error:e.message }; }
}
function getPrintSettings() {
  try {
    var ps = getDispSheet(SpreadsheetApp.getActiveSpreadsheet()).getPageSetup();
    return { ok:true, top:cmFromPt(ps.getTopMargin()), bottom:cmFromPt(ps.getBottomMargin()),
             left:cmFromPt(ps.getLeftMargin()), right:cmFromPt(ps.getRightMargin()) };
  } catch(e) { return { ok:false, top:1.5, bottom:1.5, left:1.5, right:1.5 }; }
}
function cmFromPt(pt) { return (pt==null)?1.5:Math.round((pt/28.3465)*10)/10; }

// ══════════════════════════════════════════════════════════════
//  인쇄 설정
// ══════════════════════════════════════════════════════════════
function applyPrintSettings(sheet, m) {
  try {
    m = m || {top:1.5,bottom:1.5,left:1.5,right:1.5};
    var pt = 28.3465;
    // getPageSetup()이 없는 환경도 있으므로 안전하게 처리
    if (typeof sheet.getPageSetup !== 'function') return;
    var ps = sheet.getPageSetup();
    ps.setPaperSize(SpreadsheetApp.PaperSize.A4);
    ps.setPageOrientation(SpreadsheetApp.PageOrientation.PORTRAIT);
    ps.setTopMargin(m.top*pt); ps.setBottomMargin(m.bottom*pt);
    ps.setLeftMargin(m.left*pt); ps.setRightMargin(m.right*pt);
    ps.setFitToWidth(1); ps.setFitToHeight(2);
    ps.setPrintGridlines(false); ps.setPrintHeadings(false);
    ps.setPageOrder(SpreadsheetApp.PageOrder.TOP_TO_BOTTOM);
    ps.setHorizontalCentered(true); ps.setVerticalCentered(false);
    var lr=sheet.getLastRow(), lc=sheet.getLastColumn();
    if (lr>0&&lc>0) ps.setPrintRange(sheet.getRange(1,1,lr,lc));
    sheet.setPageSetup(ps); // 변경사항 명시적 저장
  } catch(e) {
    // 인쇄 설정 실패는 무시 (일지 생성에는 영향 없음)
    Logger.log('applyPrintSettings 건너뜀: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  양식 템플릿 생성 (배열 일괄 처리로 최적화)
// ══════════════════════════════════════════════════════════════
function buildTemplate(sheet, dateStr, managerVal, recordLabel) {
  recordLabel = recordLabel || '아동';
  managerVal  = managerVal  || '';

  var NAVY='#1a3a5c',NAVY2='#334e6e',LBL='#e8edf5',WHITE='#ffffff';
  var SUM_BG='#f0f9e8',SUM_FG='#16a34a',SEC_BG='#dce6f4',GRAY='#f8fafc';
  var T=14, LAST=52;

  // 행 높이
  var rh={1:42,2:6,3:28,4:6,5:20,6:22,7:24,8:24,9:6,10:20,11:26,12:6,13:20,14:26,15:6,
          16:20,17:22,18:26,19:6,20:20,21:22,22:26,23:26,24:26,25:26,26:26,
          27:8,28:30,29:6,30:20,31:90,32:6,33:20,34:24,35:28,36:28,37:28,38:28,39:28,40:28,
          41:6,42:20,43:24,44:26,45:26,46:26,47:26,48:26,49:6,50:22,51:40,52:20};
  Object.keys(rh).forEach(function(r){ sheet.setRowHeight(Number(r),rh[r]); });
  [72,40,60,60,68,60,68,48,48,52,52,52,52,52].forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });

  // 52×14 배열 초기화
  var vals=[], bgs=[], fws=[], fcs=[], has=[], vas=[];
  for (var r=0;r<LAST;r++){
    vals.push(new Array(T).fill(''));    bgs.push(new Array(T).fill(WHITE));
    fws.push(new Array(T).fill('normal')); fcs.push(new Array(T).fill('#1e293b'));
    has.push(new Array(T).fill('center')); vas.push(new Array(T).fill('middle'));
  }
  function c(r,col,v,bg,fw,ha){ if(v!==undefined)vals[r][col]=v; if(bg)bgs[r][col]=bg; if(fw)fws[r][col]=fw; if(ha)has[r][col]=ha; }
  function rowFill(r,cs,ce,v,bg,fw,ha){ for(var j=cs;j<=ce;j++)c(r,j,v,bg,fw,ha); }
  function sec(r,label){ rowFill(r,0,T-1,label,SEC_BG,'bold','left'); fcs[r].fill(NAVY); }

  rowFill(0,0,T-1,'운영일지('+recordLabel+')',NAVY,'bold','center'); fcs[0].fill('#fff');
  c(2,0,'일자',LBL,'bold'); c(2,1,dateStr,WHITE); c(2,4,'운영시간',LBL,'bold'); c(2,5,'09:00~18:00',WHITE);
  c(2,8,'담당자',LBL,'bold'); c(2,9,managerVal,WHITE);
  sec(4,'■ 아동현황 (취학구분)');
  ['구분','성별','취학전','탈학교','초등학교','중학교','고등학교','기타','계'].forEach(function(v,i){ c(5,i,v,LBL,'bold'); });
  c(6,0,'아동\n현황',LBL,'bold'); c(6,1,'남',LBL,'bold');
  for(var i=2;i<8;i++){c(6,i,'0',WHITE);c(7,i,'0',WHITE);} c(6,8,'',SUM_BG); fws[6][8]='bold'; fcs[6][8]=SUM_FG;
  c(7,1,'여',LBL,'bold'); c(7,8,'',SUM_BG); fws[7][8]='bold'; fcs[7][8]=SUM_FG;
  sec(9,'■ 급식현황'); c(10,0,'조식',LBL,'bold');c(10,1,'0',WHITE);c(10,2,'중식',LBL,'bold');c(10,3,'0',WHITE);c(10,4,'석식',LBL,'bold');c(10,5,'0',WHITE);
  sec(12,'■ 교사현황'); c(13,0,'종사자',LBL,'bold');c(13,1,'0',WHITE);c(13,2,'교사',LBL,'bold');c(13,3,'0',WHITE);c(13,4,'강사',LBL,'bold');c(13,5,'0',WHITE);c(13,6,'기타',LBL,'bold');c(13,7,'0',WHITE);
  sec(15,'■ 아동출석 (출석구분)');
  ['정원','현원','출석','공결','대체출석','결석','기타'].forEach(function(v,i){ c(16,i,v,LBL,'bold'); c(17,i,'0',WHITE); });
  sec(19,'■ 프로그램 현황');
  ['No','프로그램명','시간','대상','인원','담당자'].forEach(function(v,i){ c(20,i,v,LBL,'bold'); });
  c(20,6,'내용',LBL,'bold');
  for(var row=21;row<=25;row++){ c(row,0,String(row-20),GRAY); for(var j=1;j<14;j++)c(row,j,'',WHITE); }
  rowFill(26,0,T-1,'','#c8d5e8');
  rowFill(27,0,T-1,'운영일지('+recordLabel+')  ─  2 페이지',NAVY2,'bold'); fcs[27].fill('#fff');
  sec(29,'■ 특이사항 / 메모'); rowFill(30,0,T-1,'',WHITE); vas[30].fill('top'); has[30].fill('left');
  sec(32,'■ 업무일지');
  c(33,0,'구분',LBL,'bold');c(33,2,'내용',LBL,'bold');c(33,8,'처리결과',LBL,'bold');c(33,12,'담당자',LBL,'bold');
  for(var row2=34;row2<=39;row2++){[0,1,2,3,4,5,6,7,8,9,10,11,12,13].forEach(function(j){c(row2,j,'',WHITE);});}
  sec(41,'■ 안전점검');
  c(42,0,'점검항목',LBL,'bold');c(42,3,'점검결과',LBL,'bold');c(42,10,'비고',LBL,'bold');
  for(var row3=43;row3<=47;row3++){[0,1,2,3,4,5,6,7,8,9,10,11,12,13].forEach(function(j){c(row3,j,'',WHITE);});}
  c(49,10,'담당',LBL,'bold');c(49,12,'센터장',LBL,'bold');
  c(51,10,'(인)'); fcs[51][10]='#aaa'; c(51,12,'(인)'); fcs[51][12]='#aaa';

  // 배열 일괄 적용
  var full = sheet.getRange(1,1,LAST,T);
  full.setValues(vals); full.setBackgrounds(bgs); full.setFontWeights(fws);
  full.setFontColors(fcs); full.setHorizontalAlignments(has); full.setVerticalAlignments(vas);
  full.setFontSize(11); full.setWrap(true);

  // 수식
  sheet.getRange('I7').setFormula('=SUM(C7:H7)');
  sheet.getRange('I8').setFormula('=SUM(C8:H8)');

  // 병합
  ['A1:N1','B3:D3','F3:H3','J3:N3','A5:N5','A7:A8',
   'A10:N10','A13:N13','A16:N16','A20:N20',
   'G21:N21','G22:N22','G23:N23','G24:N24','G25:N25','G26:N26',
   'A27:N27','A28:N28','A30:N30','A31:N31','A33:N33',
   'A34:B34','C34:H34','I34:L34','M34:N34',
   'A35:B35','C35:H35','I35:L35','M35:N35','A36:B36','C36:H36','I36:L36','M36:N36',
   'A37:B37','C37:H37','I37:L37','M37:N37','A38:B38','C38:H38','I38:L38','M38:N38',
   'A39:B39','C39:H39','I39:L39','M39:N39','A40:B40','C40:H40','I40:L40','M40:N40',
   'A42:N42','A43:C43','D43:J43','K43:N43',
   'A44:C44','D44:J44','K44:N44','A45:C45','D45:J45','K45:N45',
   'A46:C46','D46:J46','K46:N46','A47:C47','D47:J47','K47:N47','A48:C48','D48:J48','K48:N48',
   'K50:L50','M50:N50','K51:L51','M51:N51','K52:L52','M52:N52'
  ].forEach(function(a){ try{sheet.getRange(a).merge();}catch(e){} });

  // 테두리
  var S=SpreadsheetApp.BorderStyle.SOLID, M=SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  sheet.getRange(1,1,LAST,T).setBorder(true,true,true,true,true,true,'#aaaaaa',S);
  sheet.getRange('A1:N1').setBorder(true,true,true,true,null,null,NAVY,M);
  sheet.getRange('A28:N52').setBorder(true,true,true,true,null,null,'#7a96bb',M);
  SpreadsheetApp.flush();
}
