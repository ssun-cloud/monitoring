const DRIVE_FOLDER_ID = '1Zslt_Mkwrj5CwOjL7_hdB7q8ZolmhRTG';
const MANAGER_EMAIL = 'smlee@jpurme.org';

const SHEET_APPLY = '신청';
const SHEET_MONITOR = '모니터링';
const SHEET_UNIFIED = '통합현황';
const SHEET_RECOMMEND = '상점추천';
const SHEET_JOURNAL = '활동일지';
const SHEET_SETTINGS = '설정';

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let apply = ss.getSheetByName(SHEET_APPLY);
  if (!apply) apply = ss.insertSheet(SHEET_APPLY);
  apply.getRange(1, 1, 1, 9).setValues([[
    '타임스탬프', '신청자', '상호명', '주소', '신청항목', '연락처', '메모', '상태', '개인정보동의'
  ]]);
  apply.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#E8F5E9');
  apply.setColumnWidth(1, 160);
  apply.setColumnWidth(4, 220);
  apply.setColumnWidth(7, 240);

  let monitor = ss.getSheetByName(SHEET_MONITOR);
  if (!monitor) monitor = ss.insertSheet(SHEET_MONITOR);
  monitor.getRange(1, 1, 1, 14).setValues([[
    '타임스탬프', '장소명', '주소', '모니터링날짜',
    '점자메뉴판', '경사로·출입접근성', '화장실접근성', '안내판가독성',
    '조치필요여부', '종합의견', '추가확인항목', '기록자', '상태', '사진링크'
  ]]);
  monitor.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#E3F2FD');
  monitor.setColumnWidth(1, 160);
  monitor.setColumnWidth(3, 220);
  monitor.setColumnWidth(10, 260);
  monitor.setColumnWidth(14, 420);

  let unified = ss.getSheetByName(SHEET_UNIFIED);
  if (!unified) unified = ss.insertSheet(SHEET_UNIFIED);
  unified.getRange(1, 1, 1, 10).setValues([[
    '타임스탬프', '유형', '장소·상호명', '주소', '신청·점검항목',
    '조치필요여부', '종합내용', '담당자·기록자', '상태', '사진링크'
  ]]);
  unified.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#FFF8E1');
  unified.setColumnWidth(1, 160);
  unified.setColumnWidth(3, 180);
  unified.setColumnWidth(4, 230);
  unified.setColumnWidth(5, 330);
  unified.setColumnWidth(7, 260);
  unified.setColumnWidth(10, 420);

  let recommend = ss.getSheetByName(SHEET_RECOMMEND);
  if (!recommend) recommend = ss.insertSheet(SHEET_RECOMMEND);
  recommend.getRange(1, 1, 1, 8).setValues([[
    '타임스탬프', '상호명', '주소', '좋은점', '추천인', '연락처', '한줄소개', '상태'
  ]]);
  recommend.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#FCE4EC');
  recommend.setColumnWidth(1, 160);
  recommend.setColumnWidth(2, 160);
  recommend.setColumnWidth(3, 230);
  recommend.setColumnWidth(4, 260);
  recommend.setColumnWidth(7, 260);

  let journal = ss.getSheetByName(SHEET_JOURNAL);
  if (!journal) journal = ss.insertSheet(SHEET_JOURNAL);
  journal.getRange(1, 1, 1, 8).setValues([[
    '타임스탬프', '활동날짜', '활동시간', '팀', '팀원(작성자)', '방문상점', '활동소감', '사진링크'
  ]]);
  journal.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#F3E8FF');
  journal.setColumnWidth(1, 160);
  journal.setColumnWidth(5, 170);
  journal.setColumnWidth(6, 320);
  journal.setColumnWidth(7, 360);
  journal.setColumnWidth(8, 420);

  let settings = ss.getSheetByName(SHEET_SETTINGS);
  if (!settings) settings = ss.insertSheet(SHEET_SETTINGS);
  if (settings.getLastRow() === 0) {
    settings.getRange(1, 1, 1, 3).setValues([['구분', '이름', '값']]);
    settings.getRange(2, 1, 2, 3).setValues([
      ['team', '1팀', '홍길동, 김철수'],
      ['team', '2팀', '김종로, 이선민']
    ]);
  }
  settings.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#E0F2FE');
  settings.setColumnWidth(1, 100);
  settings.setColumnWidth(2, 130);
  settings.setColumnWidth(3, 300);

  const old = ss.getSheetByName('배리어프리 마을 만들기');
  if (old) ss.deleteSheet(old);

  const sheet1 = ss.getSheetByName('시트1');
  if (sheet1) ss.deleteSheet(sheet1);

  rebuildUnified();
  SpreadsheetApp.getUi().alert('시트 세팅 완료!');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('배리어프리 관리')
    .addItem('통합현황 갱신', 'refreshUnified')
    .addItem('시트 초기 세팅', 'setup')
    .addToUi();
}

function makeResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function safeSendEmail_(subject, body) {
  try {
    MailApp.sendEmail(MANAGER_EMAIL, subject, body);
  } catch (err) {
    Logger.log('메일 발송 실패: ' + err.toString());
  }
}

function normalizeDate_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(value).trim();
  const m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[0] : s;
}

function parseTimestamp_(value) {
  if (value instanceof Date) return value.getTime();
  const s = String(value || '').trim();
  const m = s.match(/(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})\.\s*(오전|오후)?\s*(\d{1,2})?:?(\d{1,2})?:?(\d{1,2})?/);
  if (m) {
    let hour = Number(m[5] || 0);
    if (m[4] === '오후' && hour < 12) hour += 12;
    if (m[4] === '오전' && hour === 12) hour = 0;
    return new Date(
      Number(m[1]),
      Number(m[2]) - 1,
      Number(m[3]),
      hour,
      Number(m[6] || 0),
      Number(m[7] || 0)
    ).getTime();
  }
  const parsed = new Date(s).getTime();
  return isNaN(parsed) ? 0 : parsed;
}

function getSettings_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  const props = PropertiesService.getScriptProperties();
  const result = {
    monitorPinSet: Boolean(props.getProperty('MONITOR_PIN')),
    managerPinSet: Boolean(props.getProperty('MANAGER_PIN')),
    teams: []
  };

  if (!sheet) return result;
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const type = String(rows[i][0] || '').trim();
    const name = String(rows[i][1] || '').trim();
    const value = String(rows[i][2] || '').trim();
    if (type === 'team' && name) {
      result.teams.push({
        name: name,
        members: value ? value.split(',').map(s => s.trim()).filter(Boolean) : []
      });
    }
  }
  return result;
}

function checkPin_(pinType, pin) {
  const props = PropertiesService.getScriptProperties();
  const key = pinType === 'manager' ? 'MANAGER_PIN' : 'MONITOR_PIN';
  const saved = props.getProperty(key);
  if (!saved) return false;
  return String(pin || '') === saved;
}

function requireManagerPin_(pin) {
  if (checkPin_('manager', pin)) return null;
  return makeResponse({ result: 'error', message: 'manager authorization required' });
}

function requireMonitorOrManagerPin_(monitorPin, managerPin) {
  if (checkPin_('monitor', monitorPin) || checkPin_('manager', managerPin)) return null;
  return makeResponse({ result: 'error', message: 'monitor authorization required' });
}

function setPinsOnce() {
  PropertiesService.getScriptProperties().setProperties({
    MONITOR_PIN: '7051',
    MANAGER_PIN: '0517'
  });
  Logger.log('비밀번호 저장 완료');
}

function rebuildUnified() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const unified = ss.getSheetByName(SHEET_UNIFIED);
  const applySheet = ss.getSheetByName(SHEET_APPLY);
  const monitorSheet = ss.getSheetByName(SHEET_MONITOR);
  const recommendSheet = ss.getSheetByName(SHEET_RECOMMEND);
  if (!unified) return;

  const lastRow = unified.getLastRow();
  if (lastRow > 1) {
    unified.getRange(2, 1, lastRow - 1, 10).clearContent();
  }

  const rows = [];

  if (applySheet) {
    const data = applySheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[0]) continue;
      rows.push([
        r[0],
        '보급 신청',
        r[2] || '',
        r[3] || '',
        r[4] || '',
        '',
        r[6] || '',
        r[1] || '',
        r[7] || '신청 접수',
        ''
      ]);
    }
  }

  if (monitorSheet) {
    const data = monitorSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[0]) continue;

      const ratings = [
        r[4] ? '점자메뉴판:' + r[4] : '',
        r[5] ? '경사로:' + r[5] : '',
        r[6] ? '화장실:' + r[6] : '',
        r[7] ? '안내판:' + r[7] : ''
      ].filter(Boolean).join(' / ');

      rows.push([
        r[0],
        '모니터링',
        r[1] || '',
        r[2] || '',
        ratings,
        r[8] || '',
        r[9] || '',
        r[11] || '',
        r[12] || '확인 중',
        r[13] || ''
      ]);
    }
  }

  if (recommendSheet) {
    const data = recommendSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[0]) continue;
      rows.push([
        r[0],
        '상점 추천',
        r[1] || '',
        r[2] || '',
        r[3] || '',
        '',
        r[6] || '',
        r[4] || '',
        r[7] || '추천',
        ''
      ]);
    }
  }


  rows.sort((a, b) => parseTimestamp_(b[0]) - parseTimestamp_(a[0]));

  if (rows.length > 0) {
    unified.getRange(2, 1, rows.length, 10).setValues(rows);
  }
}

function refreshUnified() {
  rebuildUnified();
  SpreadsheetApp.getUi().alert('통합현황 갱신 완료!');
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);
    Logger.log('doPost type: ' + data.type);

    if (data.type === 'apply') {
      const sheet = ss.getSheetByName(SHEET_APPLY);
      sheet.appendRow([
        data.timestamp,
        data.applyname || '',
        data.name || '',
        data.addr || '',
        data.items || '',
        data.phone || '',
        data.memo || '',
        '신청 접수',
        data.privacy || 'N'
      ]);
      rebuildUnified();

      safeSendEmail_(
        `[배리어프리 신청] ${data.name || ''} - ${data.items || ''}`,
        `새로운 신청이 접수되었습니다.\n\n` +
        `신청자: ${data.applyname || '미입력'}\n` +
        `상호명: ${data.name || ''}\n` +
        `주소: ${data.addr || ''}\n` +
        `신청 항목: ${data.items || ''}\n` +
        `연락처: ${data.phone || ''}\n` +
        `메모: ${data.memo || '없음'}\n` +
        `접수 시간: ${data.timestamp || ''}`
      );

    } else if (data.type === 'monitor') {
      const authError = requireMonitorOrManagerPin_(data.monitorPin, data.managerPin);
      if (authError) return authError;
      const sheet = ss.getSheetByName(SHEET_MONITOR);
      sheet.appendRow([
        data.timestamp,
        data.place || '',
        data.addr || '',
        data.date || '',
        data.r1 || '',
        data.r2 || '',
        data.r3 || '',
        data.r4 || '',
        data.action || '',
        data.comment || '',
        data.extras || '',
        data.author || '',
        '확인 중',
        ''
      ]);
      rebuildUnified();
      
       // ① 모니터링 접수 알림 (항상 발송)
      const ratingLines = [
        data.r1 ? `점자메뉴판: ${data.r1}` : '',
        data.r2 ? `경사로·출입접근성: ${data.r2}` : '',
        data.r3 ? `화장실접근성: ${data.r3}` : '',
        data.r4 ? `안내판가독성: ${data.r4}` : ''
      ].filter(Boolean).join('\n');
 
      safeSendEmail_(
        `[배리어프리 모니터링] ${data.place || ''} 접수`,
        `새로운 모니터링 기록이 접수되었습니다.\n\n` +
        `장소명: ${data.place || ''}\n` +
        `주소: ${data.addr || ''}\n` +
        `모니터링 날짜: ${data.date || ''}\n` +
        `기록자: ${data.author || ''}\n\n` +
        `[ 점검 항목별 결과 ]\n${ratingLines}\n\n` +
        `조치 필요 여부: ${data.action || '미입력'}\n` +
        `종합의견: ${data.comment || '없음'}\n` +
        `접수 시간: ${data.timestamp || ''}`
      );
 
      // ② 조치 필요 시 추가 알림
      if (data.action && data.action.trim() === '예') {
        safeSendEmail_(
          `[⚠️ 조치 필요] ${data.place || ''} - 배리어프리 모니터링`,
          `조치가 필요한 모니터링 기록이 접수되었습니다.\n\n` +
          `장소명: ${data.place || ''}\n` +
          `주소: ${data.addr || ''}\n` +
          `모니터링 날짜: ${data.date || ''}\n` +
          `기록자: ${data.author || ''}\n\n` +
          `[ 점검 항목별 결과 ]\n${ratingLines}\n\n` +
          `종합의견: ${data.comment || '없음'}\n` +
          `추가 확인 항목: ${data.extras || '없음'}\n\n` +
          `※ 해당 기록을 확인하고 조치 여부를 업데이트해 주세요.`
        );
      }

    } else if (data.type === 'recommend') {
      const sheet = ss.getSheetByName(SHEET_RECOMMEND);
      sheet.appendRow([
        data.timestamp,
        data.name || '',
        data.addr || '',
        data.items || '',
        data.recommender || '',
        data.phone || '',
        data.memo || '',
        '추천'
      ]);
      rebuildUnified();

      safeSendEmail_(
        `[배리어프리 추천] ${data.name || ''} - ${data.items || ''}`,
        `새로운 상점 추천이 접수되었습니다.\n\n` +
        `상호명: ${data.name || ''}\n` +
        `주소: ${data.addr || ''}\n` +
        `좋은 점: ${data.items || ''}\n` +
        `추천인: ${data.recommender || ''}\n` +
        `연락처: ${data.phone || '미입력'}\n` +
        `한줄소개: ${data.memo || '없음'}\n` +
        `접수 시간: ${data.timestamp || ''}`
      );

    } else if (data.type === 'journal') {
      const authError = requireMonitorOrManagerPin_(data.monitorPin, data.managerPin);
      if (authError) return authError;
      const sheet = ss.getSheetByName(SHEET_JOURNAL);
      const monitorSheet = ss.getSheetByName(SHEET_MONITOR);
      const photoLinks = [];

      if (monitorSheet && data.date && data.author) {
        const monitorRows = monitorSheet.getDataRange().getValues();
        const authorNames = String(data.author).split(',').map(s => s.trim()).filter(Boolean);
        const targetDate = normalizeDate_(data.date);

        for (let i = 1; i < monitorRows.length; i++) {
          const r = monitorRows[i];
          if (!r[0]) continue;
          if (normalizeDate_(r[3]) !== targetDate) continue;

          const rowAuthor = String(r[11] || '').trim();
          const matched = authorNames.some(name => rowAuthor.includes(name));
          if (!matched) continue;

          const storeName = String(r[1] || '');
          String(r[13] || '').split('\n').filter(Boolean).forEach(url => {
            photoLinks.push(storeName + ': ' + url);
          });
        }
      }

      sheet.appendRow([
        data.timestamp,
        data.date || '',
        data.time || '',
        data.team || '',
        data.author || '',
        data.stores || '',
        data.memo || '',
        photoLinks.join('\n') || ''
      ]);
      rebuildUnified();

    } else if (data.type === 'uploadPhoto') {
      const authError = requireMonitorOrManagerPin_(data.monitorPin, data.managerPin);
      if (authError) return authError;
      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const folderName = (data.place || 'unknown') + '_' + (data.date || '');
      const existFolders = folder.getFoldersByName(folderName);
      const subFolder = existFolders.hasNext() ? existFolders.next() : folder.createFolder(folderName);
      const blob = Utilities.newBlob(
        Utilities.base64Decode(data.photo),
        'image/jpeg',
        (data.place || 'photo') + '_' + (Number(data.photoIndex || 0) + 1) + '.jpg'
      );
      const file = subFolder.createFile(blob);
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (shareErr) {
        Logger.log('공유 설정 실패, 파일 URL만 반환: ' + shareErr.toString());
      }
      return makeResponse({ result: 'ok', url: file.getUrl() });

    } else if (data.type === 'addPhotoLink') {
      const authError = requireMonitorOrManagerPin_(data.monitorPin, data.managerPin);
      if (authError) return authError;
      const sheet = ss.getSheetByName(SHEET_MONITOR);
      const rows = sheet.getDataRange().getValues();
      let targetRow = -1;

      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.timestamp)) {
          targetRow = i + 1;
          break;
        }
      }

      if (targetRow === -1 && data.place && data.date) {
        const targetDate = normalizeDate_(data.date);
        for (let i = rows.length - 1; i >= 1; i--) {
          const rowPlace = String(rows[i][1] || '');
          const rowDate = normalizeDate_(rows[i][3]);
          if (rowPlace === String(data.place) && rowDate === targetDate) {
            targetRow = i + 1;
            break;
          }
        }
      }

      if (targetRow === -1 && rows.length > 1) targetRow = rows.length;
      if (targetRow === -1) {
        return makeResponse({ result: 'error', message: '사진 링크를 넣을 모니터링 행을 찾지 못했습니다.' });
      }

      const existing = String(sheet.getRange(targetRow, 14).getValue() || '').trim();
      sheet.getRange(targetRow, 14).setValue(existing ? existing + '\n' + data.url : data.url);
      rebuildUnified();
      return makeResponse({ result: 'ok' });

    } else if (data.type === 'updateStatus') {
      const authError = requireManagerPin_(data.managerPin);
      if (authError) return authError;
      let sheetName, statusCol;
      if (data.recordType === 'apply') {
        sheetName = SHEET_APPLY; statusCol = 8;
      } else if (data.recordType === 'recommend') {
        sheetName = SHEET_RECOMMEND; statusCol = 8;
      } else {
        sheetName = SHEET_MONITOR; statusCol = 13;
      }
      const sheet = ss.getSheetByName(sheetName);
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.timestamp)) {
          sheet.getRange(i + 1, statusCol).setValue(data.status);
          break;
        }
      }
      rebuildUnified();

    } else if (data.type === 'checkPin') {
      return makeResponse({ result: 'ok', valid: checkPin_(data.pinType || 'monitor', data.pin) });
    }

    return makeResponse({ result: 'ok' });

  } catch (err) {
    Logger.log('doPost 오류: ' + err.toString());
    return makeResponse({ result: 'error', message: err.toString() });
  }
}

function doUploadPhoto(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const folderName = (data.place || 'unknown') + '_' + (data.date || '');
    const folders = folder.getFoldersByName(folderName);
    const subFolder = folders.hasNext() ? folders.next() : folder.createFolder(folderName);
    const rawPhoto = data.photo && data.photo.data ? data.photo.data : data.photo;
    const blob = Utilities.newBlob(
      Utilities.base64Decode(rawPhoto),
      'image/jpeg',
      (data.place || 'photo') + '_' + (Number(data.photoIndex || 0) + 1) + '.jpg'
    );
    const file = subFolder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      Logger.log('공유 설정 실패, 파일 URL만 반환: ' + shareErr.toString());
    }
    return makeResponse({ result: 'ok', url: file.getUrl() });
  } catch (err) {
    Logger.log('사진 업로드 오류: ' + err.toString());
    return makeResponse({ result: 'error', message: err.toString() });
  }
}

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const type = e.parameter.type;

  if (type === 'settings') {
    return makeResponse({ result: 'ok', settings: getSettings_() });
  }

  if (type === 'map') {
    const records = [];
    const addRecord = rec => {
      if (rec.addr) records.push(rec);
    };

    const applySheet = ss.getSheetByName(SHEET_APPLY);
    const applyRows = applySheet ? applySheet.getDataRange().getValues() : [];
    for (let i = 1; i < applyRows.length; i++) {
      const r = applyRows[i];
      if (!r[0]) continue;
      addRecord({
        type: 'apply',
        timestamp: String(r[0]),
        name: String(r[2] || ''),
        addr: String(r[3] || ''),
        items: String(r[4] || ''),
        action: String(r[7] || '신청 접수')
      });
    }

    const monitorSheet = ss.getSheetByName(SHEET_MONITOR);
    const monitorRows = monitorSheet ? monitorSheet.getDataRange().getValues() : [];
    for (let i = 1; i < monitorRows.length; i++) {
      const r = monitorRows[i];
      if (!r[0]) continue;
      addRecord({
        type: 'monitor',
        timestamp: String(r[0]),
        name: String(r[1] || ''),
        addr: String(r[2] || ''),
        date: normalizeDate_(r[3]),
        action: String(r[8] || ''),
        extras: String(r[10] || '')
      });
    }

    const recommendSheet = ss.getSheetByName(SHEET_RECOMMEND);
    const recommendRows = recommendSheet ? recommendSheet.getDataRange().getValues() : [];
    for (let i = 1; i < recommendRows.length; i++) {
      const r = recommendRows[i];
      if (!r[0]) continue;
      addRecord({
        type: 'recommend',
        timestamp: String(r[0]),
        name: String(r[1] || ''),
        addr: String(r[2] || ''),
        items: String(r[3] || ''),
        action: String(r[7] || '추천')
      });
    }

    records.sort((a, b) => parseTimestamp_(b.timestamp) - parseTimestamp_(a.timestamp));
    return makeResponse({ result: 'ok', records });
  }

  if (type === 'apply') {
    const authError = requireManagerPin_(e.parameter.managerPin);
    if (authError) return authError;
    const sheet = ss.getSheetByName(SHEET_APPLY);
    const rows = sheet ? sheet.getDataRange().getValues() : [];
    const records = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      records.push({
        type: 'apply',
        timestamp: String(r[0]),
        applyname: String(r[1] || ''),
        name: String(r[2] || ''),
        addr: String(r[3] || ''),
        items: String(r[4] || ''),
        phone: String(r[5] || ''),
        memo: String(r[6] || ''),
        action: String(r[7] || '신청 접수')
      });
    }
    return makeResponse({ result: 'ok', records });
  }

  if (type === 'monitor') {
    const authError = requireManagerPin_(e.parameter.managerPin);
    if (authError) return authError;
    const sheet = ss.getSheetByName(SHEET_MONITOR);
    const rows = sheet ? sheet.getDataRange().getValues() : [];
    const records = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      records.push({
        type: 'monitor',
        timestamp: String(r[0]),
        name: String(r[1] || ''),
        addr: String(r[2] || ''),
        date: normalizeDate_(r[3]),
        r1: String(r[4] || ''),
        r2: String(r[5] || ''),
        r3: String(r[6] || ''),
        r4: String(r[7] || ''),
        action: String(r[8] || ''),
        comment: String(r[9] || ''),
        extras: String(r[10] || ''),
        author: String(r[11] || ''),
        status: String(r[12] || ''),
        photoLinks: String(r[13] || '')
      });
    }
    return makeResponse({ result: 'ok', records });
  }

  if (type === 'recommend') {
    const authError = requireManagerPin_(e.parameter.managerPin);
    if (authError) return authError;
    const sheet = ss.getSheetByName(SHEET_RECOMMEND);
    const rows = sheet ? sheet.getDataRange().getValues() : [];
    const records = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      records.push({
        type: 'recommend',
        timestamp: String(r[0]),
        name: String(r[1] || ''),
        addr: String(r[2] || ''),
        items: String(r[3] || ''),
        recommender: String(r[4] || ''),
        phone: String(r[5] || ''),
        memo: String(r[6] || ''),
        action: String(r[7] || '추천')
      });
    }
    return makeResponse({ result: 'ok', records });
  }

  if (type === 'journal') {
    const authError = requireManagerPin_(e.parameter.managerPin);
    if (authError) return authError;
    const sheet = ss.getSheetByName(SHEET_JOURNAL);
    const rows = sheet ? sheet.getDataRange().getValues() : [];
    const records = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      records.push({
        timestamp: String(r[0]),
        date: normalizeDate_(r[1]),
        time: String(r[2] || ''),
        team: String(r[3] || ''),
        author: String(r[4] || ''),
        stores: String(r[5] || ''),
        memo: String(r[6] || ''),
        photoLinks: String(r[7] || '')
      });
    }
    return makeResponse({ result: 'ok', records });
  }

  return makeResponse({ result: 'error', message: 'unknown type' });
}

function authorizeDrive() {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  Logger.log(folder.getName());
}

function authorizeDriveCreateFolder() {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const test = folder.createFolder('_권한테스트_' + new Date().getTime());
  test.setTrashed(true);
  Logger.log('Drive 폴더 생성 권한 확인 완료');
}

function testDriveAccess() {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const file = folder.createFile('권한테스트.txt', 'Drive 업로드 권한 테스트', MimeType.PLAIN_TEXT);
  Logger.log(file.getUrl());
  file.setTrashed(true);
}

function testMail() {
  MailApp.sendEmail('smlee@jpurme.org', '테스트 메일', '잘 되나요?');
}
