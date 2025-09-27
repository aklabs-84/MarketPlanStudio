/** 마케팅 기획 도구 날짜별 시트 저장 - Google Apps Script **/

const DEFAULT_HEADERS = [
  '타임스탬프', '기획자명', '상품명', '상품설명', '상품카테고리',
  'MBTI질문', 'MBTI연결아이디어', '포스터메인문구', '포스터비주얼', '타겟고객',
  '영상콘셉트', '영상장면', '영상길이', 'MBTI페이지제목', 'MBTI설명',
  '포스터메인카피', '포스터서브카피', '예산일정'
];

const HEADER_KEYS = {
  plannerName:      ['기획자명', '기획자', 'planner', 'plannername'],
  productName:      ['상품명', '상품이름', 'product', 'productname'],
  productDesc:      ['상품설명', '상품설명', 'productdescription', 'description'],
  productCategory:  ['상품카테고리', '카테고리', 'category', 'productcategory'],
  mbtiQuestion:     ['MBTI질문', 'mbti질문', 'mbtiquestion'],
  mbtiConnection:   ['MBTI연결아이디어', 'mbti연결', 'mbticonnection'],
  posterMain:       ['포스터메인문구', '포스터메인', 'postermain'],
  posterVisual:     ['포스터비주얼', '비주얼', 'postervisual'],
  targetAudience:   ['타겟고객', '타겟', 'target', 'targetaudience'],
  videoConcept:     ['영상콘셉트', '영상개념', 'videoconcept'],
  scenes:           ['영상장면', '장면', 'scenes'],
  videoDuration:    ['영상길이', '영상시간', 'videoduration'],
  mbtiPageTitle:    ['MBTI페이지제목', 'mbti제목', 'mbtipagetitle'],
  mbtiDescription:  ['MBTI설명', 'mbti설명', 'mbtidescription'],
  posterMainCopy:   ['포스터메인카피', '메인카피', 'postermaincopy'],
  posterSubCopy:    ['포스터서브카피', '서브카피', 'postersubcopy'],
  budgetTimeline:   ['예산일정', '예산', 'budget', 'budgettimeline'],
  ts:               ['타임스탬프', '타임스태프', 'timestamp', '시간', '작성시각']
};

function norm_(s){ return String(s||'').trim().toLowerCase().replace(/\s+/g,''); }

// 날짜 기반 시트명 생성 함수
function getDateSheetName_(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `마케팅_${year}-${month}-${day}`;
}

// 날짜별 시트 가져오기 또는 생성
function getOrCreateDateSheet_(ss, date) {
  const sheetName = getDateSheetName_(date);
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    // 새 시트 생성
    sheet = ss.insertSheet(sheetName);
    
    // 시트 탭 색상 설정 (마케팅용 색상)
    const colors = ['#4285f4', '#34a853', '#fbbc05', '#ea4335', '#9c27b0', '#ff9800'];
    const colorIndex = Math.floor(Math.random() * colors.length);
    sheet.setTabColor(colors[colorIndex]);
  }
  
  return sheet;
}

function ensureHeaders_(sh){
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,DEFAULT_HEADERS.length).setValues([DEFAULT_HEADERS]);
    
    // 헤더 스타일링
    const headerRange = sh.getRange(1, 1, 1, DEFAULT_HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    headerRange.setHorizontalAlignment('center');
    
    // 열 너비 자동 조정
    sh.autoResizeColumns(1, DEFAULT_HEADERS.length);
    return;
  }
  
  const lastCol = sh.getLastColumn();
  const row1 = sh.getRange(1,1,1,lastCol).getValues()[0];
  const normRow = row1.map(norm_);

  // 각 키의 대표 헤더가 없으면 추가, 비표준 변형이면 표준 명칭으로 교체
  const want = {
    plannerName: '기획자명', productName: '상품명', productDesc: '상품설명', 
    productCategory: '상품카테고리', mbtiQuestion: 'MBTI질문', mbtiConnection: 'MBTI연결아이디어',
    posterMain: '포스터메인문구', posterVisual: '포스터비주얼', targetAudience: '타겟고객',
    videoConcept: '영상콘셉트', scenes: '영상장면', videoDuration: '영상길이',
    mbtiPageTitle: 'MBTI페이지제목', mbtiDescription: 'MBTI설명', 
    posterMainCopy: '포스터메인카피', posterSubCopy: '포스터서브카피', 
    budgetTimeline: '예산일정', ts: '타임스탬프'
  };
  
  Object.entries(HEADER_KEYS).forEach(([k, alts])=>{
    const idx = normRow.findIndex(h => alts.includes(h));
    if (idx === -1) {
      // 맨 뒤에 새로 추가
      sh.getRange(1, sh.getLastColumn()+1).setValue(want[k]);
      sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold');
    } else {
      // 표준 명칭으로 교체
      sh.getRange(1, idx+1).setValue(want[k]);
    }
  });
}

function findCol_(sh, key){ // key in HEADER_KEYS
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  const normHeaders = headers.map(norm_);
  const alts = HEADER_KEYS[key];
  const idx = normHeaders.findIndex(h => alts.includes(h));
  return idx === -1 ? null : idx+1; // 1-based
}

function doPost(e) {
  try {
    // 스프레드시트 ID (실제 ID로 변경 필요)
    const SPREADSHEET_ID = '1zFo4PB_I437On423PWcKE4fycxzi7lENtyd-XaCX1jA';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 현재 날짜로 시트 결정
    const currentDate = new Date();
    const sh = getOrCreateDateSheet_(ss, currentDate);

    ensureHeaders_(sh);

    // POST 데이터 파싱
    let data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        Logger.log('JSON parse error: ' + parseErr);
        data = e.parameter || {};
      }
    } else {
      data = e ? (e.parameter || {}) : {};
    }

    Logger.log('Received data: ' + JSON.stringify(data));

    // 값 준비
    const rawTs = data.timestamp || new Date().toLocaleString('ko-KR');

    // 열 찾기
    const tsCol = findCol_(sh,'ts');
    const plannerNameCol = findCol_(sh,'plannerName');
    const productNameCol = findCol_(sh,'productName');
    const productDescCol = findCol_(sh,'productDesc');
    const productCategoryCol = findCol_(sh,'productCategory');
    const mbtiQuestionCol = findCol_(sh,'mbtiQuestion');
    const mbtiConnectionCol = findCol_(sh,'mbtiConnection');
    const posterMainCol = findCol_(sh,'posterMain');
    const posterVisualCol = findCol_(sh,'posterVisual');
    const targetAudienceCol = findCol_(sh,'targetAudience');
    const videoConceptCol = findCol_(sh,'videoConcept');
    const scenesCol = findCol_(sh,'scenes');
    const videoDurationCol = findCol_(sh,'videoDuration');
    const mbtiPageTitleCol = findCol_(sh,'mbtiPageTitle');
    const mbtiDescriptionCol = findCol_(sh,'mbtiDescription');
    const posterMainCopyCol = findCol_(sh,'posterMainCopy');
    const posterSubCopyCol = findCol_(sh,'posterSubCopy');
    const budgetTimelineCol = findCol_(sh,'budgetTimeline');

    const rowLen = sh.getLastColumn();
    const row = new Array(rowLen).fill('');

    if (tsCol) row[tsCol-1] = rawTs;
    if (plannerNameCol) row[plannerNameCol-1] = data.plannerName || '';
    if (productNameCol) row[productNameCol-1] = data.productName || '';
    if (productDescCol) row[productDescCol-1] = data.productDescription || '';
    if (productCategoryCol) row[productCategoryCol-1] = data.productCategory || '';
    if (mbtiQuestionCol) row[mbtiQuestionCol-1] = data.mbtiQuestion || '';
    if (mbtiConnectionCol) row[mbtiConnectionCol-1] = data.mbtiConnection || '';
    if (posterMainCol) row[posterMainCol-1] = data.posterMain || '';
    if (posterVisualCol) row[posterVisualCol-1] = data.posterVisual || '';
    if (targetAudienceCol) row[targetAudienceCol-1] = data.targetAudience || '';
    if (videoConceptCol) row[videoConceptCol-1] = data.videoConcept || '';
    if (scenesCol) row[scenesCol-1] = data.scenes || '';
    if (videoDurationCol) row[videoDurationCol-1] = data.videoDuration || '';
    if (mbtiPageTitleCol) row[mbtiPageTitleCol-1] = data.mbtiPageTitle || '';
    if (mbtiDescriptionCol) row[mbtiDescriptionCol-1] = data.mbtiDescription || '';
    if (posterMainCopyCol) row[posterMainCopyCol-1] = data.posterMainCopy || '';
    if (posterSubCopyCol) row[posterSubCopyCol-1] = data.posterSubCopy || '';
    if (budgetTimelineCol) row[budgetTimelineCol-1] = data.budgetTimeline || '';

    sh.appendRow(row);

    // 데이터 행 스타일링
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      const dataRange = sh.getRange(lastRow, 1, 1, rowLen);
      dataRange.setBorder(true, true, true, true, true, true);
      
      // 교대로 배경색 적용
      if (lastRow % 2 === 0) {
        dataRange.setBackground('#f8f9fa');
      }
    }

    // 성공 응답
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success', 
        message: '마케팅 기획서가 저장되었습니다!',
        sheetName: sh.getName(),
        row: sh.getLastRow()
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('Error: ' + err);
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error', 
        message: String(err)
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// OPTIONS 요청 처리 (CORS 대응)
function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.JSON);
}

// 디버그용 GET 요청
function doGet() {
  return ContentService
    .createTextOutput('마케팅 기획 Apps Script is working! 🚀')
    .setMimeType(ContentService.MimeType.TEXT);
}

// 모든 날짜 시트의 데이터를 통합하여 요약 시트 생성
function createMarketingSummarySheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheetName = '마케팅_전체요약';
    
    // 기존 요약 시트 삭제 후 새로 생성
    let summarySheet = ss.getSheetByName(summarySheetName);
    if (summarySheet) {
      ss.deleteSheet(summarySheet);
    }
    summarySheet = ss.insertSheet(summarySheetName);
    
    // 요약 시트 헤더 설정
    const summaryHeaders = ['날짜', ...DEFAULT_HEADERS];
    summarySheet.getRange(1, 1, 1, summaryHeaders.length).setValues([summaryHeaders]);
    summarySheet.getRange(1, 1, 1, summaryHeaders.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('white')
      .setHorizontalAlignment('center');
    
    const allSheets = ss.getSheets();
    const dateSheets = allSheets.filter(sheet => {
      const name = sheet.getName();
      return /^마케팅_\d{4}-\d{2}-\d{2}$/.test(name); // 마케팅_YYYY-MM-DD 형식만
    });
    
    let summaryRow = 2;
    
    // 각 날짜 시트에서 데이터 수집
    dateSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const dateOnly = sheetName.replace('마케팅_', ''); // 마케팅_ 접두사 제거
      const lastRow = sheet.getLastRow();
      
      if (lastRow > 1) { // 헤더 외에 데이터가 있는 경우
        const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
        
        data.forEach(row => {
          if (row.some(cell => cell !== '')) { // 빈 행이 아닌 경우
            const summaryRowData = [dateOnly, ...row];
            summarySheet.getRange(summaryRow, 1, 1, summaryRowData.length).setValues([summaryRowData]);
            summaryRow++;
          }
        });
      }
    });
    
    // 열 너비 자동 조정
    summarySheet.autoResizeColumns(1, summaryHeaders.length);
    
    // 요약 시트를 첫 번째 위치로 이동
    ss.moveSheet(summarySheet, 1);
    
    SpreadsheetApp.getUi().alert(`마케팅 요약 시트가 생성되었습니다!\n총 ${dateSheets.length}개의 날짜 시트에서 데이터를 통합했습니다.`);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('요약 시트 생성 중 오류 발생: ' + error.toString());
  }
}

// 특정 날짜의 마케팅 데이터 삭제
function deleteMarketingDataByDate() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('마케팅 날짜별 데이터 삭제', 'YYYY-MM-DD 형식으로 삭제할 날짜를 입력하세요:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const dateInput = response.getResponseText().trim();
    
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateInput)) {
      ui.alert('올바른 날짜 형식(YYYY-MM-DD)을 입력해주세요.');
      return;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = `마케팅_${dateInput}`;
    const sheet = ss.getSheetByName(sheetName);
    
    if (sheet) {
      const confirmResponse = ui.alert(`${sheetName} 시트를 삭제하시겠습니까?`, ui.ButtonSet.YES_NO);
      if (confirmResponse == ui.Button.YES) {
        ss.deleteSheet(sheet);
        ui.alert(`${sheetName} 시트가 삭제되었습니다.`);
      }
    } else {
      ui.alert(`${sheetName} 시트를 찾을 수 없습니다.`);
    }
  }
}

// 상품 카테고리별 통계 생성
function createCategoryStatsSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const statsSheetName = '마케팅_카테고리통계';
    
    // 기존 통계 시트 삭제 후 새로 생성
    let statsSheet = ss.getSheetByName(statsSheetName);
    if (statsSheet) {
      ss.deleteSheet(statsSheet);
    }
    statsSheet = ss.insertSheet(statsSheetName);
    
    // 모든 마케팅 날짜 시트에서 데이터 수집
    const allSheets = ss.getSheets();
    const dateSheets = allSheets.filter(sheet => {
      const name = sheet.getName();
      return /^마케팅_\d{4}-\d{2}-\d{2}$/.test(name);
    });
    
    const categoryCount = {};
    
    dateSheets.forEach(sheet => {
      const categoryCol = findCol_(sheet, 'productCategory');
      if (!categoryCol) return;
      
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      
      const categories = sheet.getRange(2, categoryCol, lastRow - 1, 1).getValues().flat();
      
      categories.forEach(category => {
        if (category) {
          categoryCount[category] = (categoryCount[category] || 0) + 1;
        }
      });
    });
    
    // 통계 시트 헤더
    statsSheet.getRange(1, 1).setValue('상품 카테고리');
    statsSheet.getRange(1, 2).setValue('기획 수');
    statsSheet.getRange(1, 3).setValue('비율(%)');
    
    const headerRange = statsSheet.getRange(1, 1, 1, 3);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    headerRange.setHorizontalAlignment('center');
    
    // 데이터 입력
    const totalCount = Object.values(categoryCount).reduce((sum, count) => sum + count, 0);
    const sortedCategories = Object.entries(categoryCount).sort((a, b) => b[1] - a[1]);
    
    sortedCategories.forEach(([category, count], index) => {
      const percentage = totalCount > 0 ? Math.round((count / totalCount) * 100) : 0;
      const row = index + 2;
      
      statsSheet.getRange(row, 1).setValue(category);
      statsSheet.getRange(row, 2).setValue(count);
      statsSheet.getRange(row, 3).setValue(`${percentage}%`);
    });
    
    // 열 너비 자동 조정
    statsSheet.autoResizeColumns(1, 3);
    
    SpreadsheetApp.getUi().alert(`상품 카테고리 통계 시트가 생성되었습니다!\n총 ${totalCount}개의 기획서를 분석했습니다.`);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('통계 시트 생성 중 오류 발생: ' + error.toString());
  }
}

// 테스트 함수
function testMarketingInsert() {
  const testData = {
    timestamp: new Date().toLocaleString('ko-KR'),
    plannerName: '테스트 기획자',
    productName: '친환경 텀블러',
    productDescription: '오래가는 보냉 보틀',
    productCategory: '생활용품',
    mbtiQuestion: '주말에 나는... A. 집콕 (I) / B. 친구 만나기 (E)',
    mbtiConnection: 'ENFP → 활동적인 당신에게, 언제 어디서든 시원한 텀블러 필요!',
    posterMain: '오늘도 시원하게, 친환경 텀블러',
    posterVisual: '파란색, 시원한 물방울, 손에 든 보틀',
    targetAudience: '20~30대 직장인, 운동을 좋아하는 사람들',
    videoConcept: '더위에 지친 순간, 한 모금으로 살아나는 장면',
    scenes: '첫번째 장면 | 두번째 장면 | 세번째 장면',
    videoDuration: '30초',
    mbtiPageTitle: '내 성격에 맞는 음료 찾기!',
    mbtiDescription: '몇 가지 질문으로 나에게 딱 맞는 상품을 찾아보세요',
    posterMainCopy: '환경도 지키고 스타일도 지킨다',
    posterSubCopy: '친환경 텀블러 하나로 내 일상에 변화를!',
    budgetTimeline: ''
  };
  
  const e = {
    postData: {
      contents: JSON.stringify(testData)
    }
  };
  
  const result = doPost(e);
  Logger.log('Test result: ' + result.getContent());
}

// 메뉴 추가
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 마케팅 데이터 관리')
    .addItem('📋 전체 요약 시트 생성', 'createMarketingSummarySheet')
    .addItem('📈 카테고리별 통계 시트 생성', 'createCategoryStatsSheet')
    .addItem('🗑️ 날짜별 데이터 삭제', 'deleteMarketingDataByDate')
    .addItem('🧪 테스트 데이터 추가', 'testMarketingInsert')
    .addToUi();
}
