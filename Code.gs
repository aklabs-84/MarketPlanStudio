function doPost(e) {
  try {
    // ⭐ 스프레드시트 ID (URL에서 추출해서 입력하세요)
    const SPREADSHEET_ID = '1zFo4PB_I437On423PWcKE4fycxzi7lENtyd-XaCX1jA';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName('마케팅기획') || ss.insertSheet('마케팅기획');

    // 헤더 설정 (첫 실행시에만)
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,18).setValues([[
        '타임스탬프', '기획자명', '상품명', '상품설명', '상품카테고리',
        'MBTI질문', 'MBTI연결아이디어', '포스터메인문구', '포스터비주얼', '타겟고객',
        '영상콘셉트', '영상장면', '영상길이', 'MBTI페이지제목', 'MBTI설명',
        '포스터메인카피', '포스터서브카피'
      ]]);

      // 헤더 스타일링
      const headerRange = sh.getRange(1, 1, 1, 18);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('white');
    }

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

    // 새 행에 데이터 추가
    sh.appendRow([
      data.timestamp || new Date().toLocaleString('ko-KR'),
      data.plannerName || '',
      data.productName || '',
      data.productDescription || '',
      data.productCategory || '',
      data.mbtiQuestion || '',
      data.mbtiConnection || '',
      data.posterMain || '',
      data.posterVisual || '',
      data.targetAudience || '',
      data.videoConcept || '',
      data.scenes || '',
      data.videoDuration || '',
      data.mbtiPageTitle || '',
      data.mbtiDescription || '',
      data.posterMainCopy || '',
      data.posterSubCopy || '',
      data.budgetTimeline || ''
    ]);

    // 성공 응답
    return ContentService
      .createTextOutput(JSON.stringify({status: 'success', message: '마케팅 기획서가 저장되었습니다!'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('Error: ' + err);
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', message: String(err)}))
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
    .createTextOutput('Apps Script is working!')
    .setMimeType(ContentService.MimeType.TEXT);
}
