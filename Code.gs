function doGet(e) {
  var action = e.parameter.action;
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. API LẤY CẤU HÌNH VÀ BỐC ĐỀ THI NGẪU NHIÊN
  if (action == "getExam") {
    var monHoc = e.parameter.monHoc;
    
    // Kiểm tra cấu hình và thời gian làm bài
    var configSheet = sheet.getSheetByName("CaiDatDeThi");
    var configData = configSheet.getDataRange().getValues();
    var examConfig = null;
    
    for (var i = 1; i < configData.length; i++) {
      if (configData[i][0].toString().toLowerCase() === monHoc.toLowerCase()) {
        examConfig = {
          monHoc: configData[i][0],
          thoiGian: configData[i][1],
          ngayBD: configData[i][2] ? new Date(configData[i][2]) : null,
          ngayKT: configData[i][3] ? new Date(configData[i][3]) : null,
          soCau: parseInt(configData[i][4]) || 10
        };
        break;
      }
    }
    
    if (!examConfig) {
      return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Không tìm thấy môn thi này!"})).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Kiểm tra thời hạn giới hạn ngày làm bài
    var now = new Date();
    if (examConfig.ngayBD && now < examConfig.ngayBD) {
      return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Đề thi chưa đến giờ mở! Vui lòng quay lại sau."})).setMimeType(ContentService.MimeType.JSON);
    }
    if (examConfig.ngayKT && now > examConfig.ngayKT) {
      return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Đề thi đã hết hạn làm bài!"})).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Bốc câu hỏi từ ngân hàng đề
    var dbSheet = sheet.getSheetByName("NganHangDe");
    var qData = dbSheet.getDataRange().getValues();
    var filteredQuestions = [];
    
    for (var i = 1; i < qData.length; i++) {
      if (qData[i][1].toString().toLowerCase() === monHoc.toLowerCase()) {
        filteredQuestions.push({
          id: qData[i][0],
          type: qData[i][2],
          question: qData[i][3],
          image: qData[i][4],
          a: qData[i][5],
          b: qData[i][6],
          c: qData[i][7],
          d: qData[i][8]
        });
      }
    }
    
    if (filteredQuestions.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Ngân hàng đề chưa có câu hỏi cho môn này!"})).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Trộn ngẫu nhiên câu hỏi
    filteredQuestions.sort(function() { return 0.5 - Math.random() });
    var examQuestions = filteredQuestions.slice(0, examConfig.soCau);
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      config: { thoiGian: examConfig.thoiGian, monHoc: examConfig.monHoc },
      questions: examQuestions
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  // 2. LẤY DANH SÁCH MÔN HỌC ĐỂ HIỂN THỊ LÊN TRANG CHỦ
  if (action == "getSubjects") {
    var configSheet = sheet.getSheetByName("CaiDatDeThi");
    var configData = configSheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < configData.length; i++) {
      if(configData[i][0]) list.push(configData[i][0]);
    }
    return ContentService.createTextOutput(JSON.stringify({status: "success", subjects: list})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var params = JSON.parse(e.postData.contents);
  
  if (params.action == "submitExam") {
    var dbSheet = sheet.getSheetByName("NganHangDe");
    var data = dbSheet.getDataRange().getValues();
    var userAnswers = params.answers; 
    
    var score = 0;
    var total = 0;
    
    var correctAnswers = {};
    var qTypes = {};
    for (var i = 1; i < data.length; i++) {
      correctAnswers[data[i][0]] = data[i][9].toString().trim().toLowerCase(); 
      qTypes[data[i][0]] = data[i][2];
    }
    
    for (var qId in userAnswers) {
      total++;
      var userAns = userAnswers[qId].toString().trim().toLowerCase();
      var keyAns = correctAnswers[qId];
      
      if (qTypes[qId] === "FILL_BLANK") {
        // Hỗ trợ kiểm tra điền khuyết có nhiều đáp án đúng (cách nhau bởi dấu |)
        var multipleKeys = keyAns.split("|");
        if (multipleKeys.indexOf(userAns) !== -1) {
          score++;
        }
      } else {
        // Kiểm tra trắc nghiệm thông thường
        if (userAns === keyAns) score++;
      }
    }
    
    var resultSheet = sheet.getSheetByName("KetQua");
    var diemSo = total > 0 ? (score / total) * 10 : 0;
    resultSheet.appendRow([new Date(), params.studentName, params.monHoc, score, total, diemSo.toFixed(2)]);
    
    return ContentService.createTextOutput(JSON.stringify({status: "success", score: score, total: total}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}
