/**
 * HỆ THỐNG TỰ ĐỘNG HÓA ĐIỀU PHỐI LEADS & GIÁM SÁT SLA 4H/24H - CÔNG TY NAM VIỆT
 * Nền tảng: Google Sheets + Google Apps Script + Telegram Bot / Gmail (Miễn phí 100%)
 * Hướng dẫn: Vào File Google Sheets Tổng -> Tiện ích mở rộng -> Apps Script -> Dán toàn bộ code này.
 */

// ================= CẤU HÌNH HỆ THỐNG =================
const CONFIG = {
  TELEGRAM_BOT_TOKEN: "NHAP_TOKEN_BOT_TELEGRAM_CUA_BAN", // Lấy từ @BotFather
  TPKD_TELEGRAM_ID: "NHAP_CHAT_ID_TRUONG_PHONG",
  TPKD_EMAIL: "tpkd@namviet.vn",
  VIP_BUDGET_THRESHOLD: 500000000, // 500 triệu VNĐ
  MAX_BACKLOG_CAP: 30,             // Ngưỡng trần 30 leads chưa gọi/NVKD
  SLA_WARN_HOURS: 4,               // Nhắc nhở sau 4 giờ
  SLA_RECLAIM_HOURS: 24            // Thu hồi sau 24 giờ
};

/**
 * HÀM CHÍNH: Chạy tự động mỗi 15 phút (Cài đặt trong Triggers)
 * Thực hiện trọn vẹn 4 bước AI/Logic + Đồng bộ 8 File NVKD
 */
function runLeadAutomation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsLeads = ss.getSheetByName("Leads_Master");
  const wsMatrix = ss.getSheetByName("Ma_Tran_8_NVKD");
  const wsBlack = ss.getSheetByName("Blacklist_CSKH");

  // 1. Đồng bộ kết quả gọi điện từ 8 file con của NVKD về File Tổng trước
  syncBackFromSalesSheets(wsLeads, wsMatrix);

  const leadsData = wsLeads.getDataRange().getValues();
  const matrixData = wsMatrix.getDataRange().getValues();
  const blackData = wsBlack.getDataRange().getValues();

  // Tạo từ điển Blacklist / CSKH cũ (Key: SĐT hoặc Email)
  const blackMap = {};
  for (let i = 1; i < blackData.length; i++) {
    const phone = String(blackData[i][0]).trim();
    const email = String(blackData[i][1]).trim().toLowerCase();
    const type = blackData[i][2]; // "Blacklist" | "Khách cũ" | "Khiếu nại CSKH"
    const owner = blackData[i][3];
    if (phone) blackMap[phone] = { type, owner };
    if (email) blackMap[email] = { type, owner };
  }

  // Đếm Backlog hiện tại chưa gọi của 8 NVKD & Lưu danh sách tỉnh/ngành
  const nvList = [];
  const existingKeys = {}; // Phát hiện trùng MST / Tên miền / SĐT

  for (let i = 1; i < matrixData.length; i++) {
    nvList.push({
      row: i + 1,
      idAndName: matrixData[i][0],          // VD: "NV01 - Nguyễn Văn An"
      pkg: String(matrixData[i][1]),        // VD: "PROD_RETAIL"
      provinces: String(matrixData[i][2]),  // Danh sách tỉnh phụ trách (đã chia rõ Miền Trung)
      tgId: String(matrixData[i][3]),       // Telegram Chat ID
      email: String(matrixData[i][4]),      // Email NVKD
      backlog: 0
    });
  }

  // Đếm backlog thực tế từ Leads_Master & Lập chỉ mục trùng lặp
  for (let i = 1; i < leadsData.length; i++) {
    const id = leadsData[i][0];
    const mst = String(leadsData[i][3]).trim();
    const phone = String(leadsData[i][4]).trim();
    const status = leadsData[i][12];
    const owner = leadsData[i][11];

    if (status === "ĐÃ PHÂN CÔNG") {
      const nv = nvList.find(n => n.idAndName === owner);
      if (nv) nv.backlog++;
    }
    if (status !== "ĐÃ GỘP HỒ SƠ" && id) {
      if (mst && mst !== "CHƯA XÁC ĐỊNH") existingKeys[mst] = id;
      if (phone) existingKeys[phone] = id;
    }
  }

  const now = new Date();

  // Duyệt từng dòng Lead trong Leads_Master
  for (let r = 1; r < leadsData.length; r++) {
    const rowIdx = r + 1;
    const leadId = leadsData[r][0];
    const repName = leadsData[r][1];
    const company = leadsData[r][2];
    const mst = String(leadsData[r][3]).trim();
    const phone = String(leadsData[r][4]).trim();
    const email = String(leadsData[r][5]).trim().toLowerCase();
    const city = String(leadsData[r][6]).trim();
    const pkg = String(leadsData[r][8]).trim();
    const budget = Number(leadsData[r][9]) || 0;
    const priority = String(leadsData[r][10]).trim();
    let owner = leadsData[r][11];
    let status = leadsData[r][12];
    let assignedAt = leadsData[r][14];   // Cột O: Gio_Giao_Viec
    let firstCallAt = leadsData[r][15];  // Cột P: Gio_Goi_Dau
    let warnSent = leadsData[r][18];     // Cột S: Da_Gui_Nhac_4h

    // ---------------- XỬ LÝ LEAD MỚI CHƯA PHÂN LOẠI ----------------
    if (status === "MỚI" || status === "") {
      
      // BƯỚC 1: KIỂM TRA BLACKLIST, CSKH CŨ & TRÙNG LẶP
      const historyHit = blackMap[phone] || blackMap[email];
      if (historyHit) {
        if (historyHit.type === "Blacklist") {
          wsLeads.getRange(rowIdx, 13).setValue("TỪ CHỐI - BLACKLIST");
          wsLeads.getRange(rowIdx, 14).setValue("Do Not Contact - Đóng tự động.");
          continue;
        } else if (historyHit.type === "Khiếu nại CSKH") {
          wsLeads.getRange(rowIdx, 12).setValue("CHỜ TPKD CHỈ ĐỊNH");
          wsLeads.getRange(rowIdx, 13).setValue("CẦN TPKD DUYỆT");
          wsLeads.getRange(rowIdx, 14).setValue("Cảnh báo CSKH: Từng khiếu nại. Trình TPKD.");
          sendAlert(CONFIG.TPKD_TELEGRAM_ID, CONFIG.TPKD_EMAIL, `🔍 CẢNH BÁO CSKH: Lead ${leadId} (${company}) từng khiếu nại, cần TPKD duyệt!`);
          continue;
        } else if (historyHit.type === "Khách cũ" && historyHit.owner) {
          assignLeadToNV(wsLeads, rowIdx, historyHit.owner, now, "Khách cũ Upsell - Giao thẳng NV chăm sóc.");
          continue;
        }
      }

      // Kiểm tra trùng MST hoặc SĐT với Lead trước đó
      const masterId = existingKeys[mst] || existingKeys[phone];
      if (masterId && masterId !== leadId) {
        wsLeads.getRange(rowIdx, 11).setValue("0 - ĐÃ GỘP");
        wsLeads.getRange(rowIdx, 12).setValue("Theo Master " + masterId);
        wsLeads.getRange(rowIdx, 13).setValue("ĐÃ GỘP HỒ SƠ");
        wsLeads.getRange(rowIdx, 14).setValue(`Gộp vào hồ sơ ${masterId}. Lưu ${repName} làm liên hệ phụ.`);
        continue;
      }

      // BƯỚC 2: ĐẢO LUẬT CHECK VIP >= 500 TRIỆU TRƯỚC KHI LỌC PHẠM VI SMB
      if (budget >= CONFIG.VIP_BUDGET_THRESHOLD) {
        wsLeads.getRange(rowIdx, 11).setValue("TPKD DUYỆT (VIP)");
        wsLeads.getRange(rowIdx, 12).setValue("CHỜ TPKD CHỈ ĐỊNH");
        wsLeads.getRange(rowIdx, 13).setValue("CẦN TPKD DUYỆT");
        wsLeads.getRange(rowIdx, 14).setValue(`VIP Ngân sách ${(budget/1e6).toFixed(0)}tr (≥500tr). Đẩy TPKD chỉ định KAM.`);
        sendAlert(CONFIG.TPKD_TELEGRAM_ID, CONFIG.TPKD_EMAIL, `🚨 CƠ HỘI VIP ≥500TR: ${leadId} - ${company} (${(budget/1e6).toFixed(0)} triệu). Cần TPKD chỉ định KAM ngay!`);
        continue;
      }

      // Kiểm tra ngoài phạm vi SMB (chỉ chạy với ngân sách < 500tr)
      if (pkg === "NGOÀI PHẠM VI") {
        wsLeads.getRange(rowIdx, 13).setValue("TỪ CHỐI / ĐÓNG LEAD");
        wsLeads.getRange(rowIdx, 14).setValue("Ngoài gói SMB. Tự động gửi email từ chối lịch sự.");
        continue;
      }

      // Kiểm tra thiếu thông tin
      if (pkg === "CHƯA XÁC ĐỊNH" || company === "CHƯA XÁC ĐỊNH") {
        wsLeads.getRange(rowIdx, 13).setValue("CHỜ BỔ SUNG THÔNG TIN");
        wsLeads.getRange(rowIdx, 14).setValue("Gửi Form khảo sát bổ sung tự động.");
        continue;
      }

      // Tìm NVKD phù hợp theo Gói SP & Tỉnh thành (đã quy hoạch Miền Trung)
      let targetNV = nvList.find(n => n.pkg.includes(pkg) && (n.provinces.includes("Toàn quốc") || n.provinces.includes(city)));
      if (!targetNV) {
        // Đa ngành hoặc chưa khớp rõ tỉnh -> Giao cho nhóm Đa ngành có tải thấp nhất
        const multiNVs = nvList.filter(n => n.pkg.includes("ĐA_NGANH"));
        multiNVs.sort((a, b) => a.backlog - b.backlog);
        targetNV = multiNVs[0];
      }

      // BƯỚC 3: KIỂM TRA TẢI TRỌNG TRẦN 30 LEADS (OVERFLOW)
      if (targetNV && targetNV.backlog >= CONFIG.MAX_BACKLOG_CAP) {
        if (priority !== "1 - CAO") {
          // Chuyển Lead TB/Thấp sang nhân sự ít việc nhất cùng/gần khu vực
          const availableNVs = nvList.filter(n => n.backlog < CONFIG.MAX_BACKLOG_CAP).sort((a, b) => a.backlog - b.backlog);
          if (availableNVs.length > 0) {
            targetNV = availableNVs[0];
          } else {
            wsLeads.getRange(rowIdx, 13).setValue("CẦN TPKD DUYỆT");
            wsLeads.getRange(rowIdx, 14).setValue("Toàn bộ team vượt trần 30 leads. Chờ TPKD điều phối.");
            continue;
          }
        }
      }

      if (targetNV) {
        assignLeadToNV(wsLeads, rowIdx, targetNV.idAndName, now, `Tư vấn giải pháp ${pkg} trong 24h.`);
        targetNV.backlog++;
      }
    }

    // ---------------- BƯỚC 4: GIÁM SÁT ĐỒNG HỒ SLA 4H & 24H ----------------
    if (status === "ĐÃ PHÂN CÔNG" && assignedAt instanceof Date && !firstCallAt) {
      const hoursElapsed = (now - assignedAt) / (1000 * 60 * 60);
      const nvInfo = nvList.find(n => n.idAndName === owner);

      if (hoursElapsed >= CONFIG.SLA_RECLAIM_HOURS) {
        // QUÁ HẠN 24H: Thu hồi về kho chung, ghi log vi phạm, báo đỏ TPKD
        wsLeads.getRange(rowIdx, 12).setValue("KHO CHUNG (QUÁ HẠN 24H)");
        wsLeads.getRange(rowIdx, 13).setValue("CẦN TPKD DUYỆT");
        wsLeads.getRange(rowIdx, 18).setValue(owner); // Cột R: NVKD_Bi_Thu_Hoi (Xếp hạng vi phạm SLA)
        sendAlert(CONFIG.TPKD_TELEGRAM_ID, CONFIG.TPKD_EMAIL,
          `🚨 CẢNH BÁO ĐỎ SLA 24H: Lead ${leadId} (${company}) giao cho ${owner} đã quá hạn 24h không gọi. Hệ thống đã THU HỒI về kho chung!`);
      } else if (hoursElapsed >= CONFIG.SLA_WARN_HOURS && warnSent !== "ĐÃ NHẮC 4H") {
        // QUÁ 4H: Gửi tin nhắn nhắc nhở NVKD
        if (nvInfo) {
          sendAlert(nvInfo.tgId, nvInfo.email,
            `⏰ NHẮC NHỞ SLA 4H: Lead ${leadId} (${company} - SĐT: ${phone}) đã giao quá 4 giờ chưa thấy cập nhật gọi. Vui lòng xử lý ngay!`);
        }
        wsLeads.getRange(rowIdx, 19).setValue("ĐÃ NHẮC 4H");
      }
    }
  }

  // Cuối cùng, đồng bộ danh sách sạch sang 8 File Google Sheets riêng của 8 NVKD
  syncToSalesSheets(wsLeads, wsMatrix);
}

function assignLeadToNV(ws, rowIdx, nvName, timestamp, actionNote) {
  ws.getRange(rowIdx, 12).setValue(nvName);
  ws.getRange(rowIdx, 13).setValue("ĐÃ PHÂN CÔNG");
  ws.getRange(rowIdx, 14).setValue(actionNote);
  ws.getRange(rowIdx, 15).setValue(timestamp); // Cột O: Gio_Giao_Viec
}

/**
 * Đẩy dữ liệu riêng từng NVKD sang File Sheet con của họ (Bảo mật tuyệt đối)
 * Cột F (cột 6) ở bảng Ma_Tran_8_NVKD lưu URL hoặc ID File Google Sheet riêng của từng NVKD
 */
function syncToSalesSheets(wsLeads, wsMatrix) {
  const leads = wsLeads.getDataRange().getValues();
  const matrix = wsMatrix.getDataRange().getValues();

  for (let i = 1; i < matrix.length; i++) {
    const nvName = matrix[i][0];
    const sheetId = String(matrix[i][5]).trim(); // ID File Google Sheet riêng của NVKD
    if (!sheetId) continue;

    try {
      const childSS = SpreadsheetApp.openById(sheetId);
      const childWs = childSS.getSheets()[0];
      
      // Lọc Lead của riêng NV này, sắp xếp 1-CAO lên trên cùng
      const myLeads = leads.slice(1).filter(r => r[11] === nvName && r[12] === "ĐÃ PHÂN CÔNG");
      myLeads.sort((a, b) => String(a[10]).localeCompare(String(b[10])) || (Number(b[9]) - Number(a[9])));

      const output = [["Mã Hồ Sơ", "Ưu Tiên", "Doanh Nghiệp", "Người Đại Diện", "SĐT", "Email", "Tỉnh/TP", "Gói SP", "Ngân Sách", "Hành Động", "Trạng Thái Gọi (Chọn 'Đã gọi')"]];
      myLeads.forEach(r => {
        output.push([r[0], r[10], r[2], r[1], r[4], r[5], r[6], r[8], r[9], r[13], r[15] ? "Đã gọi" : "Chưa gọi"]);
      });

      childWs.clearContents();
      childWs.getRange(1, 1, output.length, output[0].length).setValues(output);
    } catch (e) {
      Logger.log("Lỗi đồng bộ file NV " + nvName + ": " + e.message);
    }
  }
}

/**
 * Đọc ngược trạng thái "Đã gọi" từ 8 File Con về File Tổng để tắt đồng hồ SLA
 */
function syncBackFromSalesSheets(wsLeads, wsMatrix) {
  const leads = wsLeads.getDataRange().getValues();
  const matrix = wsMatrix.getDataRange().getValues();
  const calledIds = new Set();

  for (let i = 1; i < matrix.length; i++) {
    const sheetId = String(matrix[i][5]).trim();
    if (!sheetId) continue;
    try {
      const childWs = SpreadsheetApp.openById(sheetId).getSheets()[0];
      const rows = childWs.getDataRange().getValues();
      for (let j = 1; j < rows.length; j++) {
        if (String(rows[j][10]).trim().toLowerCase() === "đã gọi") {
          calledIds.add(rows[j][0]);
        }
      }
    } catch (e) {}
  }

  const now = new Date();
  for (let r = 1; r < leads.length; r++) {
    if (calledIds.has(leads[r][0]) && !leads[r][15]) {
      wsLeads.getRange(r + 1, 16).setValue(now); // Ghi nhận Gio_Goi_Dau -> Tắt đồng hồ SLA
      wsLeads.getRange(r + 1, 13).setValue("ĐÃ LIÊN HỆ");
    }
  }
}

/**
 * Hàm gửi thông báo đa kênh: Telegram Bot (tức thì) & Dự phòng Gmail
 */
function sendAlert(chatId, email, message) {
  // 1. Gửi Telegram nếu đã cấu hình Token và Chat ID
  if (CONFIG.TELEGRAM_BOT_TOKEN && !CONFIG.TELEGRAM_BOT_TOKEN.includes("NHAP_") && chatId) {
    try {
      const url = `https://api.telegram.org/bot${CONFIG.TELEGRAM_BOT_TOKEN}/sendMessage`;
      UrlFetchApp.fetch(url, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify({ chat_id: chatId, text: message })
      });
    } catch (e) {
      Logger.log("Telegram error: " + e.message);
    }
  }
  // 2. Gửi Email thông báo qua Gmail tích hợp sẵn
  if (email && email.includes("@")) {
    try {
      MailApp.sendEmail(email, "[Nam Việt CRM] Cảnh báo Điều phối Leads & SLA", message);
    } catch (e) {}
  }
}

