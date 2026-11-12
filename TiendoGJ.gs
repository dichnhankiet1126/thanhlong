/*
 * V3.1 11/11/2026 – P057 BEFORE HEAVY MAINTENANCE (CANDIDATE ONLY)
 * Source: deployed V3 package provided to the user; never auto-deploy.
 * No business fields, trigger definitions, receipts, script properties,
 * queue semantics, visitId or WebApp interfaces altered.
 */
/*
 * TOYOTA PHÚ MỸ HƯNG – ĐIỀU PHỐI EM/SCC
 * CLEANUP 11/11/2026 – CANDIDATE (CHƯA DEPLOY)
 * Nguồn: FULL Mã.gs P0.5.7.4 do người vận hành cung cấp.
 * Chỉ tinh gọn menu; 12 hàm Mã.gs cũ đã chuyển ra kho lưu trữ,
 * bỏ lịch sử banner đã được lưu lại. Không sửa logic nghiệp vụ.
 * CẤM tự đổi trigger, Script Properties, visitId, Data EM/SCC, giờ hoàn thành,
 * giờ ra, queue/receipt, WebApp /exec hoặc dữ liệu Production.
 */

/* Shared CVDV schema adapter. Does not alter Main/Data_CVDV layout. */
function mainQtNew_(sheet) {
  if(!sheet)throw Error('[MAIN_QT] Missing CVDV advisor sheet');
  var f=String(sheet.getRange('F2').getDisplayValue()||'').trim();
  if(f==='Tình trạng QT'){
    if(String(sheet.getRange('G2').getDisplayValue()||'').trim()!=='Số ngày lưu xưởng'||
       sheet.getMaxColumns()<11)throw Error('[MAIN_QT] Malformed 11-column advisor '+sheet.getName());
    return true;
  }
  if(f==='Số ngày lưu xưởng'){
    if(sheet.getMaxColumns()<10)throw Error('[MAIN_QT] Missing J '+sheet.getName());
    return false;
  }
  throw Error('[MAIN_QT] Unrecognized advisor F2 '+sheet.getName()+': '+f);
}
function mainQtWidth_(sheet){return mainQtNew_(sheet)?11:10;}
function mainQtCol_(sheet,legacyCol){return legacyCol>=6&&mainQtNew_(sheet)?legacyCol+1:legacyCol;}
function mainQtBusiness_(row){
  var cols=row.length===11?[0,1,2,3,4,5,7,9,10]:[0,1,2,3,4,6,8,9];
  return cols.every(function(c){var v=row[c];return v==null||String(v).trim()==='';});
}
function mainQtEqual_(a,b){
  if(a instanceof Date&&b instanceof Date)return a.getTime()===b.getTime();
  return String(a==null?'':a)===String(b==null?'':b);
}
function mainQtValidateRow_(source,target,sourceRow,targetRow,original){
  var qt=mainQtNew_(source),width=qt?11:10;
  if(qt!==mainQtNew_(target))return false;
  var a=source.getRange(sourceRow,1,1,width).getValues()[0];
  var b=target.getRange(targetRow,1,1,width).getValues()[0];
  var fields=qt?[0,1,2,3,4,5,7,9,10]:[0,1,2,3,4,6,8,9];
  if(!fields.every(function(i){return mainQtEqual_(a[i],original[i])&&mainQtEqual_(b[i],original[i]);}))
    return false;
  var aNotes=source.getRange(sourceRow,1,1,width).getNotes()[0];
  var bNotes=target.getRange(targetRow,1,1,width).getNotes()[0];
  return fields.every(function(i){return String(aNotes[i]||'')===String(bNotes[i]||'');});
}
function mainQtEnsureDropdown_(sheet,row){
  if(!mainQtNew_(sheet))return;
  sheet.getRange(row,6).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Chưa QT','Đã QT'],true).setAllowInvalid(false).build());
}
function TPMH_QT_MAIN_PreflightReadOnly(){
  var cv=SpreadsheetApp.openById(DELIVERY_PROMISE_SYNC_CFG.CVDV_SPREADSHEET_ID);
  var names=Object.values(DELIVERY_PROMISE_SYNC_CFG.ADVISOR_SHEET_MAP);
  var rows=names.map(function(n){var sh=cv.getSheetByName(n);
    if(!sh)throw Error('[MAIN_QT] Missing '+n);
    return {sheet:n,columns:mainQtWidth_(sh),promise:mainQtCol_(sh,7)};});
  var result={readOnly:true,mixed:rows.some(function(x){return x.columns===10;})&&
    rows.some(function(x){return x.columns===11;}),sheets:rows};
  console.log('[MAIN_QT_PREFLIGHT] '+JSON.stringify(result));return result;
}

const CFG = {
  SHEET_DISPATCH: 'Điều phối',
  SHEET_DATA_EM: 'Data EM',
  SHEET_DATA_SCC: 'Data SCC',
  SHEET_KT: 'Theo dõi KT',
  SHEET_KT_LAI: 'Theo dõi KT lại',
  SHEET_WAITING_DELIVERY: 'Xe chờ giao',
  SHEET_CVDV_DATA: 'Data_CVDV',
  FIRST_DATA_ROW: 6,
  WAITING_DELIVERY_FIRST_ROW: 6,
  CVDV_DATA_FIRST_ROW: 2,
  CVDV_RECONCILE_MAX_DAYS: 60,
  WASH_STATUS_VALUES: Object.freeze(['Chờ rửa', 'Đang rửa', 'Hoàn thành']),
  WD: {NGAY: 1, BIEN_SO: 2, CVDV: 3, MODEL: 4, NGAY_GIAO: 5, GIO_GIAO: 6, TRANG_THAI: 7, VISIT_ID: 8},
  CVDV_DATA: {NGAY: 1, BIEN_SO: 3, TINH_TRANG: 5, CVDV: 11},
  DP: {
    NGAY: 1,           // A
    BIEN_SO: 2,        // B
    CVDV: 3,           // C
    MODEL: 4,          // D
    NHOM_EM: 5,        // E
    NGUOI_KT: 6,       // F
    KTV1: 7,           // G - Nhóm SCC
    KTV2: 8,           // H - legacy/ẩn, không dùng cho Data SCC mới
    CV: 9,             // I
    CAP_BD: 10,        // J
    GTGT: 11,          // K
    SO_MON: 12,        // L
    GIO_NHAN: 13,      // M
    NGAY_GIAO: 14,     // N
    GIO_GIAO: 15,      // O
    TD_EM: 16,         // P
    TD_SCC: 17,        // Q
    TD_RUA_XE: 18,     // R
    GHI_CHU: 19,       // S
    MOC_TD_EM: 20,     // T
    MOC_TD_SCC: 21,    // U
    VISIT_ID: 22,      // V
    LAST_EM_STATUS: 23, // W
    LAST_SCC_STATUS: 24 // X
  },
  // Data SCC schema mới: A:N nghiệp vụ, O = visitId
  DATA_SCC: {
    NGAY: 1,
    BIEN_SO: 2,
    CVDV: 3,
    MODEL: 4,
    NGUOI_KT: 5,
    CONG_VIEC: 6,
    SO_MON: 7,
    NHOM_SCC: 8,
    GIO_NHAN: 9,
    NGAY_GIAO: 10,
    GIO_GIAO: 11,
    GIO_HOAN_THANH: 12,
    TIEN_DO_CUOI: 13,
    GHI_CHU: 14,
    VISIT_ID: 15
  }
};

/*********************************************************
 * V13.20.3 - XE CHỜ GIAO: NGÀY CHUẨN; HONDA CHI RA THEO CHUNG CU
 *********************************************************/
const WAITING_CANONICAL_CFG = Object.freeze({
  CANONICAL_LOOKBACK_DAYS: 30,
  HONDA_ADVISORS: Object.freeze(['lap', 'tien'])
});
/*********************************************************
 * V13.14 - ĐỒNG BỘ HẸN GIAO ĐIỀU PHỐI -> CVDV
 *
 * Nguồn:
 * - Điều phối A = Ngày tiếp nhận
 * - Điều phối B = Biển số
 * - Điều phối C = CVDV
 * - Điều phối N = Ngày giao
 * - Điều phối O = Giờ giao
 *
 * Đích:
 * - File CVDV -> đúng sheet theo cột C
 * - Match: Ngày tiếp nhận + BSX
 * - Ghi: cột G = HẸN GIAO
 *********************************************************/
const DELIVERY_PROMISE_SYNC_CFG = Object.freeze({
  CVDV_SPREADSHEET_ID:
    '1XtfqfG7l6ny1xioLcAeP9BsI88F0TDCDTa3VqHbMdpA',
  CVDV_FIRST_DATA_ROW: 3,
  CVDV_COL_DATE: 1,
  CVDV_COL_PLATE: 3,
  CVDV_COL_PROMISE: 7,
  LOG_SHEET: 'LOG_HEN_GIAO_CVDV',
  ADVISOR_SHEET_MAP: Object.freeze({
    'ha': 'Hà',
    'hai': 'Hải',
    'duy': 'Duy',
    'tinh': 'Tính',
    'minh': 'Minh',
    'thanh': 'Thành',
    'nha': 'Nhã',
    'phong': 'Phong',
    'cong': 'Công'
  })
});

/*********************************************************
 * V13.17 - CROSS-CHECK CỐ VẤN
 *
 * Nguồn chuẩn Cố vấn: Điều phối cột C; khi xe đã rời Điều phối thì dùng snapshot Xe chờ giao cột C.
 * Khóa nhận diện lượt: EXACT Ngày tiếp nhận + BSX.
 *
 * Sửa tự động khi an toàn:
 * - Lễ tân: sửa cột G CỐ VẤN nếu có đúng 1 match.
 * - CVDV: nếu có đúng 1 record toàn hệ thống nhưng nằm sai sheet,
 *   chuyển nguyên record A:J sang đúng sheet, giữ status/Hẹn giao/note.
 *
 * Không tự xử lý khi:
 * - Main không map được CVDV;
 * - Lễ tân có >1 match cùng Ngày + BSX;
 * - CVDV có >1 match cùng Ngày + BSX;
 * - Record thay đổi trong lúc re-check.
 *********************************************************/
const CVDV_CROSSCHECK_CFG = Object.freeze({
  RECEPTION_SPREADSHEET_ID:
    '10dxPgBNvoYZ4oDydtf1-dNsGLxfHxJiRJ5l7uoJIqso',
  RECEPTION_SHEET_NAME: 'Giờ ra vào',
  RECEPTION_FIRST_DATA_ROW: 2,
  RECEPTION_COL_PLATE: 2,
  RECEPTION_COL_ENTRY: 5,
  RECEPTION_COL_ADVISOR: 7,
  RECEPTION_LOOKBACK_ROWS: 2500,
  CVDV_FIRST_DATA_ROW: 3,
  CVDV_WIDTH: 10,
  PERIODIC_DAYS: 2,
  PERIODIC_MAX_ROWS: 100,
  LOG_SHEET: 'LOG_CVDV_CROSSCHECK'
});
/*********************************************************
 * CONCURRENCY / RETRY
 *********************************************************/
const CONCURRENCY_CFG = Object.freeze({
  DIRTY_PREFIX: 'TPMH_DIRTY_VISIT_',
  DIRTY_FALLBACK_PREFIX: 'TPMH_DIRTY_FALLBACK_',
  QUARANTINE_PREFIX: 'TPMH_DIRTY_QUARANTINE_',
  // V13.19.2: WebApp vẫn ưu tiên tuyệt đối. Khi WebApp rảnh, worker chạy
  // PHASE 0 clear-only trước queue để dọn ngay các visit đã append Xe chờ giao
  // nhưng bị ngắt trước compact Điều phối. Sau đó mới drain pending queue.
  WORKER_BATCH: 8,
  WORKER_TIME_BUDGET_MS: 22000,
  WORKER_LOCK_MS: 800,
  SAFE_ROW_COMPACT_MAX_MS: 3000,
  QUEUE_LOCK_MS: 3000,
  MAX_ATTEMPTS: 20,
  MANUAL_RECOVERY_BATCH: 4,
  MANUAL_RECOVERY_TIME_BUDGET_MS: 25000,
  // P0.2: Bật lại AUTO reconcile tăng dần 2 phút/lượt,
  // 8 dòng/chu kỳ, tối đa 2 recovery; không quét toàn bộ file.
  // Manual recovery/menu vẫn được giữ nguyên.
  AUTO_RECONCILE_ENABLED: true,
  RECONCILE_BATCH: 1,  // Menu recovery thủ công giữ nguyên
  RECONCILE_INTERVAL_MS: 2 * 60 * 1000,
  RECONCILE_LAST_RUN_KEY: 'TPMH_RECONCILE_LAST_RUN_MS',
  RECONCILE_CURSOR_KEY: 'TPMH_AUTO_RECOVERY_NEXT_ROW_P02',
  RECONCILE_MAX_SCAN_ROWS: 8,
  RECONCILE_MAX_REPAIRS: 2,
  RECONCILE_TIME_BUDGET_MS: 4000
});
const WEBAPP_PRIORITY_GATE = Object.freeze({
  PROP_UNTIL: 'TPMH_WEBAPP_CONFIRM_PRIORITY_UNTIL'
});
function isWebAppConfirmPriorityActive_() {
  const raw =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        WEBAPP_PRIORITY_GATE.PROP_UNTIL
      );
  const until = Number(raw || 0);
  if (!until) return false;
  if (until <= Date.now()) {
    // Tự dọn cờ hết hạn để không tạo trạng thái kẹt vĩnh viễn.
    try {
      PropertiesService
        .getScriptProperties()
        .deleteProperty(
          WEBAPP_PRIORITY_GATE.PROP_UNTIL
        );
    }
    catch (err) {}
    return false;
  }
  return true;
}
const WAITING_CVDV_GUARD = Object.freeze({
  UNSAFE_PREFIX: 'TPMH_WD_CVDV_UNSAFE_'
});
function waitingCvdvUnsafeKey_(visitId) {
  const id = clean_(visitId);
  return id ? WAITING_CVDV_GUARD.UNSAFE_PREFIX + id : '';
}

function clearWaitingVisitCvdvUnsafe_(visitId) {
  const key = waitingCvdvUnsafeKey_(visitId);
  if (!key) return;
  PropertiesService
    .getScriptProperties()
    .deleteProperty(key);
}
function isWaitingVisitCvdvUnsafe_(visitId) {
  const key = waitingCvdvUnsafeKey_(visitId);
  if (!key) return false;
  return !!PropertiesService
    .getScriptProperties()
    .getProperty(key);
}
function isWaitingCvdvGuardRecoverable_(visitId, duplicateVisitId, statusKey, decision) {
  return !!(
    clean_(visitId) &&
    duplicateVisitId !== true &&
    statusKey === 'cho giao' &&
    decision &&
    decision.safeToComplete === true &&
    decision.matchType === 'exact'
  );
}

/*********************************************************
 * MENU
 *********************************************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // Chỉ kiểm tra, không dùng các hàm sửa dữ liệu để thay thế audit.
  const menuKiemTra = ui.createMenu('🔎 Kiểm tra & đối soát')
    .addItem('🆔 Kiểm tra trùng visitId', 'kiemTraDuplicateVisitIdToanHeThong')
    .addItem('🔎 Audit Hoàn thành thiếu Data EM/SCC', 'p057CompletionGapAuditReadOnly')
    .addItem('📬 Xem trạng thái pending queue', 'xemTrangThaiPendingQueue')
    .addItem('🔎 Kiểm tra tự phục hồi P049', 'p049ProductionPreflightReadOnly');

  // Các thao tác có thể ghi dữ liệu; giữ tên hàm nguyên vẹn.
  const menuBaoTri = ui.createMenu('🛠 Bảo trì & khôi phục')
    .addItem('🔁 Recovery xe bị kẹt (timeout-safe)', 'dongBoLaiDongBiKet')
    .addItem('🧾 Khôi phục Hoàn thành có giờ xác nhận', 'p057RecoverSelectedCompletedWithOperatorProof')
    .addItem('🚪 Đối soát/đồng bộ xe đã ra cổng', 'dongBoTrangThaiCVDV')
    .addItem('⚡ Dọn xe đã Hoàn thành', 'donDongDaHoanThanh')
    .addSeparator()
    .addItem('🕒 [Dự phòng] Hẹn giao → CVDV', 'dongBoHenGiaoDongDangChon')
    .addItem('🔎 [Dự phòng] Kiểm tra CVDV dòng chọn', 'kiemTraCheoCvdvDongDangChon')
    .addSeparator()
    .addItem('🔤 Khóa biển số dạng Text', 'ensureAllPlateColumnsText_')
    .addItem('🔄 Đồng bộ dropdown Xe chờ giao', 'dongBoDropdownXeChoGiao')
    .addSeparator()
    .addItem('🙈 Chỉ hiện Data hôm nay', 'anDuLieuCu')
    .addItem('👁 Hiện toàn bộ Data', 'hienToanBoData');

  // Các hàm cài/gỡ trigger và công tắc P049 vẫn có thể được Admin
  // gọi trực tiếp trong Apps Script Editor khi bảo trì có kiểm soát.
  // Không đưa vào menu vận hành để tránh tạo/gỡ trigger ngoài ý muốn.
  ui.createMenu('🚗 Điều phối')
    .addItem('🔄 Đồng bộ lại dòng đang chọn', 'dongBoLaiDongDangChon')
    .addItem('⏱ Ghi giờ hiện tại', 'ghiGioHienTai')
    .addItem('🔎 Kiểm tra xe phát sinh chưa giao', 'p046KiemTraCongViecPhatSinh')
    .addItem('↩️ Xác nhận chưa giao – mở việc bổ sung', 'p046XacNhanXeChuaGiaoPhatSinhViec')
    .addItem('📊 Cập nhật thông tin tải xưởng', 'capNhatTaiXuong')
    .addSeparator()
    .addSubMenu(menuKiemTra)
    .addSubMenu(menuBaoTri)
    .addToUi();
}
/*********************************************************
 * TRIGGER CHÍNH
 *********************************************************/
function xuLyChinhSua(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const name = sh.getName();
  if (
    !isProtectedPlateSheet_(name) &&
    name !== CFG.SHEET_WAITING_DELIVERY
  ) {
    return;
  }

  // P0.4.4: Persist the edit event *before* waiting on ScriptLock.
  // This is an observed Apps Script event time, never an inferred
  // service-completion time. Only a real P/Q edit gets this evidence.
  let p044JournalSaved = false;
  const p044CriticalStatusEdit =
    name === CFG.SHEET_DISPATCH &&
    e.range.getLastRow() >= CFG.FIRST_DATA_ROW &&
    e.range.getColumn() <= CFG.DP.TD_SCC &&
    e.range.getLastColumn() >= CFG.DP.TD_EM;

  if (p044CriticalStatusEdit) {
    e.p044CapturedEventAt = new Date().toISOString();
    // P0.5.7: an independent SOURCE-BOUND receipt survives a queue failure.
    // This is recorded at the true onEdit event, NEVER at worker runtime.
    try { p057CaptureStatusReceiptFromEdit_(e); }
    catch (p057ReceiptErr) {
      console.error('[P057_RECEIPT_SAVE_FAILED] '+String(p057ReceiptErr.message||p057ReceiptErr));
    }
    try {
      p044JournalSaved = !!enqueueDirtyVisitFromEdit_(e);
      console.log('[P044_EARLY_EVENT] ' + JSON.stringify({
        at: e.p044CapturedEventAt,
        range: e.range.getA1Notation(),
        persisted: p044JournalSaved
      }));
    } catch (journalErr) {
      // The direct sync still gets one chance to complete. A later
      // failure must remain visible as an execution error, not be
      // hidden behind a false PASS.
      console.error('[P044_EARLY_EVENT_FAILED] ' + journalErr.message);
    }
  }

  // Do not overwrite a persisted completion-event timestamp by
  // enqueueing the same edit a second time after a later error.
  function p044EnqueueIfNeeded_(sourceError) {
    if (p044JournalSaved) return true;
    const queued = !!enqueueDirtyVisitFromEdit_(e, sourceError);
    if (queued) p044JournalSaved = true;
    return queued;
  }

  // Keep all existing P0.4.3 flow, WebApp priority and locks.
  if (isWebAppConfirmPriorityActive_()) {
    if (
      name === CFG.SHEET_DISPATCH ||
      name === CFG.SHEET_WAITING_DELIVERY
    ) {
      p044EnqueueIfNeeded_();
    }
    return;
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(2500)) {
    if (
      name === CFG.SHEET_DISPATCH ||
      name === CFG.SHEET_WAITING_DELIVERY
    ) {
      p044EnqueueIfNeeded_();
    }
    return;
  }
  try {
    assertVisitSetup_(sh.getParent());
    if (name === CFG.SHEET_WAITING_DELIVERY) {
      handleWaitingDeliveryEdit_(e);
      return;
    }
    if (name === CFG.SHEET_DISPATCH) {
      guardEditedDispatchIdentityP03_(sh, e);
      sanitizeCopiedVisitMetadata_(sh, e.range);
    }
    if (editTouchesColumn_(e.range, 2)) {
      normalizeEditedPlateRange_(e);
    }
    invalidateHistoryCacheForEdit_(e);
    if (name !== CFG.SHEET_DISPATCH) return;
    const first = Math.max(e.range.getRow(), CFG.FIRST_DATA_ROW);
    const last = e.range.getRow() + e.range.getNumRows() - 1;
    for (let row = last; row >= first; row--) {
      if (isWebAppConfirmPriorityActive_()) {
        try {
          p044EnqueueIfNeeded_();
        } catch (queueErr) {
          console.log('Không enqueue được edit khi nhường WebApp: ' +
            queueErr.message);
        }
        break;
      }
      try {
        xuLyMotDongDieuPhoi_(sh, row, e);
      } catch (rowErr) {
        if (!p0533IsPendingProofError_(rowErr)) throw rowErr;
        // One business-proof block must not abort other rows in a paste.
        // Never suppress unless a durable, visit-bound queue entry exists.
        const heldVisitId = clean_(sh.getRange(row, CFG.DP.VISIT_ID).getDisplayValue());
        if (!heldVisitId) throw rowErr;
        if (!p044JournalSaved) {
          const deferredJob = buildDirtyVisitJobForRow_(sh, name, row, e, rowErr);
          if (!deferredJob || !enqueueDirtyVisit_(deferredJob)) throw rowErr;
        }
        p0533LogProofHold_('onEdit', {
          sheet: name,
          rowHint: row,
          visitId: clean_(sh.getRange(row, CFG.DP.VISIT_ID).getDisplayValue()),
          plate: normalizePlate_(sh.getRange(row, CFG.DP.BIEN_SO).getDisplayValue())
        }, rowErr);
      }
    }
  } catch (err) {
    if (
      name === CFG.SHEET_DISPATCH ||
      name === CFG.SHEET_WAITING_DELIVERY
    ) {
      try {
        p044EnqueueIfNeeded_(err);
      } catch (queueErr) {
        console.log('Không enqueue được dirty visit sau lỗi onEdit: ' +
          queueErr.message);
      }
    }
    throw err;
  } finally {
    try {
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  }
}
/*********************************************************
 * PENDING VISIT QUEUE
 *
 * Mỗi visitId là một Script Property riêng -> cùng visit
 * enqueue nhiều lần chỉ overwrite một key, không tạo duplicate.
 *
 * Queue chỉ lưu "cần xử lý lại". Worker luôn đọc state HIỆN TẠI
 * từ Sheet, không tin event.value cũ.
 *********************************************************/
function dirtyVisitPropertyKey_(sheetName, visitId) {
  return (
    CONCURRENCY_CFG.DIRTY_PREFIX +
    encodeURIComponent(sheetName) +
    '_' +
    clean_(visitId)
  );
}
function dirtyFallbackPropertyKey_(job) {
  const raw = [
    job.sheet || '',
    job.rowHint || '',
    job.plate || '',
    job.dateKey || '',
    job.advisor || '',
    job.model || ''
  ].join('|');
  const digest =
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5,
      raw,
      Utilities.Charset.UTF_8
    )
      .map(function(b) {
        const v =
          (b < 0 ? b + 256 : b)
            .toString(16);
        return v.length === 1
          ? '0' + v
          : v;
      })
      .join('');
  return (
    CONCURRENCY_CFG.DIRTY_FALLBACK_PREFIX +
    digest
  );
}
function enqueueDirtyVisitFromEdit_(e, sourceError) {
  if (!e || !e.range) return false;
  const sh = e.range.getSheet();
  const name = sh.getName();
  if (
    name !== CFG.SHEET_DISPATCH &&
    name !== CFG.SHEET_WAITING_DELIVERY
  ) {
    return false;
  }
  const firstDataRow =
    name === CFG.SHEET_DISPATCH
      ? CFG.FIRST_DATA_ROW
      : CFG.WAITING_DELIVERY_FIRST_ROW;
  const firstRow =
    Math.max(
      e.range.getRow(),
      firstDataRow
    );
  const lastRow =
    e.range.getRow() +
    e.range.getNumRows() -
    1;
  if (lastRow < firstRow) return false;
  let queued = 0;
  // Queue TẤT CẢ dòng bị ảnh hưởng, không chỉ dòng đầu của multi-row paste.
  for (let row = firstRow; row <= lastRow; row++) {
    const job =
      buildDirtyVisitJobForRow_(
        sh,
        name,
        row,
        e,
        sourceError
      );
    if (
      job &&
      enqueueDirtyVisit_(
        job
      )
    ) {
      queued++;
    }
  }
  return queued > 0;
}
function buildDirtyVisitJobForRow_(sh, name, row, e, sourceError) {
  const visitCol =
    name === CFG.SHEET_DISPATCH
      ? CFG.DP.VISIT_ID
      : CFG.WD.VISIT_ID;
  const plateCol =
    name === CFG.SHEET_DISPATCH
      ? CFG.DP.BIEN_SO
      : CFG.WD.BIEN_SO;
  const visitId =
    clean_(
      sh
        .getRange(
          row,
          visitCol
        )
        .getValue()
    );
  const plate =
    normalizePlate_(
      sh
        .getRange(
          row,
          plateCol
        )
        .getDisplayValue()
    );
  const job = {
    sheet: name,
    visitId: visitId,
    rowHint: row,
    plate: plate,
    colStart: e.range.getColumn(),
    numCols: e.range.getNumColumns(),
    oldValue:
      e.oldValue === undefined
        ? ''
        : String(e.oldValue),
    reason:
      sourceError
        ? 'onEdit-error: ' + sourceError.message
        : 'onEdit-lock-busy',
    updatedAt: new Date().toISOString(),
    // P0.4.2: mốc SỰ KIỆN người dùng sửa P/Q, không lấy giờ chạy worker.
    completionEventAt: (name === CFG.SHEET_DISPATCH &&
      e.range.getColumn() <= CFG.DP.TD_SCC &&
      e.range.getLastColumn() >= CFG.DP.TD_EM)
      ? (e.p044CapturedEventAt || new Date().toISOString()) : '',
    completionEventSides: (name === CFG.SHEET_DISPATCH) ?
      [CFG.DP.TD_EM, CFG.DP.TD_SCC].filter(function(col) {
        return editTouchesColumn_(e.range, col) &&
          isComplete_(sh.getRange(row, col).getDisplayValue());
      }) : [],
    attempts: 0,
    revision: Utilities.getUuid()
  };
  if (name === CFG.SHEET_DISPATCH) {
    const width =
      Math.max(
        CFG.DP.MODEL,
        CFG.DP.NGAY,
        CFG.DP.CVDV
      );
    const values =
      sh
        .getRange(
          row,
          1,
          1,
          width
        )
        .getValues()[0];
    const display =
      sh
        .getRange(
          row,
          1,
          1,
          width
        )
        .getDisplayValues()[0];
    job.dateKey =
      cvdvDateKey_(
        values[CFG.DP.NGAY - 1],
        sh.getParent()
      );
    job.advisor =
      normalizeKey_(
        display[CFG.DP.CVDV - 1]
      );
    job.model =
      normalizeKey_(
        display[CFG.DP.MODEL - 1]
      );
  }
  return job;
}
function queueMutex_() {
  // DocumentLock có thể không khả dụng nếu project đổi context.
  // UserLock không phù hợp vì worker/trigger có thể chạy dưới execution khác.
  // Dùng ScriptLock riêng theo critical section queue bằng Properties revision;
  // nếu main ScriptLock đang bận, queue producer KHÔNG thể lấy cùng lock.
  // Vì vậy queue dùng DocumentLock khi có; fallback về UserLock chỉ để serialize
  // producer cùng user, còn CAS revision ở worker bảo vệ khỏi xóa job mới.
  const documentLock =
    LockService.getDocumentLock();
  if (documentLock) {
    return documentLock;
  }
  const userLock =
    LockService.getUserLock();
  if (userLock) {
    return userLock;
  }
  throw new Error(
    'Không có LockService phù hợp để ghi pending queue.'
  );
}
/**
 * P0.4.3 – confirmation evidence is per side, never just one timestamp
 * for the entire visit. A later note/assignment edit must not erase an
 * earlier P/Q completion click while the worker is busy.
 */
function p043ConfirmationEvidenceMap_(job) {
  const out = {};
  if (!job) return out;
  const raw = job.p043CompletionEvents || {};
  [CFG.DP.TD_EM, CFG.DP.TD_SCC].forEach(function(col) {
    const t = clean_(raw[String(col)] || '');
    if (t) out[String(col)] = t;
  });
  const legacyTime = clean_(job.completionEventAt || job.p042CompletionEventAt || '');
  const legacySides = job.completionEventSides || job.p042CompletionEventSides || [];
  if (legacyTime && Array.isArray(legacySides)) legacySides.forEach(function(col) {
    if ((col === CFG.DP.TD_EM || col === CFG.DP.TD_SCC) &&
        !out[String(col)]) out[String(col)] = legacyTime;
  });
  return out;
}
function p043MergeCompletionEvidence_(fresh, previous) {
  if (!fresh || fresh.sheet !== CFG.SHEET_DISPATCH) return fresh;
  const inherited = p043ConfirmationEvidenceMap_(previous);
  const received = p043ConfirmationEvidenceMap_(fresh);
  const merged = Object.assign({}, inherited, received);
  [CFG.DP.TD_EM, CFG.DP.TD_SCC].forEach(function(col) {
    // An explicit edit clearing/reopening one status cancels its queued
    // completion event. A note edit leaves the evidence unchanged.
    const start = Number(fresh.colStart || 0);
    const end = start + Math.max(1, Number(fresh.numCols || 1)) - 1;
    if (!fresh.isNewWebAppInsert &&
        clean_(fresh.source || '') !== 'WEBAPP_FAST_COMMIT' &&
        start <= col && col <= end &&
        !(fresh.completionEventSides || []).includes(col) &&
        !received[String(col)]) delete merged[String(col)];
  });
  fresh.p043CompletionEvents = merged;
  return fresh;
}
function p043HasConfirmationEvidence_(e, statusCol) {
  if (!e) return false;
  const map = p043ConfirmationEvidenceMap_(e);
  return !!map[String(statusCol)];
}
function enqueueDirtyVisit_(job) {
  if (!job || !job.sheet) return false;
  const props =
    PropertiesService
      .getScriptProperties();
  const queueLock =
    queueMutex_();
  let locked = false;
  try {
    // Chờ lâu hơn để giảm khả năng mất producer.
    // Nếu vẫn timeout, throw để execution được ghi nhận là lỗi thay vì im lặng mất event.
    locked =
      queueLock.tryLock(
        CONCURRENCY_CFG.QUEUE_LOCK_MS
      );
    if (!locked) {
      throw new Error(
        'Không thể ghi pending queue sau ' +
        CONCURRENCY_CFG.QUEUE_LOCK_MS +
        ' ms.'
      );
    }
    const key =
      job.visitId
        ? dirtyVisitPropertyKey_(
            job.sheet,
            job.visitId
          )
        : dirtyFallbackPropertyKey_(
            job
          );
    const currentRaw =
      props.getProperty(
        key
      );
    if (currentRaw) {
      try {
        const current =
          JSON.parse(
            currentRaw
          );
        job.attempts =
          Number(
            current.attempts || 0
          );
        job.createdAt =
          current.createdAt ||
          job.updatedAt;
        p043MergeCompletionEvidence_(job, current);
      }
      catch (err) {}
    }
    // Initial completion event (without an older queue entry).
    p043MergeCompletionEvidence_(job, null);
    if (!job.createdAt) {
      job.createdAt =
        job.updatedAt ||
        new Date().toISOString();
    }
    // Mỗi enqueue luôn có revision mới.
    job.revision =
      Utilities.getUuid();
    props.setProperty(
      key,
      JSON.stringify(
        job
      )
    );
    return true;
  }
  finally {
    if (locked) {
      queueLock.releaseLock();
    }
  }
}
function loadDirtyVisitJobs_() {
  const scriptProps =
    PropertiesService
      .getScriptProperties();
  const props =
    scriptProps.getProperties();
  const jobs = [];
  Object.keys(
    props
  ).forEach(
    function(key) {
      if (
        key.indexOf(
          CONCURRENCY_CFG.DIRTY_PREFIX
        ) !== 0 &&
        key.indexOf(
          CONCURRENCY_CFG.DIRTY_FALLBACK_PREFIX
        ) !== 0
      ) {
        return;
      }
      try {
        const job =
          JSON.parse(
            props[key]
          );
        job._propertyKey =
          key;
        jobs.push(
          job
        );
      }
      catch (err) {
        const quarantineKey =
          CONCURRENCY_CFG.QUARANTINE_PREFIX +
          Utilities.getUuid();
        scriptProps.setProperty(
          quarantineKey,
          JSON.stringify({
            originalKey: key,
            raw: String(props[key] || ''),
            error: err.message,
            quarantinedAt:
              new Date().toISOString()
          })
        );
        scriptProps.deleteProperty(
          key
        );
      }
    }
  );
  jobs.sort(
    function(a, b) {
      return String(a.updatedAt || '')
        .localeCompare(
          String(b.updatedAt || '')
        );
    }
  );
  return jobs;
}
function removeDirtyVisitJob_(job) {
  if (!job || !job._propertyKey) return false;
  const props =
    PropertiesService
      .getScriptProperties();
  const queueLock =
    queueMutex_();
  let locked = false;
  try {
    locked =
      queueLock.tryLock(
        CONCURRENCY_CFG.QUEUE_LOCK_MS
      );
    if (!locked) {
      return false;
    }
    const currentRaw =
      props.getProperty(
        job._propertyKey
      );
    if (!currentRaw) {
      return true;
    }
    let current;
    try {
      current =
        JSON.parse(
          currentRaw
        );
    }
    catch (err) {
      return false;
    }
    // CAS: chỉ xóa nếu property vẫn là revision worker đã đọc.
    // Nếu onEdit vừa enqueue revision mới, giữ job mới lại.
    if (
      clean_(current.revision) !==
      clean_(job.revision)
    ) {
      return false;
    }
    props.deleteProperty(
      job._propertyKey
    );
    return true;
  }
  finally {
    if (locked) {
      queueLock.releaseLock();
    }
  }
}
function updateDirtyVisitAttempt_(job, err) {
  if (!job || !job._propertyKey) return false;
  const props =
    PropertiesService
      .getScriptProperties();
  const queueLock =
    queueMutex_();
  let locked = false;
  try {
    locked =
      queueLock.tryLock(
        CONCURRENCY_CFG.QUEUE_LOCK_MS
      );
    if (!locked) {
      return false;
    }
    const currentRaw =
      props.getProperty(
        job._propertyKey
      );
    if (!currentRaw) {
      return true;
    }
    let current;
    try {
      current =
        JSON.parse(
          currentRaw
        );
    }
    catch (parseErr) {
      return false;
    }
    // Nếu producer vừa ghi revision mới, worker cũ không được overwrite.
    if (
      clean_(current.revision) !==
      clean_(job.revision)
    ) {
      return false;
    }
    job.attempts =
      Number(
        job.attempts || 0
      ) + 1;
    job.updatedAt =
      new Date().toISOString();
    job.lastError =
      err
        ? String(err.message || err)
        : '';
    const cleanJob =
      Object.keys(job)
        .filter(function(key) {
          return key !== '_propertyKey';
        })
        .reduce(function(out, key) {
          out[key] = job[key];
          return out;
        }, {});
    if (
      job.attempts >=
      CONCURRENCY_CFG.MAX_ATTEMPTS
    ) {
      const quarantineKey =
        CONCURRENCY_CFG.QUARANTINE_PREFIX +
        encodeURIComponent(
          job.sheet || ''
        ) +
        '_' +
        encodeURIComponent(
          job.visitId ||
          job.plate ||
          Utilities.getUuid()
        );
      props.setProperty(
        quarantineKey,
        JSON.stringify(
          Object.assign(
            {},
            cleanJob,
            {
              quarantinedAt:
                new Date().toISOString()
            }
          )
        )
      );
      props.deleteProperty(
        job._propertyKey
      );
      return true;
    }
    // attempt update cũng là revision mới.
    cleanJob.revision =
      Utilities.getUuid();
    props.setProperty(
      job._propertyKey,
      JSON.stringify(
        cleanJob
      )
    );
    return true;
  }
  finally {
    if (locked) {
      queueLock.releaseLock();
    }
  }
}
/** P0.5.3.3: Only the two explicit MISSING PROOF gates are deferrable.
 * Do not mask mismatch, invalid time, sheet/lock/API failures or validation bugs.
 */
function p0533PendingProofCode_(err) {
  const message = String(err && (err.message || err) || '');
  const code = String(err && err.code || message.split(':')[0]);
  return code === 'P043_PENDING_COMPLETION_EVENT_REQUIRED' ||
    code === 'P04_COMPLETION_WITHOUT_STORED_WORK' ? code : '';
}
function p0533IsPendingProofError_(err) {
  return !!p0533PendingProofCode_(err);
}
function p0533LogProofHold_(source, job, err) {
  console.warn('[P0533_COMPLETION_PROOF_PENDING] ' + JSON.stringify({
    source: source,
    reason: p0533PendingProofCode_(err),
    sheet: String(job && job.sheet || ''),
    rowHint: Number(job && job.rowHint || 0),
    plate: String(job && job.plate || ''),
    visitId: String(job && job.visitId || ''),
    detail: String(err && (err.message || err) || '').slice(0, 480)
  }));
}
/** CAS update under the existing queue mutex; no extra trigger/property key.
 * Preserve the current job rather than burning MAX_ATTEMPTS for a business hold.
 * A new user event replaces this revision; periodic worker rechecks at
 * 2/4/8/16/30-minute intervals, without repeating warning spam.
 */
function p0533DeferDirtyVisit_(job, err) {
  if (!p0533IsPendingProofError_(err) || !job || !job._propertyKey ||
      !clean_(job.visitId) || !clean_(job.plate)) return false;
  const lock = queueMutex_();
  if (!lock.tryLock(CONCURRENCY_CFG.QUEUE_LOCK_MS)) return false;
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty(job._propertyKey);
    if (!raw) return 'gone'; // Already resolved/removed.
    const current = JSON.parse(raw);
    if (clean_(current.revision) !== clean_(job.revision)) {
      // A newer edit superseded this worker's evidence snapshot.
      return 'superseded';
    }
    const now = new Date();
    const firstHold = current.p0533Deferred !== true;
    const checks = firstHold ? 1 : Math.min(1000, Number(current.p0533Checks || 0) + 1);
    const delayMs = Math.min(1800000, 120000 * Math.pow(2, Math.min(4, checks - 1)));
    current.p0533Deferred = true;
    current.p0533Checks = checks;
    current.p0533FirstBlockedAt = current.p0533FirstBlockedAt || now.toISOString();
    current.p0533LastBlockedAt = now.toISOString();
    current.p0533Reason = p0533PendingProofCode_(err);
    current.p0533Detail = String(err.message || err).slice(0, 480);
    current.p0533NextCheckAt = now.getTime() + delayMs;
    current.updatedAt = now.toISOString();
    current.revision = Utilities.getUuid();
    props.setProperty(job._propertyKey, JSON.stringify(current));
    return firstHold ? 'held-first' : 'held-repeat';
  } finally {
    lock.releaseLock();
  }
}
/** Read-only audit: neither getScriptProperties nor JSON.parse modifies queues. */
function p0533ProofDeferredAuditReadOnly() {
  const all = PropertiesService.getScriptProperties().getProperties();
  const now = Date.now();
  const result = {version: 'P0.5.3.3', readOnly: true,
    deferred: [], due: 0, waiting: 0, invalid: 0};
  Object.keys(all).forEach(function(key) {
    if (key.indexOf(CONCURRENCY_CFG.DIRTY_PREFIX) !== 0 &&
        key.indexOf(CONCURRENCY_CFG.DIRTY_FALLBACK_PREFIX) !== 0) return;
    try {
      const job = JSON.parse(all[key]);
      if (job.p0533Deferred !== true) return;
      const due = Number(job.p0533NextCheckAt || 0) <= now;
      if (due) result.due++; else result.waiting++;
      result.deferred.push({sheet: job.sheet, plate: job.plate,
        visitId: job.visitId, reason: job.p0533Reason,
        firstBlockedAt: job.p0533FirstBlockedAt,
        nextCheckAt: job.p0533NextCheckAt, checks: job.p0533Checks, due: due});
    } catch (ignored) { result.invalid++; }
  });
  console.log('[P0533_AUDIT_READ_ONLY] ' + JSON.stringify(result));
  return result;
}
function resolveDispatchRowsByVisitId_(sh, visitId) {
  const id =
    clean_(
      visitId
    );
  if (!id) return [];
  const last =
    sh.getLastRow();
  if (
    last <
    CFG.FIRST_DATA_ROW
  ) {
    return [];
  }
  const ids =
    sh
      .getRange(
        CFG.FIRST_DATA_ROW,
        CFG.DP.VISIT_ID,
        last - CFG.FIRST_DATA_ROW + 1,
        1
      )
      .getValues();
  const rows = [];
  ids.forEach(
    function(value, index) {
      if (
        clean_(
          value[0]
        ) === id
      ) {
        rows.push(
          CFG.FIRST_DATA_ROW +
          index
        );
      }
    }
  );
  return rows;
}
function resolveFallbackDispatchRow_(sh, job) {
  const last =
    sh.getLastRow();
  if (
    last <
    CFG.FIRST_DATA_ROW
  ) {
    return {
      rows: [],
      reason: 'empty'
    };
  }
  const width =
    Math.max(
      CFG.DP.MODEL,
      CFG.DP.NGAY,
      CFG.DP.BIEN_SO,
      CFG.DP.CVDV
    );
  const values =
    sh
      .getRange(
        CFG.FIRST_DATA_ROW,
        1,
        last - CFG.FIRST_DATA_ROW + 1,
        width
      )
      .getValues();
  const display =
    sh
      .getRange(
        CFG.FIRST_DATA_ROW,
        1,
        last - CFG.FIRST_DATA_ROW + 1,
        width
      )
      .getDisplayValues();
  const matches = [];
  for (let i = 0; i < values.length; i++) {
    const plate =
      normalizePlate_(
        display[i][CFG.DP.BIEN_SO - 1]
      );
    if (
      !plate ||
      plate !== job.plate
    ) {
      continue;
    }
    if (
      job.dateKey &&
      cvdvDateKey_(
        values[i][CFG.DP.NGAY - 1],
        sh.getParent()
      ) !== job.dateKey
    ) {
      continue;
    }
    if (
      job.advisor &&
      normalizeKey_(
        display[i][CFG.DP.CVDV - 1]
      ) !== job.advisor
    ) {
      continue;
    }
    if (
      job.model &&
      normalizeKey_(
        display[i][CFG.DP.MODEL - 1]
      ) !== job.model
    ) {
      continue;
    }
    matches.push(
      CFG.FIRST_DATA_ROW +
      i
    );
  }
  return {
    rows: matches,
    reason:
      matches.length === 1
        ? 'unique'
        : matches.length === 0
          ? 'not-found'
          : 'ambiguous'
  };
}
function buildRetryEvent_(sh, row, job) {
  const maxCols =
    sh.getMaxColumns();
  const startCol =
    Math.max(
      1,
      Math.min(
        Number(job.colStart || CFG.DP.GHI_CHU),
        maxCols
      )
    );
  const numCols =
    Math.max(
      1,
      Math.min(
        Number(job.numCols || 1),
        maxCols - startCol + 1
      )
    );
  return {
    range:
      sh.getRange(
        row,
        startCol,
        1,
        numCols
      ),
    source:
      sh.getParent(),
    oldValue:
      job.oldValue,
    isRetryWorker:
      true,
    // V13.19: job từ WebApp phải đi qua đúng nhánh new insert
    // để đồng bộ hậu kỳ sau khi request đã trả SUCCESS.
    isNewWebAppInsert:
      !!job.isNewWebAppInsert,
    asyncSource:
      clean_(job.source || ''),
    p042CompletionEventAt: clean_(job.completionEventAt || ''),
    p042CompletionEventSides: job.completionEventSides || [],
    p043CompletionEvents: p043ConfirmationEvidenceMap_(job)
  };
}
/**
 * P0.5.8.1 read-only check: does a dirty job still point to a live plate/day?
 * No UUID allocation, no Data mutation, no implicit advisor correction.
 * Caller holds the existing Worker ScriptLock.
 */
function p058ActiveDirtyOwnerEvidence_(ss, sh, job) {
  if (!ss || !sh || !job) return [];
  const target = normalizePlate_(job.plate || '');
  if (!target) return [];
  const first = CFG.FIRST_DATA_ROW;
  const last = sh.getLastRow();
  if (last < first) return [];
  const values = sh.getRange(first, 1, last - first + 1, CFG.DP.VISIT_ID).getValues();
  const displayed = sh.getRange(first, 1, last - first + 1, CFG.DP.VISIT_ID).getDisplayValues();
  const evidence = [];
  for (let i = 0; i < values.length; i++) {
    const row = first + i;
    if (normalizePlate_(displayed[i][CFG.DP.BIEN_SO - 1]) !== target) continue;
    if (job.dateKey) {
      const dateKey = cvdvDateKey_(values[i][CFG.DP.NGAY - 1], ss);
      if (dateKey !== job.dateKey) continue;
    } else if (Number(job.rowHint || 0) !== row) {
      // Without a date, never conflate an unrelated visit with a stale job.
      continue;
    }
    evidence.push({row: row,
      visitId: clean_(values[i][CFG.DP.VISIT_ID - 1]),
      advisor: clean_(displayed[i][CFG.DP.CVDV - 1])});
  }
  return evidence;
}

function processDirtyVisit_(ss, job) {
  if (!job || !job.sheet) {
    return {
      done: true,
      reason: 'invalid-job'
    };
  }
  if (
    job.sheet ===
    CFG.SHEET_DISPATCH
  ) {
    const sh =
      ss.getSheetByName(
        CFG.SHEET_DISPATCH
      );
    if (!sh) {
      return {
        done: false,
        error:
          new Error(
            'Không tìm thấy Điều phối.'
          )
      };
    }
    let rows = [];
    if (job.visitId) {
      rows =
        resolveDispatchRowsByVisitId_(
          sh,
          job.visitId
        );
      if (rows.length === 0) {
        const stillActive = p058ActiveDirtyOwnerEvidence_(ss, sh, job);
        if (stillActive.length) {
          console.error('[P058_DIRTY_VISIT_ID_LOST_OR_REKEYED] ' +
            JSON.stringify({plate:job.plate, originalVisitId:job.visitId,
              candidates:stillActive}));
          return {done:false, error:new Error('P058_DIRTY_OWNER_UNRESOLVED: ' +
            job.plate + ' / oldId=' + job.visitId + ' / stillActive=' +
            JSON.stringify(stillActive))};
        }
        // Only a genuinely absent owner may be treated as left-dispatch.
        return {done:true, reason:'visit-left-dispatch'};
      }
      if (rows.length > 1) {
        return {
          done: false,
          error:
            new Error(
              'Pending visitId ' +
              job.visitId +
              ' đang có ' +
              rows.length +
              ' dòng Điều phối.'
            )
        };
      }
    }
    else {
      const fallback =
        resolveFallbackDispatchRow_(
          sh,
          job
        );
      rows =
        fallback.rows;
      if (rows.length === 0) {
        const stillActive = p058ActiveDirtyOwnerEvidence_(ss, sh, job);
        if (stillActive.length) {
          console.error('[P058_FALLBACK_IDENTITY_CHANGED] ' +
            JSON.stringify({plate:job.plate, oldAdvisor:job.advisor,
              candidates:stillActive}));
          return {done:false, error:new Error('P058_FALLBACK_OWNER_UNRESOLVED: ' +
            job.plate + ' / stillActive=' + JSON.stringify(stillActive))};
        }
        return {done:true, reason:'fallback-not-found'};
      }
      if (rows.length > 1) {
        return {
          done: false,
          error:
            new Error(
              'Pending fallback không xác định duy nhất ownership.'
            )
        };
      }
      const newVisitId =
        ensureVisitId_(
          sh,
          rows[0]
        );
      job.visitId =
        newVisitId;
    }
    const row =
      rows[0];
    const currentVisitId =
      clean_(
        sh
          .getRange(
            row,
            CFG.DP.VISIT_ID
          )
          .getValue()
      );
    if (
      job.visitId &&
      currentVisitId !==
      job.visitId
    ) {
      return {
        done: false,
        error:
          new Error(
            'Ownership thay đổi trước retry.'
          )
      };
    }
    xuLyMotDongDieuPhoi_(
      sh,
      row,
      buildRetryEvent_(
        sh,
        row,
        job
      )
    );
    return {
      done: true,
      reason: 'processed'
    };
  }
  if (
    job.sheet ===
    CFG.SHEET_WAITING_DELIVERY
  ) {
    const sh =
      waitingDeliverySheet_(
        ss
      );
    let rows = [];
    if (job.visitId) {
      rows =
        waitingDeliveryRowsByVisitId_(
          sh,
          job.visitId
        );
    }
    else {
      // Xe chờ giao chuẩn phải luôn có visitId.
      // Không fallback theo biển số vì row cũ có thể đã biến mất và biển số
      // có thể thuộc một lượt khác.
      return {
        done: true,
        reason: 'waiting-missing-visitId-skip'
      };
    }
    if (rows.length === 0) {
      return {
        done: true,
        reason: 'visit-left-waiting'
      };
    }
    if (rows.length > 1) {
      return {
        done: false,
        error:
          new Error(
            'Pending Xe chờ giao ownership không duy nhất.'
          )
      };
    }
    const row =
      rows[0];
    const currentVisitId =
      clean_(
        sh
          .getRange(
            row,
            CFG.WD.VISIT_ID
          )
          .getValue()
      );
    if (
      job.visitId &&
      currentVisitId !==
      job.visitId
    ) {
      return {
        done: false,
        error:
          new Error(
            'visitId Xe chờ giao thay đổi trước retry.'
          )
      };
    }
    if (
      isComplete_(
        sh
          .getRange(
            row,
            CFG.WD.TRANG_THAI
          )
          .getDisplayValue()
      )
    ) {
      clearAndCompactWaitingDeliveryRow_(
        sh,
        row,
        currentVisitId
      );
    }
    return {
      done: true,
      reason: 'processed-waiting'
    };
  }
  return {
    done: true,
    reason: 'unsupported-sheet'
  };
}
/*********************************************************
 * WORKER PENDING VISIT
 *********************************************************/
function shouldRunWorkerReconcile_() {
  // Script Properties có thể gần đầy bởi pending/receipt của WebApp.
  // Không dùng setProperty cho nhịp housekeeping để tránh làm worker lỗi.
  const key = 'TPMH_AUTO_RECONCILE_TICK_P02';
  try {
    const cache = CacheService.getScriptCache();
    if (cache.get(key)) return false;
    cache.put(key, '1', Math.max(60,
      Math.ceil(CONCURRENCY_CFG.RECONCILE_INTERVAL_MS / 1000)));
    return true;
  } catch (err) {
    // Nếu CacheService lỗi, không tự quét liên tục gây tranh WebApp.
    console.log('[AUTO_RECOVERY_P02][CACHE_SKIP] ' + err.message);
    return false;
  }
}
function workerJobPriority_(job) {
  if (!job) return 99;
  if (job.sheet === CFG.SHEET_WAITING_DELIVERY) return 0;
  if (job.sheet !== CFG.SHEET_DISPATCH) return 90;

  const start = Number(job.colStart || 0);
  const count = Math.max(1, Number(job.numCols || 1));
  const end = start + count - 1;
  // P:R là vùng tiến độ. Edit ở đây có khả năng làm xe đủ điều kiện rời Điều phối.
  if (start <= CFG.DP.TD_RUA_XE && end >= CFG.DP.TD_EM) return 1;
  if (job.isNewWebAppInsert || clean_(job.source || '') === 'WEBAPP_FAST_COMMIT') return 2;
  return 3;
}

function runPendingVisitWorker() {
  // Không tranh ScriptLock khi người dùng đang bấm XÁC NHẬN trên WebApp.
  // Pending job vẫn còn nguyên và worker sẽ xử lý ở lượt kế tiếp.
  if (isWebAppConfirmPriorityActive_()) {
    return {
      success: true,
      busy: true,
      webAppPriority: true,
      processed: 0,
      failed: 0,
      reconciled: 0
    };
  }
  const p0533NowMs = Date.now();
  const p0533AllJobs = loadDirtyVisitJobs_();
  let p0533HeldJobs = 0;
  const p0533ReadyJobs = [];
  // Phân loại pending/deferred trong một lượt quét, không thay đổi thứ tự.
  p0533AllJobs.forEach(function(job) {
    if (job.p0533Deferred !== true) {
      p0533ReadyJobs.push(job);
      return;
    }
    const nextCheckAt = Number(job.p0533NextCheckAt || 0);
    if (nextCheckAt > p0533NowMs) p0533HeldJobs++;
    else if (nextCheckAt <= p0533NowMs) p0533ReadyJobs.push(job);
    // NaN: giữ đúng hành vi cũ, không đưa job không hợp lệ vào lượt chạy.
  });
  const jobs =
    p0533ReadyJobs
      .sort(function(a, b) {
        const pa = workerJobPriority_(a);
        const pb = workerJobPriority_(b);
        if (pa !== pb) return pa - pb;
        const ta = Date.parse(a.updatedAt || a.createdAt || 0) || 0;
        const tb = Date.parse(b.updatedAt || b.createdAt || 0) || 0;
        return ta - tb;
      })
      .slice(
        0,
        CONCURRENCY_CFG.WORKER_BATCH
      );
  // Chỉ đánh dấu mốc reconcile sau khi lấy được ScriptLock, tránh
  // bỏ lỡ cả chu kỳ 2 phút khi WebApp đang chiếm ưu tiên.
  let doReconcile = false;
  // V13.19.2: KHÔNG early-return chỉ vì queue rỗng.
  // Mỗi lượt trigger vẫn thử ScriptLock ngắn để chạy PHASE 0 clear-only.
  // Lý do: một visit có thể đã append thành công sang Xe chờ giao nhưng
  // execution bị ngắt trước clearAndCompactDispatchRow_(). Khi đó queue có thể
  // đã hết nhưng Điều phối vẫn còn dòng. PHASE 0 tự dọn đúng case này.
  const lock =
    LockService
      .getScriptLock();
  if (
    !lock.tryLock(
      CONCURRENCY_CFG.WORKER_LOCK_MS
    )
  ) {
    return {
      success: true,
      busy: true,
      processed: 0,
      remaining: jobs.length
    };
  }
  let processed = 0;
  let failed = 0;
  let deferred = 0;
  try {
    // Priority có thể bật trong lúc worker đang chờ tryLock().
    // Sau khi vừa lấy được lock phải kiểm tra lại trước mọi mutation.
    if (isWebAppConfirmPriorityActive_()) {
      return {
        success: true,
        busy: true,
        webAppPriority: true,
        processed: 0,
        failed: 0,
        reconciled: 0,
        yieldedAfterLock: true
      };
    }
    const ss =
      laySpreadsheetWeb_();
    assertVisitSetup_(
      ss
    );
    doReconcile =
      CONCURRENCY_CFG.AUTO_RECONCILE_ENABLED === true &&
      shouldRunWorkerReconcile_();
    // V13.19.2 - PHASE 0: AUTO COMPACT-ONLY
    //
    // Nếu visitId đã có DUY NHẤT ở cả Điều phối và Xe chờ giao thì việc append
    // đã PASS trước đó. Chỉ clear đúng row Điều phối, tuyệt đối không gọi lại Data EM/SCC,
    // CVDV, Hẹn giao hay Tracking. Đây là công việc nhẹ và được ưu tiên trước queue.
    //
    // Hard time-budget vẫn áp dụng cho toàn worker; WebApp có thể giành P0 ở mọi
    // điểm an toàn trước mutation.
    const workerDeadline =
      Date.now() + Number(CONCURRENCY_CFG.WORKER_TIME_BUDGET_MS || 22000);

    // V3.1: prioritize proof-bound completion-receipt recovery before
    // AutoCompact (moveRows) and P0532 full Data ID scans. These may exceed
    // the 22-second cooperative Worker budget without throwing an error.
    // Original P057 validation, ScriptLock, WebApp priority, visitId and
    // completion timestamp requirements remain unchanged.
    // P0.5.7: bounded retry of MISSING COMPLETED Data only when a TRUE
    // onEdit receipt still matches the unique live visit and status.
    // Existing ScriptLock / WebApp priority / 22s budget remain unchanged.
    const p057StageStartedAt = Date.now();
    let p057Fast={checked:0,replayed:0,cleared:0,blocked:0,errors:[]};
    if (!isWebAppConfirmPriorityActive_() && Date.now()<workerDeadline) {
      try {
        p057Fast=p057RecoverCompletedReceiptUnderLock_(ss,
          Math.min(workerDeadline,Date.now()+TPMH_P057_CFG.WORKER_BUDGET_MS));
      } catch(p057Err) {
        p057Fast.errors.push(String(p057Err.message||p057Err));
        console.error('[P057_WORKER_REPLAY_FAILED] '+String(p057Err.message||p057Err));
      }
    }
    // Log on every worker tick, including the zero-receipt case; a completed
    // Apps Script execution alone does not show whether P057 was attempted.
    console.log('[P057_STAGE] '+JSON.stringify({
      checked:Number(p057Fast.checked||0),
      replayed:Number(p057Fast.replayed||0),
      blocked:Number(p057Fast.blocked||0),
      errors:(p057Fast.errors||[]).length,
      yieldedForBudget:!!p057Fast.yieldedForBudget,
      ms:Date.now()-p057StageStartedAt,
      remainingBudgetMs:Math.max(0,workerDeadline-Date.now())
    }));

    let quickCompact = {
      compacted: 0,
      plates: [],
      yieldedForWebApp: false
    };
    if (
      !isWebAppConfirmPriorityActive_() &&
      Date.now() < workerDeadline
    ) {
      const quickStart = Date.now();
      quickCompact =
        quickCompactDispatchAlreadyWaitingUnderLock_(
          ss,
          workerDeadline
        );
      console.log(
        '[PERF][AutoCompact] ' +
        JSON.stringify({
          compacted: Number(quickCompact.compacted || 0),
          plates: quickCompact.plates || [],
          packedRows: Number(quickCompact.packedRows || 0),
          packedBlocks: Number(quickCompact.packedBlocks || 0),
          packBlocked: !!quickCompact.packBlocked,
          packReason: clean_(quickCompact.packReason || ''),
          yieldedForWebApp: !!quickCompact.yieldedForWebApp,
          ms: Date.now() - quickStart
        })
      );
    }

    // P0.5.3.2: Bounded OPEN Data reconciliation is reserved BEFORE dirty jobs.
    // Before this change it was AFTER the queue plus P049/P053 stages, so a
    // sustained queue could starve recovery even with a 1-minute trigger.
    // Existing WebApp priority and the 22s worker deadline still apply.
    // Same visitId, assigned OPEN status only. No invented completion stamps.
    let p047Fast = { checked: 0, repaired: 0, skipped: 0, errors: [], busy: false };
    if (!isWebAppConfirmPriorityActive_() && Date.now() < workerDeadline) {
      const p047Start = Date.now();
      p047Fast = p047RestoreMissingOpenDataUnderLock_(
        ss, Math.min(workerDeadline, Date.now() + 8000)
      );
      console.log('[P0532_OPEN_DATA_EARLY] ' + JSON.stringify({
        checked: p047Fast.checked, repaired: p047Fast.repaired,
        skipped: p047Fast.skipped, errors: p047Fast.errors.length,
        busy: !!p047Fast.busy, ms: Date.now() - p047Start
      }));
    }

    // Nếu WebApp bật priority trong lúc PHASE 0 đang chuẩn bị, nhường ngay.
    // Các queue job vẫn giữ nguyên để lượt worker sau xử lý.
    if (
      isWebAppConfirmPriorityActive_() ||
      Date.now() >= workerDeadline
    ) {
      return {
        success: true,
        busy: false,
        webAppPriority: isWebAppConfirmPriorityActive_(),
        processed: 0,
        failed: 0,
        deferredPending: p0533HeldJobs,
        autoCompacted: Number(quickCompact.compacted || 0),
        autoCompactPlates: quickCompact.plates || [],
        packedRows: Number(quickCompact.packedRows || 0),
        packedBlocks: Number(quickCompact.packedBlocks || 0),
        reconciled: 0,
        yieldedAfterCompact: true,
        p057ReceiptChecked: Number(p057Fast.checked||0),
        p057ReceiptReplayed: Number(p057Fast.replayed||0),
        p057ReceiptBlocked: Number(p057Fast.blocked||0),
        p057ReceiptErrors: p057Fast.errors||[]
      };
    }

    // Queue rỗng vẫn có thể vừa compact được xe. Nếu không còn việc gì khác,
    // trả kết quả ngay sau PHASE 0 để giảm thời gian giữ ScriptLock.
    if (
      jobs.length === 0 &&
      !doReconcile
    ) {
      // P0.4.7: the queue can be empty after a missed/delayed producer.
      // Self-heal missing OPEN Data EM/SCC before the short-circuit.
      // This never alters statuses, timestamps, or completed records.
      const p049Fast = p049AutoTickUnderLock_(ss, workerDeadline);
      const p053Fast = p053OneQueueTickUnderLock_(ss, workerDeadline);
      // p047Fast already ran early; never perform the same pass twice.
      return {
        success: true,
        busy: false,
        processed: 0,
        failed: 0,
        deferredPending: p0533HeldJobs,
        autoCompacted: Number(quickCompact.compacted || 0),
        autoCompactPlates: quickCompact.plates || [],
        packedRows: Number(quickCompact.packedRows || 0),
        packedBlocks: Number(quickCompact.packedBlocks || 0),
        reconciled: 0,
        p047OpenRecovered: p047Fast.repaired,
        p047OpenChecked: p047Fast.checked,
        p053AutoStatus: p053Fast.status,
        p053AutoActivated: p053Fast.activated,
        p049AutoStatus: p049Fast.status,
        p049AutoRecovered: p049Fast.recovered,
        p049AutoPlate: p049Fast.plate || ''
      };
    }

    // Sau PHASE 0 mới xử lý pending queue, vẫn trong hard time-budget.
    for (let jobIndex = 0; jobIndex < jobs.length; jobIndex++) {
      if (
        isWebAppConfirmPriorityActive_() ||
        Date.now() >= workerDeadline
      ) {
        break;
      }
      const job = jobs[jobIndex];
      try {
        const workerJobStart =
          Date.now();

        const result =
          processDirtyVisit_(
            ss,
            job
          );

        console.log(
          '[PERF][RowWorker] ' +
          JSON.stringify({
            visitId: clean_(job.visitId || ''),
            plate: clean_(job.plate || ''),
            source: clean_(job.source || ''),
            attempt: Number(job.attempts || 0),
            ms: Date.now() - workerJobStart,
            done: !!result.done,
            reason: clean_(result.reason || '')
          })
        );

        if (result.done) {
          const removed =
            removeDirtyVisitJob_(
              job
            );
          if (removed) {
            processed++;
          }
        }
        else {
          updateDirtyVisitAttempt_(
            job,
            result.error ||
              new Error(
                result.reason ||
                'Pending chưa xử lý được.'
              )
          );
          failed++;
        }
      }
      catch (err) {
        if (
          err &&
          err.code === 'WEBAPP_PRIORITY_YIELD'
        ) {
          console.log(
            '[WORKER] Nhường WebApp, giữ nguyên job ' +
            clean_(job.visitId || job.plate || '')
          );
          break;
        }

        if (p0533IsPendingProofError_(err)) {
          const holdResult = p0533DeferDirtyVisit_(job, err);
          if (holdResult === 'held-first' || holdResult === 'held-repeat') {
            deferred++;
            // One warning per hold; repeated checks are recorded in queue
            // audit without flooding execution logs every few minutes.
            if (holdResult === 'held-first') p0533LogProofHold_('worker', job, err);
          } else if (holdResult === 'gone' || holdResult === 'superseded') {
            // No false pending warning for a resolved or newer event.
            console.log('[P0533_STALE_JOB_SKIPPED] ' +
              String(job.visitId || job.plate || '') + ' / ' + holdResult);
          } else {
            // Queue could not be checkpointed; this IS an execution issue.
            console.error('[P0533_DEFER_CHECKPOINT_FAILED] ' +
              String(job.visitId || job.plate || '') + ' / ' + err.message);
            updateDirtyVisitAttempt_(job, err);
            failed++;
          }
          continue;
        }
        updateDirtyVisitAttempt_(
          job,
          err
        );
        failed++;
      }
    }
    // After queued high-priority jobs, repair open Data that had no queue.
    // The existing bounded reconcile still handles completion with evidence.
    const p049Fast = p049AutoTickUnderLock_(ss, workerDeadline);
      const p053Fast = p053OneQueueTickUnderLock_(ss, workerDeadline);
    // p047Fast already ran early, ahead of queued work.
    let reconciliation = {
      processed: 0,
      errors: []
    };
    // P0.2: Tự đối soát tăng dần, bounded/cursor, không quét toàn bộ
    // những row đang làm và không tranh WebApp hơn 4 giây/lượt.
    if (
      doReconcile &&
      !isWebAppConfirmPriorityActive_() &&
      Date.now() < workerDeadline
    ) {
      reconciliation =
        reconcileDispatchStuckIncrementalUnderLock_(
          ss,
          workerDeadline
        );
    }
    return {
      success: failed === 0,
      busy: false,
      processed: processed,
      failed: failed,
      deferred: deferred,
      deferredPending: p0533HeldJobs + deferred,
      autoCompacted: Number(quickCompact.compacted || 0),
      autoCompactPlates: quickCompact.plates || [],
      packedRows: Number(quickCompact.packedRows || 0),
      packedBlocks: Number(quickCompact.packedBlocks || 0),
      reconciled:
        reconciliation.processed,
      recoveredData: Number(reconciliation.recoveredData || 0),
      movedToWaiting: Number(reconciliation.movedToWaiting || 0),
      reconcileChecked: Number(reconciliation.checked || 0),
      reconcileErrors:
        reconciliation.errors,
      p047OpenRecovered: p047Fast.repaired,
      p047OpenChecked: p047Fast.checked,
      p047OpenErrors: p047Fast.errors,
      p057ReceiptChecked: p057Fast.checked,
      p057ReceiptReplayed: p057Fast.replayed,
      p057ReceiptBlocked: p057Fast.blocked,
      p057ReceiptErrors: p057Fast.errors,
      p053AutoStatus: p053Fast.status,
        p053AutoActivated: p053Fast.activated,
        p049AutoStatus: p049Fast.status,
      p049AutoRecovered: p049Fast.recovered,
      p049AutoPlate: p049Fast.plate || ''
    };
  }
  finally {
    try {
      SpreadsheetApp.flush();
    }
    finally {
      lock.releaseLock();
    }
  }
}
/*********************************************************
 * P0.2 - RECONCILIATION SELF-HEAL (BOUNDED, FAIR, FAIL-CLOSED)
 *
 * Chạy SAU pending queue, dưới ScriptLock hiện có. Tối đa 8
 * dòng/chu kỳ, tối đa 2 lượt sửa/chu kỳ, 4 giây; cursor qua từng dòng
 * thay vì quét mãi từ dòng cuối. Không đổi visitId/dữ liệu lịch sử.
 *
 * 1) Xe đủ điều kiện giao nhưng không được chuyển -> chạy lại processor.
 * 2) Xe còn Điều phối và Data EM/SCC theo visitId bị thiếu ->
 *    chạy lại đồng bộ downstream từ dòng chủ sở hữu, không tự đặt
 *    Hoàn thành. Đây là đường cứu worker/onEdit bị bỏ sót.
 *
 * Không đụng CVDV/Lễ tân ở nhánh này và không tự xử lý mâu thuẫn.
 *********************************************************/
/**
 * P0.4.3 – inspect only the same visitId and exact plate. A completed
 * Data row cannot silently be downgraded by reconciliation.
 * Returns a reason or '' and never writes.
 */
function p043ReconcileSideReason_(ss, rec, active, isEM) {
  if (!active) return '';
  const sheetName = isEM ? CFG.SHEET_DATA_EM : CFG.SHEET_DATA_SCC;
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('P043_DATA_SHEET_MISSING: ' + sheetName);
  const row = findDataRow_(sh, rec);
  if (!row) return isEM ? 'MISSING_EM' : 'MISSING_SCC';
  if (normalizePlate_(sh.getRange(row, 2).getDisplayValue()) !==
      normalizePlate_(rec.plate)) {
    throw new Error('P043_CROSS_VISIT_PLATE_MISMATCH: ' + sheetName);
  }
  const dpStatus = isEM ? rec.emStatus : rec.sccStatus;
  const storedStatus = clean_(sh.getRange(row, 13).getDisplayValue());
  if (!clean_(dpStatus) && storedStatus) {
    throw new Error('P043_STATUS_CLEARED_REQUIRES_REVIEW: ' + sheetName +
      ' / ' + rec.visitId);
  }
  if (isComplete_(storedStatus) && !isComplete_(dpStatus)) {
    throw new Error('P043_COMPLETED_DATA_DOWNGRADE_BLOCKED: ' + sheetName +
      ' / ' + rec.visitId);
  }
  if (normalizeKey_(dpStatus) !== normalizeKey_(storedStatus)) {
    const committed = isEM ? rec.lastEMStatus : rec.lastSCCStatus;
    if (isComplete_(dpStatus) && !isComplete_(committed)) {
      throw new Error('P043_UNCONFIRMED_COMPLETION_TRANSITION: ' + sheetName +
        ' / ' + rec.visitId + '. Đợi pending event P/Q; không lấy T/U của Đang làm.');
    }
    return isEM ? 'STATUS_EM_MISMATCH' : 'STATUS_SCC_MISMATCH';
  }
  if (isComplete_(dpStatus) && !sh.getRange(row, 12).getValue()) {
    return isEM ? 'STAMP_EM_MISSING' : 'STAMP_SCC_MISSING';
  }
  return '';
}
function reconcileDispatchStuckIncrementalUnderLock_(ss, workerDeadline) {
  const sh = ss.getSheetByName(CFG.SHEET_DISPATCH);
  if (!sh) throw new Error('AUTO_RECONCILE: thiếu Điều phối.');
  const out = {
    processed: 0, checked: 0, recoveredData: 0, movedToWaiting: 0,
    errors: [], yieldedForWebApp: false
  };
  const first = CFG.FIRST_DATA_ROW;
  const last = sh.getLastRow();
  if (last < first) return out;
  const cache = CacheService.getScriptCache();
  const cursorKey = CONCURRENCY_CFG.RECONCILE_CURSOR_KEY;
  let stored = 0;
  try { stored = Number(cache.get(cursorKey) || 0); }
  catch (err) { console.log('[AUTO_RECOVERY_P02][CURSOR_READ] ' + err.message); }
  let cursor = Number.isInteger(stored) && stored >= first && stored <= last
    ? stored : last;
  const deadline = Math.min(
    Number(workerDeadline || Number.MAX_SAFE_INTEGER),
    Date.now() + CONCURRENCY_CFG.RECONCILE_TIME_BUDGET_MS
  );
  try {
    while (
      out.checked < CONCURRENCY_CFG.RECONCILE_MAX_SCAN_ROWS &&
      out.processed < CONCURRENCY_CFG.RECONCILE_MAX_REPAIRS &&
      Date.now() < deadline
    ) {
      if (isWebAppConfirmPriorityActive_()) {
        out.yieldedForWebApp = true;
        break;
      }
      const row = cursor;
      cursor = row <= first ? last : row - 1;
      out.checked++;
      const plate = readPlateText_(sh, row);
      if (!plate) continue;
      try {
        const rec = readDispatchRow_(sh, row);
        const id = clean_(rec.visitId);
        if (!id) {
          out.errors.push('AUTO_RECOVERY_MISSING_VISIT_ID row=' + row);
          continue;
        }
        const owners = resolveDispatchRowsByVisitId_(sh, id);
        if (owners.length !== 1 || owners[0] !== row) {
          out.errors.push('AUTO_RECOVERY_IDENTITY_CONFLICT row=' + row);
          continue;
        }
        const route = classifyWork_(rec);
        const ready = isReadyForWaitingDelivery_(rec, route);
        // P0.4.3: Theo dõi cả trạng thái và mốc hoàn thành, không
        // chỉ tồn tại Data. Đây là READ-ONLY preflight. Sai ownership
        // hoặc downgrade Hoàn thành -> Đang làm sẽ throw/fail-closed.
        const emReason = p043ReconcileSideReason_(ss, rec, route.em.active, true);
        const sccReason = p043ReconcileSideReason_(ss, rec, route.scc.active, false);
        const reason = emReason || sccReason;
        const missing = /MISSING_(EM|SCC)/.test(emReason + ' ' + sccReason);
        if (!ready && !reason) continue;
        if (Date.now() >= deadline || isWebAppConfirmPriorityActive_()) break;

        // Dùng lại processor chuẩn. Khi thiếu Data, giữ new-insert
        // mode để không dọn bản ghi cũ ngoài luồng một cách ngẫu nhiên.
        // Event giả chỉ chạm cột S, không kích hoạt normalize ngày,
        // biển số hoặc sửa giờ đang sử dụng.
        xuLyMotDongDieuPhoi_(sh, row, {
          range: sh.getRange(row, CFG.DP.GHI_CHU),
          source: ss,
          isRetryWorker: true,
          isNewWebAppInsert: missing,
          isReconciliation: true,
          isRecovery: true
        });
        SpreadsheetApp.flush();
        const remaining = resolveDispatchRowsByVisitId_(sh, id);
        if (remaining.length === 0) {
          const waiting = waitingDeliverySheet_(ss);
          if (waitingDeliveryRowsByVisitId_(waiting, id).length === 1) {
            out.processed++;
            out.movedToWaiting++;
          } else {
            out.errors.push('AUTO_RECOVERY_MOVED_NOT_VERIFIED plate=' + plate);
          }
        } else if (remaining.length === 1 && reason) {
          const after = readDispatchRow_(sh, remaining[0]);
          const emLeft = p043ReconcileSideReason_(ss, after, route.em.active, true);
          const sccLeft = p043ReconcileSideReason_(ss, after, route.scc.active, false);
          if (!emLeft && !sccLeft) {
            out.processed++;
            out.recoveredData++;
          } else {
            out.errors.push('P043_RECONCILE_INCOMPLETE plate=' + plate +
              ', EM=' + emLeft + ', SCC=' + sccLeft);
          }
        } else if (remaining.length > 1) {
          out.errors.push('AUTO_RECOVERY_POST_DUPLICATE plate=' + plate);
        }
      } catch (err) {
        out.errors.push('AUTO_RECOVERY_ROW_' + row + ': ' + String(err.message || err));
      }
    }
  } finally {
    // Cursor có thể lệch vài dòng sau moveRows(); chu kỳ sau sẽ
    // normalize về bounds mới. Giữ fairness dù một dòng lỗi mãi.
    try { cache.put(cursorKey, String(cursor), 21600); }
    catch (err) { console.log('[AUTO_RECOVERY_P02][CURSOR_WRITE] ' + err.message); }
  }
  if (out.checked || out.errors.length) {
    console.log('[AUTO_RECOVERY_P02] ' + JSON.stringify(out));
  }
  return out;
}

/*********************************************************
 * RECONCILIATION XE KẸT
 *********************************************************/
function reconcileDispatchStuckUnderLock_(ss, maxRowsToProcess, deadlineMs) {
  const sh =
    ss.getSheetByName(
      CFG.SHEET_DISPATCH
    );
  if (!sh) {
    throw new Error(
      'Không tìm thấy Điều phối.'
    );
  }
  const errors = [];
  let processed = 0;
  for (
    let row = sh.getLastRow();
    row >= CFG.FIRST_DATA_ROW;
    row--
  ) {
    if (
      (deadlineMs && Date.now() >= deadlineMs) ||
      processed >=
      Number(
        maxRowsToProcess ||
        CONCURRENCY_CFG.RECONCILE_BATCH
      )
    ) {
      break;
    }
    if (
      !readPlateText_(
        sh,
        row
      )
    ) {
      continue;
    }
    try {
      let rec =
        readDispatchRow_(
          sh,
          row
        );
      if (!rec.visitId) {
        continue;
      }
      const ownerRows =
        resolveDispatchRowsByVisitId_(
          sh,
          rec.visitId
        );
      if (
        ownerRows.length !== 1 ||
        ownerRows[0] !== row
      ) {
        errors.push(
          rec.plate +
          ': ownership visitId không duy nhất.'
        );
        continue;
      }
      const route =
        classifyWork_(
          rec
        );
      // Chỉ recovery xe mà business rule hiện tại xác nhận đã đủ điều kiện.
      if (
        !isReadyForWaitingDelivery_(
          rec,
          route
        )
      ) {
        continue;
      }
      const missing =
        validateEMCompletion_(
          rec,
          route
        )
          .concat(
            validateSCCCompletion_(
              rec,
              route
            )
          );
      if (missing.length) {
        errors.push(
          rec.plate +
          ': ' +
          uniqueArray_(
            missing
          ).join(', ')
        );
        continue;
      }
      /*****************************************************
       * RECOVERY DỮ LIỆU BỊ MẤT DO ONEDIT CŨ
       *
       * Không validateStoredDataBeforeClear_ ở đây nữa.
       * Gọi lại processor hiện hữu với event giả chạm cả P/Q.
       *
       * Processor sẽ:
       * 1. updateStatusStamp_ từ trạng thái hiện tại
       * 2. đọc lại rec
       * 3. syncVisitIdentity_
       * 4. sync Data EM / Data SCC
       * 5. sync tracking
       * 6. validateStoredDataBeforeClear_
       * 7. chỉ move + compact nếu dữ liệu đã đủ
       *
       * Vì vậy xe kẹt cũ có thể tự phục hồi mốc hoàn thành,
       * nhưng vẫn giữ nguyên toàn bộ business validation.
       *****************************************************/
      const visitIdBefore =
        rec.visitId;
      const fakeEvent = {
        range:
          sh.getRange(
            row,
            CFG.DP.TD_EM,
            1,
            CFG.DP.TD_SCC - CFG.DP.TD_EM + 1
          ),
        source:
          ss,
        isReconciliation:
          true,
        isRecovery:
          true
      };
      xuLyMotDongDieuPhoi_(
        sh,
        row,
        fakeEvent
      );
      SpreadsheetApp.flush();
      // Nếu visitId không còn ở Điều phối, processor đã move/compact thành công.
      const remainingRows =
        resolveDispatchRowsByVisitId_(
          sh,
          visitIdBefore
        );
      if (remainingRows.length === 0) {
        processed++;
        continue;
      }
      if (
        remainingRows.length !== 1
      ) {
        errors.push(
          rec.plate +
          ': visitId không duy nhất sau recovery.'
        );
        continue;
      }
      // Vẫn còn Điều phối: đọc lại và báo chính xác dữ liệu nào chưa thể phục hồi.
      const currentRow =
        remainingRows[0];
      rec =
        readDispatchRow_(
          sh,
          currentRow
        );
      const currentRoute =
        classifyWork_(
          rec
        );
      const storedErrors =
        validateStoredDataBeforeClear_(
          rec,
          currentRoute
        );
      if (storedErrors.length) {
        errors.push(
          rec.plate +
          ': ' +
          storedErrors.join(' | ')
        );
      }
    }
    catch (err) {
      errors.push(
        'Dòng ' +
        row +
        ': ' +
        err.message
      );
    }
  }
  return {
    processed: processed,
    errors: errors
  };
}
function quickCompactDispatchAlreadyWaitingUnderLock_(ss, deadlineMs) {
  // V13.19.4:
  // 1) Clear đúng row đã chắc chắn append sang Xe chờ giao (1-1 visitId + cùng BSX).
  // 2) Sau khi clear xong mới SAFE ROW COMPACT bằng moveRows().
  //    Không rewrite A:X, không setValues toàn bảng, không rebuild Data Validation.
  if (isWebAppConfirmPriorityActive_()) {
    return {
      compacted: 0,
      plates: [],
      yieldedForWebApp: true,
      identityMismatches: [],
      packedRows: 0,
      packedBlocks: 0
    };
  }
  if (deadlineMs && Date.now() >= deadlineMs) {
    return {
      compacted: 0,
      plates: [],
      deadlineReached: true,
      identityMismatches: [],
      packedRows: 0,
      packedBlocks: 0
    };
  }

  const sh = ss.getSheetByName(CFG.SHEET_DISPATCH);
  const wd = ss.getSheetByName(CFG.SHEET_WAITING_DELIVERY);
  if (!sh || !wd) {
    return {
      compacted: 0,
      plates: [],
      identityMismatches: [],
      packedRows: 0,
      packedBlocks: 0
    };
  }

  const start = CFG.FIRST_DATA_ROW;
  const last = sh.getLastRow();
  const candidates = [];
  const identityMismatches = [];

  if (last >= start) {
    const wdStart = CFG.WAITING_DELIVERY_FIRST_ROW;
    const wdLast = waitingDeliveryDataLastRow_(wd);

    if (wdLast >= wdStart) {
      const waiting = wd
        .getRange(wdStart, 1, wdLast - wdStart + 1, CFG.WD.VISIT_ID)
        .getDisplayValues();

      const waitingCount = Object.create(null);
      const waitingPlate = Object.create(null);
      waiting.forEach(function(r) {
        const id = clean_(r[CFG.WD.VISIT_ID - 1]);
        if (!id) return;
        waitingCount[id] = (waitingCount[id] || 0) + 1;
        if (!waitingPlate[id]) {
          waitingPlate[id] = normalizePlate_(r[CFG.WD.BIEN_SO - 1]);
        }
      });

      const width = CFG.DP.LAST_SCC_STATUS;
      const display = sh
        .getRange(start, 1, last - start + 1, width)
        .getDisplayValues();

      const dispatchCount = Object.create(null);
      display.forEach(function(r) {
        const id = clean_(r[CFG.DP.VISIT_ID - 1]);
        if (!id) return;
        dispatchCount[id] = (dispatchCount[id] || 0) + 1;
      });

      display.forEach(function(r, idx) {
        const id = clean_(r[CFG.DP.VISIT_ID - 1]);
        const plate = normalizePlate_(r[CFG.DP.BIEN_SO - 1]);
        if (!id || !plate) return;
        if (dispatchCount[id] !== 1 || waitingCount[id] !== 1) return;

        const wdPlate = normalizePlate_(waitingPlate[id]);
        if (!wdPlate || wdPlate !== plate) {
          identityMismatches.push({
            row: start + idx,
            visitId: id,
            dispatchPlate: plate,
            waitingPlate: wdPlate || ''
          });
          return;
        }

        candidates.push({
          row: start + idx,
          visitId: id,
          plate: plate
        });
      });
    }
  }

  if (identityMismatches.length) {
    console.log(
      '[WARN][DispatchIdentityMismatch] ' +
      JSON.stringify(identityMismatches)
    );
  }

  let cleared = 0;
  const plates = [];

  // Chưa move row ở giai đoạn này để row number của toàn bộ candidates ổn định.
  for (let i = 0; i < candidates.length; i++) {
    if (isWebAppConfirmPriorityActive_()) {
      return {
        compacted: cleared,
        plates: plates,
        yieldedForWebApp: true,
        identityMismatches: identityMismatches,
        packedRows: 0,
        packedBlocks: 0
      };
    }
    if (deadlineMs && Date.now() >= deadlineMs) {
      return {
        compacted: cleared,
        plates: plates,
        deadlineReached: true,
        identityMismatches: identityMismatches,
        packedRows: 0,
        packedBlocks: 0
      };
    }

    const c = candidates[i];
    const currentId = clean_(sh.getRange(c.row, CFG.DP.VISIT_ID).getValue());
    const currentPlate = normalizePlate_(readPlateText_(sh, c.row));
    if (currentId !== c.visitId || currentPlate !== c.plate) {
      console.log(
        '[WARN][AutoClearSkipOwnershipChanged] ' +
        JSON.stringify({row: c.row, visitId: c.visitId, plate: c.plate})
      );
      continue;
    }

    const owners = resolveDispatchRowsByVisitId_(sh, c.visitId);
    if (owners.length !== 1 || owners[0] !== c.row) {
      console.log(
        '[WARN][AutoClearSkipNonUnique] ' +
        JSON.stringify({row: c.row, visitId: c.visitId, owners: owners})
      );
      continue;
    }

    // P0.2 GATE CUỐI: Waiting đã tồn tại KHÔNG đủ để xóa Điều phối.
    // Xe có thể đã được mở lại công việc sau khi append hoặc Waiting
    // được tạo bởi bản code cũ; phải xác nhận trạng thái LIVE trước clear.
    try {
      const waitingOwnerRows = waitingDeliveryRowsByVisitId_(wd, c.visitId);
      if (waitingOwnerRows.length !== 1 ||
          normalizeKey_(wd.getRange(waitingOwnerRows[0], CFG.WD.TRANG_THAI)
            .getDisplayValue()) !== 'cho giao') {
        console.log('[P02][CompactBlockedWaitingOwnership] ' + c.visitId);
        continue;
      }
      const liveRec = readDispatchRow_(sh, c.row);
      if (clean_(liveRec.visitId) !== c.visitId ||
          normalizePlate_(liveRec.plate) !== c.plate) {
        console.log('[P02][CompactBlockedLiveIdentity] ' + c.visitId);
        continue;
      }
      const liveRoute = classifyWork_(liveRec);
      if (!isReadyForWaitingDelivery_(liveRec, liveRoute)) {
        console.log('[P02][CompactBlockedUnfinished] ' + c.plate);
        continue;
      }
      const incomplete = validateStoredDataBeforeClear_(liveRec, liveRoute);
      if (incomplete.length) {
        console.log('[P02][CompactBlockedStoredData] ' + c.plate + ' | ' + incomplete.join(' | '));
        continue;
      }
      assertVisitWorkCompletedBeforeWaiting_(liveRec, liveRoute);
    } catch (gateErr) {
      console.log('[P02][CompactBlockedGateError] ' + c.plate + ' | ' + gateErr.message);
      continue;
    }
    clearDispatchOwnedRowOnly_(sh, c.row, c.visitId);
    cleared++;
    plates.push(c.plate);
  }

  // V13.19.4: kể cả không có candidate mới, vẫn compact các gap còn lại từ
  // V13.19.3 / execution trước. moveRows() di chuyển nguyên row vật lý nên
  // chip màu + validation + note + hidden columns đi cùng record.
  let pack = {
    packedRows: 0,
    packedBlocks: 0,
    yieldedForWebApp: false,
    blocked: false
  };
  if (
    !isWebAppConfirmPriorityActive_() &&
    (!deadlineMs || Date.now() < deadlineMs)
  ) {
    pack = safeCompactDispatchRowsUnderLock_(sh, deadlineMs);
  }

  return {
    compacted: cleared,
    plates: plates,
    identityMismatches: identityMismatches,
    packedRows: Number(pack.packedRows || 0),
    packedBlocks: Number(pack.packedBlocks || 0),
    packBlocked: !!pack.blocked,
    packReason: clean_(pack.reason || ''),
    yieldedForWebApp: !!pack.yieldedForWebApp
  };
}

function dongBoLaiDongBiKet() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return {
      success: false,
      busy: true,
      message: 'Hệ thống đang bận. Hãy chạy lại sau ít phút.'
    };
  }
  try {
    const ss = laySpreadsheetWeb_();
    assertVisitSetup_(ss);

    // Phase 1: dọn local cực nhanh các lượt đã được append sang Xe chờ giao
    // nhưng clear row Điều phối bị gián đoạn/timeout.
    const quick = quickCompactDispatchAlreadyWaitingUnderLock_(ss);

    // Phase 2: recovery có giới hạn thời gian, tuyệt đối không scan 9999 dòng.
    const deadline =
      Date.now() + Number(CONCURRENCY_CFG.MANUAL_RECOVERY_TIME_BUDGET_MS || 25000);
    const result = reconcileDispatchStuckUnderLock_(
      ss,
      Number(CONCURRENCY_CFG.MANUAL_RECOVERY_BATCH || 4),
      deadline
    );

    const total = Number(quick.compacted || 0) + Number(result.processed || 0);
    try {
      ss.toast(
        total > 0
          ? 'Đã dọn ' + total + ' xe (nhanh: ' + quick.compacted +
            ', recovery: ' + result.processed + '). Nếu còn xe, chạy lại menu thêm 1 lần.'
          : 'Chưa có thêm xe đủ điều kiện để dọn.',
        '⚡ DỌN XE TIMEOUT-SAFE',
        10
      );
      if (result.errors.length) {
        ss.toast(
          result.errors.slice(0, 6).join('\n'),
          '⚠️ Xe cần kiểm tra',
          10
        );
      }
    } catch (err) {}

    return {
      success: result.errors.length === 0,
      busy: false,
      quickCompacted: quick.compacted,
      recovered: result.processed,
      processed: total,
      errors: result.errors
    };
  }
  finally {
    try {
      SpreadsheetApp.flush();
    }
    finally {
      lock.releaseLock();
    }
  }
}

/*********************************************************
 * CÀI / GỠ WORKER
 *********************************************************/
function caiTriggerPendingVisitWorker() {
  goTriggerPendingVisitWorker();
  ScriptApp
    .newTrigger(
      'runPendingVisitWorker'
    )
    .timeBased()
    .everyMinutes(
      1
    )
    .create();
  return {
    success: true,
    message:
      'Đã cài worker pending visit mỗi 1 phút.'
  };
}
function goTriggerPendingVisitWorker() {
  let removed = 0;
  ScriptApp
    .getProjectTriggers()
    .forEach(
      function(trigger) {
        if (
          trigger.getHandlerFunction() ===
          'runPendingVisitWorker'
        ) {
          ScriptApp.deleteTrigger(
            trigger
          );
          removed++;
        }
      }
    );
  return {
    success: true,
    removed: removed
  };
}
/*********************************************************
 * TIỆN ÍCH QUEUE - V13.19
 *********************************************************/
function xemTrangThaiPendingQueue() {
  const props =
    PropertiesService
      .getScriptProperties()
      .getProperties();

  let pending = 0;
  let quarantine = 0;

  Object.keys(props).forEach(function(key) {
    if (
      key.indexOf(CONCURRENCY_CFG.DIRTY_PREFIX) === 0 ||
      key.indexOf(CONCURRENCY_CFG.DIRTY_FALLBACK_PREFIX) === 0
    ) {
      pending++;
    }
    else if (
      key.indexOf(CONCURRENCY_CFG.QUARANTINE_PREFIX) === 0
    ) {
      quarantine++;
    }
  });

  const result = {
    success: true,
    pending: pending,
    quarantine: quarantine
  };

  try {
    SpreadsheetApp
      .getActive()
      .toast(
        'Pending: ' + pending + ' | Quarantine: ' + quarantine,
        '📬 QUEUE HẬU XỬ LÝ',
        6
      );
  }
  catch (err) {}

  console.log(
    '[QUEUE] ' +
    JSON.stringify(result)
  );

  return result;
}

/**
 * P0.5.7.1 - Worker manual execution diagnostics only.
 * Unchanged one-call worker; no new triggers, timestamps, statuses, or data edits.
 */
function chayWorkerPendingNgay() {
  const startedAt = Date.now();
  try {
    const result = runPendingVisitWorker();
    console.log('[P057_WORKER_RESULT] ' + JSON.stringify(result));
    console.log('[P057_WORKER_ELAPSED_MS] ' + (Date.now() - startedAt));
    return result;
  } catch (err) {
    console.error('[P057_WORKER_ERROR] ' + String(err && (err.stack || err.message) || err));
    throw err;
  }
}


/*********************************************************
 * XỬ LÝ DÒNG VỪA ĐƯỢC WEB APP GHI
 *
 * WebApp.gs đang giữ ScriptLock nên hàm này KHÔNG lấy lock lại.
 * Không tạo event giả trên toàn A:S; chỉ xử lý dữ liệu hiện có.
 *********************************************************/
function enqueueWebAppFastPostProcess_(
  sh,
  row,
  meta
) {
  if (
    !sh ||
    sh.getName() !== CFG.SHEET_DISPATCH ||
    row < CFG.FIRST_DATA_ROW
  ) {
    throw new Error(
      'FAST COMMIT: dòng Điều phối không hợp lệ.'
    );
  }

  meta = meta || {};

  const visitId =
    clean_(
      sh
        .getRange(
          row,
          CFG.DP.VISIT_ID
        )
        .getValue()
    );

  const plate =
    readPlateText_(
      sh,
      row
    );

  if (!visitId) {
    throw new Error(
      'FAST COMMIT: thiếu visitId sau khi ghi Điều phối.'
    );
  }

  if (!plate) {
    throw new Error(
      'FAST COMMIT: thiếu Biển số sau khi ghi Điều phối.'
    );
  }

  const width =
    Math.max(
      CFG.DP.MODEL,
      CFG.DP.NGAY,
      CFG.DP.CVDV
    );

  const rowValues =
    sh
      .getRange(
        row,
        1,
        1,
        width
      )
      .getValues()[0];

  const rowDisplay =
    sh
      .getRange(
        row,
        1,
        1,
        width
      )
      .getDisplayValues()[0];

  const nowIso =
    new Date().toISOString();

  const job = {
    sheet: CFG.SHEET_DISPATCH,
    visitId: visitId,
    rowHint: row,
    plate: plate,

    // A:S được xem như lượt insert hoàn chỉnh.
    // Worker luôn đọc state hiện tại từ sheet.
    colStart: 1,
    numCols: CFG.DP.GHI_CHU,
    oldValue: '',

    reason: 'webapp-fast-commit',
    source: 'WEBAPP_FAST_COMMIT',
    requestId: clean_(meta.requestId || ''),

    dateKey:
      cvdvDateKey_(
        rowValues[CFG.DP.NGAY - 1],
        sh.getParent()
      ),

    advisor:
      normalizeKey_(
        rowDisplay[CFG.DP.CVDV - 1]
      ),

    model:
      normalizeKey_(
        rowDisplay[CFG.DP.MODEL - 1]
      ),

    isNewWebAppInsert: true,

    updatedAt: nowIso,
    createdAt: nowIso,
    attempts: 0,
    revision: Utilities.getUuid()
  };

  const key =
    dirtyVisitPropertyKey_(
      CFG.SHEET_DISPATCH,
      visitId
    );

  const props =
    PropertiesService
      .getScriptProperties();

  // WebApp gọi khi đang giữ ScriptLock. Worker cũng cần ScriptLock,
  // nên không thể consume job trước khi WebApp release lock.
  // visitId là ownership key duy nhất => setProperty trực tiếp là an toàn
  // và tránh tranh thêm queue lock trong fast path.
  const scriptLock =
    LockService.getScriptLock();

  if (
    scriptLock &&
    typeof scriptLock.hasLock === 'function' &&
    scriptLock.hasLock()
  ) {
    const currentRaw =
      props.getProperty(
        key
      );

    if (currentRaw) {
      try {
        const current =
          JSON.parse(
            currentRaw
          );
        job.attempts =
          Number(
            current.attempts || 0
          );
        job.createdAt =
          current.createdAt ||
          job.createdAt;
        p043MergeCompletionEvidence_(job, current);
      }
      catch (err) {}
    }
    p043MergeCompletionEvidence_(job, null);

    job.revision =
      Utilities.getUuid();

    props.setProperty(
      key,
      JSON.stringify(
        job
      )
    );

    return true;
  }

  return enqueueDirtyVisit_(
    job
  );
}


/*********************************************************
 * COMPATIBILITY
 * V13.19: WebApp cũ gọi hàm này cũng chỉ enqueue hậu xử lý.
 *********************************************************/



/*********************************************************
 * XỬ LÝ MỘT DÒNG ĐIỀU PHỐI
 *********************************************************/
function xuLyMotDongDieuPhoi_(
  sh,
  row,
  e
) {
  const c =
    CFG.DP;
  const isNewWebAppInsert =
    !!(e && e.isNewWebAppInsert);
  const isRetryWorker =
    !!(e && e.isRetryWorker);
  const retryableAsyncErrors = [];
  const editStartCol =
    e.range.getColumn();
  const editEndCol =
    e.range.getLastColumn();
  const editedPlate =
    editStartCol <=
      c.BIEN_SO &&
    editEndCol >=
      c.BIEN_SO;
  const editedEMStatus =
    editStartCol <=
      c.TD_EM &&
    editEndCol >=
      c.TD_EM;
  const editedSCCStatus =
    editStartCol <=
      c.TD_SCC &&
    editEndCol >=
      c.TD_SCC;
  // V13.14:
  // Hẹn giao CVDV cần sync khi Cố vấn / Ngày giao / Giờ giao thay đổi.
  const editedDeliveryPromise =
    (
      editStartCol <= c.CVDV &&
      editEndCol >= c.CVDV
    ) ||
    (
      editStartCol <= c.NGAY_GIAO &&
      editEndCol >= c.NGAY_GIAO
    ) ||
    (
      editStartCol <= c.GIO_GIAO &&
      editEndCol >= c.GIO_GIAO
    );
  // V13.17: A/B/C là identity để kiểm tra chéo Cố vấn.
  const editedAdvisorIdentity =
    (
      editStartCol <= c.NGAY &&
      editEndCol >= c.NGAY
    ) ||
    (
      editStartCol <= c.BIEN_SO &&
      editEndCol >= c.BIEN_SO
    ) ||
    (
      editStartCol <= c.CVDV &&
      editEndCol >= c.CVDV
    );
  /*******************************************************
   * BIỂN SỐ
   *******************************************************/
  if (editedPlate) {
    normalizePlateCellFromEvent_(
      sh,
      row,
      e
    );
  }
  /*******************************************************
   * GIỜ NHẬN / GIỜ GIAO
   *******************************************************/
  normalizeEditedTimeCells_(
    sh,
    row,
    editStartCol,
    editEndCol,
    e
  );
  const plate =
    readPlateText_(
      sh,
      row
    );
  if (!plate) {
    // P0.5.8: empty/temporarily erased plate is NOT evidence of a closed visit.
    // Only the verified Dispatch -> Waiting transition is allowed to clear V.
    const existingId=clean_(sh.getRange(row,CFG.DP.VISIT_ID).getDisplayValue());
    const tail=sh.getRange(row,CFG.DP.LAST_EM_STATUS,1,2).getDisplayValues()[0];
    if(existingId || tail.some(function(v){return !!clean_(v);})){ 
      throw new Error('P058_PLATE_MISSING_KEEP_VISIT_ID: dòng '+row+
        ', visitId='+existingId+'. Không xóa V:X khi biển số bị trống.');
    }
    return;
  }
  // P0.4: chặn sớm trước sync/ghi Data khi xe đã có một lượt Chờ giao
  // chưa đóng hoặc có một dòng Điều phối khác cùng biển số.
  // Không xóa dòng và không đụng lịch sử; người quản trị phải xác minh lượt.
  p04AssertPhysicalVisitOwnership_(sh, row, plate);
  if (editedPlate && (!clean_(sh.getRange(row, c.CVDV).getDisplayValue()) ||
      !clean_(sh.getRange(row, c.MODEL).getDisplayValue()))) {
    if (!e._modelHistoryCache) e._modelHistoryCache = {records: null};
    autoFillAdvisorModelFromHistory_(sh, row, e._modelHistoryCache);
  }
  /*******************************************************
   * NGÀY NHẬN
   *******************************************************/
  const dateCell =
    sh.getRange(
      row,
      c.NGAY
    );
  if (!dateCell.getValue()) {
    dateCell
      .setValue(
        todayOnly_()
      )
      .setNumberFormat(
        'dd/MM/yyyy'
      );
  }
  /*******************************************************
   * NHẬP BIỂN SỐ
   * -> AUTO GIỜ NHẬN
   * -> AUTO NGÀY GIAO
   *******************************************************/
  if (editedPlate) {
    const receiveCell =
      sh.getRange(
        row,
        c.GIO_NHAN
      );
    if (
      receiveCell.getValue() === ''
    ) {
      const now =
        new Date();
      setTimeCell_(
        receiveCell,
        now.getHours(),
        now.getMinutes()
      );
    }
    const deliveryDateCell =
      sh.getRange(
        row,
        c.NGAY_GIAO
      );
    if (
      deliveryDateCell.getValue() === ''
    ) {
      const workDate =
        dateCell.getValue();
      deliveryDateCell
        .setValue(
          workDate instanceof Date
            ? stripTime_(workDate)
            : todayOnly_()
        )
        .setNumberFormat(
          'dd/MM/yyyy'
        );
    }
  }
  formatDispatchTimeCells_(
    sh,
    row
  );
  /*******************************************************
   * ĐỌC DỮ LIỆU
   *******************************************************/
  let rec =
    readDispatchRow_(
      sh,
      row
    );
  let route =
    classifyWork_(
      rec
    );
  // P0.5.3.1: LSC staged có thể đặt Q="Chờ giao việc" trước khi
  // người điều phối chọn E/F. Chỉ dọn Q mặc định nếu chứng minh
  // CHƯA có Data SCC, CHƯA có mốc SCC, không có KTV G/H riêng.
  // KHÔNG xóa trạng thái Đang làm/Hoàn thành, lịch sử hoặc giờ.
  if (p0531MrKhoiEmOnly_(rec) && !route.scc.active &&
      normalizeKey_(rec.sccStatus) === 'cho giao viec' &&
      !rec.sccConfirmationStamp && !isComplete_(rec.lastSCCStatus)) {
    const p0531Scc = sh.getParent().getSheetByName(CFG.SHEET_DATA_SCC);
    if (p0531Scc && !findDataRow_(p0531Scc, rec)) {
      sh.getRange(row, c.TD_SCC).clearContent();
      if (normalizeKey_(rec.lastSCCStatus) === 'cho giao viec') {
        sh.getRange(row, c.LAST_SCC_STATUS).clearContent();
      }
      rec = readDispatchRow_(sh, row);
      route = classifyWork_(rec);
      console.log('[P0531_MR_KHOI_EM_ONLY] ' + rec.plate +
        ' / ' + rec.visitId + ' / default SCC pending cleared');
    }
  }
  /*******************************************************
   * V13.20 - OFFLOAD CVDV BRIDGE
   *
   * MAIN KHÔNG còn tự động:
   * - mở file Lễ tân/CVDV để cross-check Cố vấn;
   * - ghi Hẹn giao N/O sang CVDV G.
   *
   * Hai trách nhiệm này thuộc worker riêng trong file CVDV V2.7.0.
   * Mục tiêu:
   * - WebApp/worker MAIN không chờ cross-workbook I/O;
   * - giảm tranh ScriptLock và timeout;
   * - mỗi file chịu trách nhiệm nghiệp vụ của mình.
   *
   * Các hàm manual fallback vẫn được giữ phía dưới code và chỉ chạy
   * khi người dùng chủ động chọn menu dự phòng.
   *******************************************************/
  /*******************************************************
   * VALIDATE EM TRƯỚC HOÀN THÀNH
   *******************************************************/
  if (
    isComplete_(
      rec.emStatus
    )
  ) {
    const missing =
      validateEMCompletion_(
        rec,
        route
      );
    if (
      missing.length > 0
    ) {
      restoreStatusAfterValidationFail_(
        sh,
        row,
        c.TD_EM,
        e
      );
      showCompletionWarning_(
        rec.plate,
        'EM',
        missing,
        sh.getParent()
      );
      return;
    }
  }
  /*******************************************************
   * VALIDATE SCC TRƯỚC HOÀN THÀNH
   *******************************************************/
  if (
    isComplete_(
      rec.sccStatus
    )
  ) {
    const missing =
      validateSCCCompletion_(
        rec,
        route
      );
    if (
      missing.length > 0
    ) {
      restoreStatusAfterValidationFail_(
        sh,
        row,
        c.TD_SCC,
        e
      );
      showCompletionWarning_(
        rec.plate,
        'SCC',
        missing,
        sh.getParent()
      );
      return;
    }
  }
  /*******************************************************
   * TIMESTAMP TRẠNG THÁI
   *******************************************************/
  // P0.4.2: worker retry KHÔNG được tạo timestamp mới thay cho lần bấm.
  p042RestoreQueuedConfirmation_(sh, row, c.TD_EM, c.MOC_TD_EM, e);
  p042RestoreQueuedConfirmation_(sh, row, c.TD_SCC, c.MOC_TD_SCC, e);
  updateStatusStamp_(
    sh,
    row,
    c.TD_EM,
    c.MOC_TD_EM,
    p043StatusReallyChanged_(sh, row, c.TD_EM, c.LAST_EM_STATUS, editedEMStatus, isRetryWorker),
    !isRetryWorker && editedEMStatus && e.p044CapturedEventAt ? e.p044CapturedEventAt : null
  );
  updateStatusStamp_(
    sh,
    row,
    c.TD_SCC,
    c.MOC_TD_SCC,
    p043StatusReallyChanged_(sh, row, c.TD_SCC, c.LAST_SCC_STATUS, editedSCCStatus, isRetryWorker),
    !isRetryWorker && editedSCCStatus && e.p044CapturedEventAt ? e.p044CapturedEventAt : null
  );
  /*******************************************************
   * ĐỌC LẠI
   *******************************************************/
  rec =
    readDispatchRow_(
      sh,
      row
    );
  route =
    classifyWork_(
      rec
    );
  // PREPARE -> COMPLETE: phục hồi OPEN trước, không sinh Data đã Hoàn thành
  // trực tiếp; chỉ dùng mốc xác nhận hợp lệ của sự kiện sửa trạng thái.
  p043GuardPendingCompletedStatus_(rec, route, e);
  p042EnsureWorkRowsBeforeCompletion_(rec, route, e);
  if (!isNewWebAppInsert) syncVisitIdentity_(rec);
  /*******************************************************
   * DATA EM
   *******************************************************/
  if (
    route.em.active
  ) {
    syncDataEM_(
      rec,
      route.em.workText
    );
  } else if (!isNewWebAppInsert) {
    removeOpenDataRowIfStale_(
      CFG.SHEET_DATA_EM,
      rec
    );
  }
  /*******************************************************
   * DATA SCC
   *******************************************************/
  if (
    route.scc.active
  ) {
    syncDataSCC_(
      rec,
      route.scc.workText
    );
  } else if (!isNewWebAppInsert) {
    removeOpenDataRowIfStale_(
      CFG.SHEET_DATA_SCC,
      rec
    );
  }
  /*******************************************************
   * TRACKING KT
   *******************************************************/
  syncTrackingSheets_(
    rec,
    route,
    e
  );
  sh.getRange(row, CFG.DP.LAST_EM_STATUS, 1, 2).setValues([[rec.emStatus, rec.sccStatus]]);

  // External sync lỗi tạm thời: giữ job lại để retry, không cho worker
  // xóa queue như thể đã xử lý hoàn tất. Các sync nội bộ đã idempotent.
  if (
    isRetryWorker &&
    retryableAsyncErrors.length > 0
  ) {
    const retryErr =
      new Error(
        retryableAsyncErrors.join(' | ')
      );
    retryErr.code =
      'ASYNC_DOWNSTREAM_RETRY';
    throw retryErr;
  }

  /*******************************************************
   * DỌN XE
   *******************************************************/
  if (
    isReadyForWaitingDelivery_(
      rec,
      route
    )
  ) {
    SpreadsheetApp.flush();
    const errors =
      validateStoredDataBeforeClear_(
        rec,
        route
      );
    if (
      errors.length > 0
    ) {
      sh.getParent()
        .toast(
          'Xe ' +
          rec.plate +
          ' chưa thể xóa khỏi Điều phối.\n' +
          errors.join(' | '),
          '⚠️ DATA CHƯA ĐỦ',
          8
        );
      return;
    }
    moveDispatchToWaitingDelivery_(
      sh,
      row,
      rec
    );
  }
}
/*********************************************************
 * ONE-TIME REPAIR v13.9
 * KHÔI PHỤC CHIP KIỂM TRA TỪ NOTE CỘT I
 *
 * Chỉ bổ sung "Kiểm tra" / "Kiểm tra lại" vào CV nếu:
 * - dòng đang có xe,
 * - Note hiện hữu đã xác định inspectionType,
 * - CV chưa có loại kiểm tra tương ứng.
 *
 * Không đụng visitId / trạng thái / Data EM / Data SCC.
 *********************************************************/
/**
 * P0.5.3.1 – Mr. Khởi là tổ trưởng EM, không phải một tổ SCC.
 * Nếu không có KTV SCC được phân công RIÊNG tại G/H, công việc của
 * ông Khởi đi theo một luồng EM. Không tự hoàn thành hay tạo giờ công.
 * Nhận diện theo F; không phụ thuộc thứ tự nhập E (nhóm EM) và F.
 */
function p0531MrKhoiEmOnly_(rec) {
  return !!rec && /^mr\.?\s*khoi$/.test(normalizeKey_(rec.inspector)) &&
    !clean_(rec.ktv1) && !clean_(rec.ktv2);
}

function classifyWork_(rec) {
  const jobs =
    splitMulti_(
      rec.workRaw
    );
  const em = [];
  const scc = [];
  /*******************************************************
   * PHÂN CÔNG THỰC TẾ
   *
   * Chỉ dùng để quyết định bên thực hiện đối với
   * CÔNG VIỆC LINH HOẠT.
   *
   * Không được dùng assignment để tự tạo thêm luồng.
   *******************************************************/
  const hasEmAssignment =
    !!clean_(
      rec.emGroup
    );
  const hasSccPersonAssignment =
    !!(clean_(rec.inspector) || clean_(rec.ktv1) || clean_(rec.ktv2));
  const hasSccAssignment =
    !!(
      hasSccPersonAssignment ||
      valueExists_(
        rec.itemCount
      )
    );
  /*******************************************************
   * LOẠI KIỂM TRA TỪ NOTE WEB APP
   *
   * "Kiểm tra / Kiểm tra lại" không cần nằm trong cột CV.
   *******************************************************/
  const inspectionType =
    clean_(
      rec.inspectionType
    );
  if (
    normalizeKey_(
      inspectionType
    ) === 'kiem tra lai'
  ) {
    pushUnique_(
      scc,
      'Kiểm tra lại'
    );
  }
  else if (
    normalizeKey_(
      inspectionType
    ) === 'kiem tra'
  ) {
    if (
      hasEmAssignment
    ) {
      pushUnique_(
        em,
        'Kiểm tra'
      );
    }
    if (
      hasSccPersonAssignment
    ) {
      pushUnique_(
        scc,
        'Kiểm tra'
      );
    }
  }
  /*******************************************************
   * CÔNG VIỆC LINH HOẠT
   *
   * BD / EM60
   * BD++
   * Bảo hành
   * ĐK
   * Kiểm tra
   *
   * Quy tắc:
   * - Có Nhóm EM       -> EM
   * - Có nhân sự SCC   -> SCC
   * - Có cả hai        -> cả EM + SCC
   * - Chưa phân công   -> chưa kích hoạt luồng
   *******************************************************/
  function flexible(
    emName,
    sccName
  ) {
    if (
      hasEmAssignment
    ) {
      pushUnique_(
        em,
        emName
      );
    }
    if (
      hasSccAssignment
    ) {
      pushUnique_(
        scc,
        sccName
      );
    }
  }
  /*******************************************************
   * DUYỆT DANH SÁCH CÔNG VIỆC
   *******************************************************/
  jobs.forEach(job => {
    const key =
      normalizeKey_(
        job
      );
    /*****************************************************
     * 1K
     *
     * CỐ ĐỊNH EM
     *****************************************************/
    if (
      key === '1k'
    ) {
      pushUnique_(
        em,
        '1K'
      );
      return;
    }
    /*****************************************************
     * HỖ TRỢ SCC
     *
     * CỐ ĐỊNH EM
     *****************************************************/
    if (
      key === 'ho tro scc'
    ) {
      pushUnique_(
        em,
        'Hỗ trợ SCC'
      );
      return;
    }
    /*****************************************************
     * BD / EM60
     *
     * LINH HOẠT EM / SCC
     *****************************************************/
    if (
      key === 'bd' ||
      key === 'em60'
    ) {
      flexible(
        'EM60',
        'BD'
      );
      return;
    }
    /*****************************************************
     * BD++
     *
     * LINH HOẠT EM / SCC
     *****************************************************/
    if (
      key === 'bd++' ||
      key === 'bdplusplus' ||
      key === 'bd + +'
    ) {
      flexible(
        'BD++',
        'BD++'
      );
      return;
    }
    /*****************************************************
     * BẢO HÀNH
     *
     * LINH HOẠT EM / SCC
     *****************************************************/
    if (
      key === 'bao hanh'
    ) {
      flexible(
        'Bảo hành',
        'Bảo hành'
      );
      return;
    }
    /*****************************************************
     * ĐĂNG KIỂM
     *
     * LINH HOẠT EM / SCC
     *****************************************************/
    if (
      key === 'dk' ||
      key === 'dang kiem'
    ) {
      flexible(
        'ĐK',
        'ĐK'
      );
      return;
    }
    /*****************************************************
     * KIỂM TRA
     *
     * LINH HOẠT EM / SCC
     *****************************************************/
    if (
      key === 'kiem tra'
    ) {
      if (hasEmAssignment) pushUnique_(em, 'Kiểm tra');
      if (hasSccPersonAssignment) pushUnique_(scc, 'Kiểm tra');
      return;
    }
    /*****************************************************
     * SC CƠ BẢN
     *
     * CỐ ĐỊNH SCC
     *****************************************************/
    if (
      key === 'sc co ban' ||
      key === 'co ban'
    ) {
      pushUnique_(
        scc,
        'Cơ bản'
      );
      return;
    }
    /*****************************************************
     * ĐẠI TU
     *
     * CỐ ĐỊNH SCC
     *****************************************************/
    if (
      key === 'dai tu'
    ) {
      pushUnique_(
        scc,
        'Đại tu'
      );
      return;
    }
    /*****************************************************
     * DÀN LẠNH
     *
     * CỐ ĐỊNH SCC
     *****************************************************/
    if (
      key === 'dan lanh'
    ) {
      pushUnique_(
        scc,
        'Dàn lạnh'
      );
      return;
    }
    /*****************************************************
     * KIỂM TRA LẠI
     *
     * CỐ ĐỊNH SCC
     *****************************************************/
    if (
      key === 'kiem tra lai'
    ) {
      pushUnique_(
        scc,
        'Kiểm tra lại'
      );
      return;
    }
    /*****************************************************
     * KIỂM TU
     *
     * CỐ ĐỊNH SCC
     *****************************************************/
    if (
      key === 'kiem tu'
    ) {
      pushUnique_(
        scc,
        'Kiểm tu'
      );
      return;
    }
    /*****************************************************
     * VSDL
     *
     * CỐ ĐỊNH SCC
     *****************************************************/
    if (
      key === 'vsdl'
    ) {
      pushUnique_(
        scc,
        'VSDL'
      );
      return;
    }
    /*****************************************************
     * 3D
     *
     * CỐ ĐỊNH SCC
     *****************************************************/
    if (
      key === '3d'
    ) {
      pushUnique_(
        scc,
        '3D'
      );
      return;
    }
  });
  // P0.5.3.1: một người thực hiện là Mr. Khởi (tổ trưởng EM),
  // không có KTV SCC độc lập => chuyển các hạng mục SCC của LSC
  // về luồng EM để chỉ yêu cầu EM hoàn thành. Giữ nguyên tên hạng mục.
  // Khi G/H có tổ SCC thực tế, quy tắc phân luồng gốc vẫn giữ nguyên.
  if (p0531MrKhoiEmOnly_(rec)) {
    scc.forEach(function(work) { pushUnique_(em, work); });
    scc.length = 0;
  }
  /*******************************************************
   * QUAN TRỌNG
   *
   * Luồng chỉ ACTIVE khi thật sự có công việc thuộc luồng đó.
   *
   * KHÔNG dùng:
   *   em.length > 0 || hasEmAssignment
   *
   * vì sẽ làm:
   *   3D + EM1 -> sinh thêm EM
   *   1K + KTV -> sinh thêm SCC
   *******************************************************/
  const emActive =
    em.length > 0;
  const sccActive =
    scc.length > 0;
  return {
    em: {
      active:
        emActive,
      workText:
        em.join(', '),
      items:
        em
    },
    scc: {
      active:
        sccActive,
      workText:
        scc.join(', '),
      items:
        scc
    }
  };
}
/*********************************************************
 * VALIDATION EM
 *********************************************************/
function validateEMCompletion_(
  rec,
  route
) {
  if (
    !route.em.active
  ) {
    const assignmentMissing = [];
    addInspectionAssignmentMissing_(rec, assignmentMissing);
    return uniqueArray_(assignmentMissing);
  }
  const missing = [];
  addMissingCommon_(
    rec,
    missing
  );
  addInspectionAssignmentMissing_(rec, missing);
  // P0.5.3.1: Mr. Khoi trong cot F, khong co KTV SCC rieng G/H,
  // la nguoi thuc hien EM hop le ngay ca khi cot E (Nhom EM) de trong.
  // Ngoai le chi ap dung cho dung Mr. Khoi-only; cac luot EM khac van
  // bat buoc khai bao Nhom EM nhu cu.
  if (
    !rec.emGroup && !p0531MrKhoiEmOnly_(rec)
  ) {
    missing.push(
      'Nhóm EM'
    );
  }
  if (
    !route.em.workText
  ) {
    missing.push(
      'Công việc EM'
    );
  }
  const needsBDLevel =
    route.em.items
      .indexOf(
        'EM60'
      ) !== -1 ||
    route.em.items
      .indexOf(
        'BD++'
      ) !== -1;
  if (
    needsBDLevel &&
    !rec.bdLevel
  ) {
    missing.push(
      'Cấp BD'
    );
  }
  return uniqueArray_(
    missing
  );
}
/*********************************************************
 * VALIDATION SCC
 *********************************************************/
function validateSCCCompletion_(
  rec,
  route
) {
  if (
    !route.scc.active
  ) {
    const assignmentMissing = [];
    addInspectionAssignmentMissing_(rec, assignmentMissing);
    return uniqueArray_(assignmentMissing);
  }
  const missing = [];
  addMissingCommon_(
    rec,
    missing
  );
  addInspectionAssignmentMissing_(rec, missing);
  if (
    !route.scc.workText
  ) {
    missing.push(
      'Công việc SCC'
    );
  }
  const inspectionJob =
    route.scc.items
      .indexOf(
        'Kiểm tra'
      ) !== -1 ||
    route.scc.items
      .indexOf(
        'Kiểm tra lại'
      ) !== -1;
  if (!inspectionJob && !rec.ktv1) {
    missing.push('Nhóm SCC');
  }
  return uniqueArray_(
    missing
  );
}function addMissingCommon_(rec, missing) {
  if (!rec.plate) missing.push('Biển số');
  if (!rec.advisor) missing.push('CVDV');
  if (!rec.model) missing.push('Model xe');
  const receiveDateOK =
    rec.date instanceof Date &&
    !isNaN(rec.date.getTime());
  const deliveryDateOK =
    rec.deliveryDate instanceof Date &&
    !isNaN(rec.deliveryDate.getTime());
  const receiveTimeOK =
    isValidTimeValue_(rec.receiveTime);
  const deliveryTimeOK =
    isValidTimeValue_(rec.deliveryTime);
  if (!receiveDateOK) missing.push('Ngày nhận hợp lệ');
  if (!receiveTimeOK) missing.push('Giờ nhận hợp lệ (00:00–23:59)');
  if (!deliveryDateOK) missing.push('Ngày giao hợp lệ');
  if (!deliveryTimeOK) missing.push('Giờ giao hợp lệ (00:00–23:59)');
  if (
    receiveDateOK &&
    deliveryDateOK
  ) {
    const receiveDay =
      new Date(
        rec.date.getFullYear(),
        rec.date.getMonth(),
        rec.date.getDate()
      ).getTime();
    const deliveryDay =
      new Date(
        rec.deliveryDate.getFullYear(),
        rec.deliveryDate.getMonth(),
        rec.deliveryDate.getDate()
      ).getTime();
    if (deliveryDay < receiveDay) {
      missing.push('Thời điểm giao không được trước thời điểm nhận');
    }
    else if (
      deliveryDay === receiveDay &&
      receiveTimeOK &&
      deliveryTimeOK
    ) {
      const receiveMinutes = timeToMinutes_(rec.receiveTime);
      const deliveryMinutes = timeToMinutes_(rec.deliveryTime);
      if (
        receiveMinutes !== null &&
        deliveryMinutes !== null &&
        deliveryMinutes < receiveMinutes
      ) {
        missing.push('Giờ giao không được trước giờ nhận trong cùng ngày');
      }
    }
  }
}
/*********************************************************
 * DATA EM
 *********************************************************/
function syncDataEM_(rec, workText) {
  syncDataForVisit_(rec, workText, true);
}
/*********************************************************
 * DATA SCC - SCHEMA NHÓM SCC
 * Không chèn/xóa cột; chỉ kiểm tra layout và chuẩn hóa H/O.
 *********************************************************/
function syncDataSCC_(rec, workText) {
  syncDataForVisit_(rec, workText, false);
}
/*********************************************************
 * TÌM DÒNG DATA THEO visitId DUY NHẤT
 * Không ghép theo biển số/ngày; gặp trùng ID phải báo lỗi.
 *********************************************************/
function findDataRow_(sh, rec) {
  if (!rec.visitId) throw new Error('Thiếu visitId; không tìm Data bằng biển số thay thế.');
  const count = sh.getLastRow() - 1;
  if (count < 1) return 0;
  const ids = sh.getRange(2, visitColumn_(sh.getName()), count, 1).getValues();
  let foundRow = 0;
  for (let i = 0; i < ids.length; i++) {
    if (clean_(ids[i][0]) !== rec.visitId) continue;
    if (foundRow) throw new Error(sh.getName() + ': nhiều dòng cùng visitId, cần kiểm tra trước khi đồng bộ.');
    foundRow = i + 2;
  }
  // P0.5.5: a matching UUID alone never authorizes a write into another car.
  // Old unguarded code would return this row, which downstream setValues() could overwrite.
  if (foundRow && rec && clean_(rec.plate)) {
    const stored = sh.getRange(foundRow, 2, 1, 2).getDisplayValues()[0];
    if (normalizePlate_(stored[0]) !== normalizePlate_(rec.plate) ||
        (clean_(rec.advisor) && normalizeKey_(stored[1]) !== normalizeKey_(rec.advisor))) {
      throw new Error('P055_STORED_VISIT_OWNER_CONFLICT: ' + sh.getName() +
        '!' + foundRow + ' / visitId=' + rec.visitId +
        ' / source=' + normalizePlate_(stored[0]) + '/' + stored[1] +
        ' / requested=' + normalizePlate_(rec.plate) + '/' + rec.advisor +
        '. BLOCK; do not replace stored work.');
    }
  }
  return foundRow;
}
/*********************************************************
 * XÓA DUPLICATE CHÍNH XÁC
 *
 * Chỉ xóa nếu trùng:
 * Ngày + Biển số + Giờ nhận
 *
 * Không xóa lượt xe khác giờ nhận.
 *********************************************************/

/*********************************************************
 * VALIDATE DATA TRƯỚC KHI DỌN
 *********************************************************/
function validateStoredDataBeforeClear_(
  rec,
  route
) {
  const errors =
    [];
  if (
    route.em.active
  ) {
    const e =
      validateStoredEM_(
        rec,
        route
      );
    if (
      e.length
    ) {
      errors.push(
        'Data EM thiếu: ' +
        e.join(', ')
      );
    }
  }
  if (
    route.scc.active
  ) {
    const e =
      validateStoredSCC_(
        rec,
        route
      );
    if (
      e.length
    ) {
      errors.push(
        'Data SCC thiếu: ' +
        e.join(', ')
      );
    }
  }
  return errors.concat(validateTrackingBeforeClear_(rec, route));
}
/*********************************************************
 * VALIDATE DATA EM
 *********************************************************/
function validateStoredEM_(
  rec,
  route
) {
  const sh =
    spreadsheetForRecord_(rec)
      .getSheetByName(
        CFG.SHEET_DATA_EM
      );
  if (!sh) {
    return [
      'Không tìm thấy sheet'
    ];
  }
  const row =
    findDataRow_(
      sh,
      rec
    );
  if (!row) {
    return [
      'Không có dòng Data'
    ];
  }
  const v =
    sh
      .getRange(
        row,
        1,
        1,
        14
      )
      .getValues()[0];
  const d =
    sh
      .getRange(
        row,
        1,
        1,
        14
      )
      .getDisplayValues()[0];
  const missing = [];
  addMissingCommon_(rec, missing);
  addInspectionAssignmentMissing_(rec, missing);
  if (!v[0]) {
    missing.push('Ngày');
  }
  if (!clean_(d[1])) {
    missing.push('Biển số');
  }
  if (!clean_(d[2])) {
    missing.push('CVDV');
  }
  if (!clean_(d[3])) {
    missing.push('Model');
  }
  // P0.5.3.1: Data EM cot E duoc de trong trong luot Mr. Khoi-only.
  // Nguoi thuc hien duoc xac nhan tai Dieu phoi cot F cung visitId.
  // Giu nguyen cac chot khac: trang thai/giay gio hoan thanh va visitId.
  if (!clean_(d[4]) && !p0531MrKhoiEmOnly_(rec)) {
    missing.push('Nhóm EM');
  }
  if (!clean_(d[5])) {
    missing.push(
      'Công việc EM'
    );
  }
  const needsBDLevel =
    route.em.items
      .indexOf('EM60') !== -1 ||
    route.em.items
      .indexOf('BD++') !== -1;
  if (
    needsBDLevel &&
    !clean_(d[6])
  ) {
    missing.push(
      'Cấp BD'
    );
  }
  if (
    v[8] === ''
  ) {
    missing.push('Giờ nhận');
  }
  if (
    v[9] === ''
  ) {
    missing.push('Ngày giao');
  }
  if (
    v[10] === ''
  ) {
    missing.push('Giờ giao');
  }
  if (
    v[11] === ''
  ) {
    missing.push(
      'Giờ hoàn thành EM'
    );
  }
  if (
    !isComplete_(
      d[12]
    )
  ) {
    missing.push(
      'Tiến độ cuối'
    );
  }
  return uniqueArray_(
    missing
  );
}
/*********************************************************
 * VALIDATE DATA SCC
 *********************************************************/
function validateStoredSCC_(
  rec,
  route
) {
  const sh = spreadsheetForRecord_(rec).getSheetByName(CFG.SHEET_DATA_SCC);
  if (!sh) return ['Không tìm thấy sheet'];
  const row = findDataRow_(sh, rec);
  if (!row) return ['Không có dòng Data'];
  const width = CFG.DATA_SCC.VISIT_ID;
  const v = sh.getRange(row, 1, 1, width).getValues()[0];
  const d = sh.getRange(row, 1, 1, width).getDisplayValues()[0];
  const missing = [];
  addMissingCommon_(rec, missing);
  if (!v[CFG.DATA_SCC.NGAY - 1]) missing.push('Ngày');
  if (!clean_(d[CFG.DATA_SCC.BIEN_SO - 1])) missing.push('Biển số');
  if (!clean_(d[CFG.DATA_SCC.CVDV - 1])) missing.push('CVDV');
  if (!clean_(d[CFG.DATA_SCC.MODEL - 1])) missing.push('Model');
  if (!clean_(d[CFG.DATA_SCC.CONG_VIEC - 1])) missing.push('Công việc SCC');
  const inspectionJob =
    route.scc.items.indexOf('Kiểm tra') !== -1 ||
    route.scc.items.indexOf('Kiểm tra lại') !== -1;
  const inspector = clean_(d[CFG.DATA_SCC.NGUOI_KT - 1]);
  const sccGroup = clean_(d[CFG.DATA_SCC.NHOM_SCC - 1]);
  const hasRepeatInspection = route.scc.items.indexOf('Kiểm tra lại') !== -1;
  if (hasRepeatInspection) {
    if (!inspector) missing.push('Người KT');
  } else if (inspectionJob) {
    if (!inspector && !sccGroup) missing.push('Người/nhóm thực hiện kiểm tra');
  } else if (!sccGroup) {
    missing.push('Nhóm SCC');
  }
  if (v[CFG.DATA_SCC.GIO_NHAN - 1] === '') missing.push('Giờ nhận');
  if (v[CFG.DATA_SCC.NGAY_GIAO - 1] === '') missing.push('Ngày giao');
  if (v[CFG.DATA_SCC.GIO_GIAO - 1] === '') missing.push('Giờ giao');
  if (v[CFG.DATA_SCC.GIO_HOAN_THANH - 1] === '') missing.push('Giờ hoàn thành SCC');
  if (!isComplete_(d[CFG.DATA_SCC.TIEN_DO_CUOI - 1])) missing.push('Tiến độ cuối');
  return uniqueArray_(missing);
}
/*********************************************************
 * THEO DÕI KT
 *********************************************************/function syncTrackingSheets_(rec, route, e) {
  const targets = trackingTargets_(rec, route);
  targets.forEach(t => {
    const status = t.side === 'EM' ? rec.emStatus : rec.sccStatus;
    const previous = t.side === 'EM' ? rec.lastEMStatus : rec.lastSCCStatus;
    syncOneTracking_(t.sheet, rec, t.repeat, previous || '', status, status !== previous, t.side);
  });
  // Lượt Web mới chưa thể có tracking lịch sử mang UUID vừa tạo.
  if (e && e.isNewWebAppInsert) return;
  // Giữ lịch sử cũ, cập nhật nhận diện cùng lượt; đóng phiên nếu đổi bên thực hiện.
  [CFG.SHEET_KT, CFG.SHEET_KT_LAI].forEach(name => {
    const sh = spreadsheetForRecord_(rec).getSheetByName(name);
    if (!sh) return;
    const col = visitColumn_(name), last = sh.getLastRow();
    if (last < 2) return;
    sh.getRange(2, col, last - 1, 2).getValues().forEach((v, i) => {
      if (v[0] !== rec.visitId) return;
      const row = i + 2, side = v[1];
      sh.getRange(row, 1, 1, 4).setValues([[rec.date, plateAsText_(rec.plate), rec.advisor, rec.model]]);
      if (!targets.some(t => t.sheet === name && t.side === side) && sh.getRange(row, 7).getValue() && !sh.getRange(row, 8).getValue()) sh.getRange(row, 8).setValue(new Date());
    });
  });
}
/*********************************************************
 * SYNC TRACKING
 *********************************************************/
function syncOneTracking_(sheetName, rec, isRepeat, oldStatus, newStatus, statusWasEdited, side) {
  side = side || 'SCC';
  const sh = spreadsheetForRecord_(rec).getSheetByName(sheetName);
  if (!sh) throw new Error('Không tìm thấy ' + sheetName);
  const col = visitColumn_(sheetName);
  let row = findTrackingRow_(sh, rec, isRepeat, side);
  const running = trackingIsRunning_(newStatus, side);
  const entering = running && !trackingIsRunning_(oldStatus, side);
  if (row && sh.getRange(row, 8).getValue() && entering) row = 0;
  const performer = side === 'EM' ? rec.emGroup : (rec.inspector || rec.ktv1 || rec.ktv2);
  if (!row) {
    row = Math.max(sh.getLastRow() + 1, 2);
    const values = Array(col + 1).fill('');
    [rec.date, plateAsText_(rec.plate), rec.advisor, rec.model, rec.inspectionText || '', performer, running ? new Date() : '', ''].forEach((v,i) => {values[i] = v;});
    if (isRepeat) values[8] = 'Lần ' + nextRepeatNumber_(sh, rec.plate, rec.visitId);
    values[col - 1] = rec.visitId;
    values[col] = side;
    preparePlateCell_(sh, row);
    sh.getRange(row, 1, 1, values.length).setValues([values]);
  } else {
    const endTime = sh.getRange(row, 8).getValue();
    // Chỉ phiên đang mở mới nhận dữ liệu cập nhật từ Điều phối.
    // Phiên đã đóng giữ nguyên lịch sử.
    if (!endTime) {
      sh.getRange(row, 1, 1, 4).setValues([[rec.date, plateAsText_(rec.plate), rec.advisor, rec.model]]);
      sh.getRange(row, 5).setValue(rec.inspectionText || '');
      sh.getRange(row, 6).setValue(performer);
    }
    if (running && !sh.getRange(row, 7).getValue() && !endTime) {
      sh.getRange(row, 7).setValue(new Date());
    }
  }
  // Phiên có giờ bắt đầu mà chưa kết thúc là bằng chứng, kể cả paste thiếu oldValue.
  if ((!running && sh.getRange(row, 7).getValue()) || isComplete_(newStatus)) {
    if (!sh.getRange(row, 8).getValue()) {
      const confirmed = isComplete_(newStatus)
        ? p042ResolvedCompletionStamp_(rec, side === 'EM') : new Date();
      sh.getRange(row, 8).setValue(confirmed);
      if (isComplete_(newStatus)) {
        const source = confirmed && confirmed.p0491Proof;
        sh.getRange(row, 8).setNote(source
          ? 'P0.4.9.1: Đóng phiên theo mốc hoàn thành công việc có xác nhận người vận hành ' +
            source.time + ' tại ' + source.sheet + '!L' + source.row +
            ' / visitId=' + source.visitId +
            '. KHÔNG khẳng định đây là giờ kết thúc kiểm tra được đo độc lập.'
          : 'P0.4.2: mốc xác nhận hoàn thành trên Điều phối, không phải giờ kết thúc thực tế riêng của Kiểm tra.');
      }
    }
  }
  sh.getRange(row, 1).setNumberFormat('dd/MM/yyyy');
  sh.getRange(row, 7, 1, 2).setNumberFormat('HH:mm');
  preparePlateCell_(sh, row);
}
/*********************************************************
 * FIND TRACKING
 *********************************************************/
function findTrackingRow_(sh, rec, isRepeat, side) {
  if (!rec.visitId) throw new Error('Thiếu visitId tracking.');
  side = side || 'SCC';
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const values = sh.getRange(2, visitColumn_(sh.getName()), last - 1, 2).getValues();
  let selectedRow = 0;
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== rec.visitId || values[i][1] !== side) continue;
    // Keep legacy newest-first semantics for repeated inspections, but block
    // a visitId that belongs to a DIFFERENT plate/advisor in tracking.
    const stored=sh.getRange(i+2, 2, 1, 2).getDisplayValues()[0];
    if ((clean_(rec.plate) && normalizePlate_(stored[0]) !== normalizePlate_(rec.plate)) ||
        (clean_(rec.advisor) && normalizeKey_(stored[1]) !== normalizeKey_(rec.advisor))) {
      throw new Error('P055_TRACKING_VISIT_OWNER_CONFLICT: '+sh.getName()+
        '!'+(i+2)+' / visitId='+rec.visitId+' / side='+side);
    }
    if (!selectedRow) selectedRow=i+2;
  }
  return selectedRow;
}
/*********************************************************
 * SỐ LẦN KT LẠI
 *********************************************************/
function nextRepeatNumber_(sh, plate, visitId) {
  const last = sh.getLastRow();
  if (last < 2) return 2;
  const col = visitColumn_(sh.getName());
  const values = sh.getRange(2, 1, last - 1, col).getDisplayValues();
  const wanted = normalizePlate_(plate);
  let max = 1;
  values.forEach(row => {
    const sameVisit =
      visitId
        ? clean_(row[col - 1]) === visitId
        : normalizePlate_(row[1]) === wanted;
    if (!sameVisit) return;
    const match = String(row[8] || '').match(/(\d+)/);
    if (match) max = Math.max(max, Number(match[1]));
  });
  return max + 1;
}
/*********************************************************
 * READ DISPATCH
 *********************************************************/
function readDispatchRow_(
  sh,
  row
) {
  const c =
    CFG.DP;
  const range =
    sh.getRange(
      row,
      1,
      1,
      21
    );
  const v =
    range.getValues()[0];
  const d =
    range.getDisplayValues()[0];
  return {
    ss: sh.getParent(),
    visitId: ensureVisitId_(sh, row),
    lastEMStatus: clean_(sh.getRange(row, c.LAST_EM_STATUS).getValue()),
    lastSCCStatus: clean_(sh.getRange(row, c.LAST_SCC_STATUS).getValue()),
    emConfirmationStamp: v[c.MOC_TD_EM - 1],
    sccConfirmationStamp: v[c.MOC_TD_SCC - 1],
    row:
      row,
    date:
      v[c.NGAY - 1],
    plate:
      readPlateText_(
        sh,
        row
      ),
    advisor:
      clean_(
        d[c.CVDV - 1]
      ),
    model:
      clean_(
        d[c.MODEL - 1]
      ),
    emGroup:
      clean_(
        d[c.NHOM_EM - 1]
      ),
    inspector:
      clean_(
        d[c.NGUOI_KT - 1]
      ),
    ktv1:
      clean_(
        d[c.KTV1 - 1]
      ),
    ktv2:
      clean_(
        d[c.KTV2 - 1]
      ),
    workRaw:
      clean_(
        d[c.CV - 1]
      ),
    bdLevel:
      clean_(
        d[c.CAP_BD - 1]
      ),
    gtgt:
      clean_(
        d[c.GTGT - 1]
      ),
    itemCount:
      v[c.SO_MON - 1],
    receiveTime:
      v[c.GIO_NHAN - 1],
    deliveryDate:
      v[c.NGAY_GIAO - 1],
    deliveryTime:
      v[c.GIO_GIAO - 1],
    emStatus:
      clean_(
        d[c.TD_EM - 1]
      ),
    sccStatus:
      clean_(
        d[c.TD_SCC - 1]
      ),
    washStatus:
      clean_(
        d[c.TD_RUA_XE - 1]
      ),
    note:
      clean_(
        d[c.GHI_CHU - 1]
      ),
    inspectionText:
      cleanInspectionNote_(
        sh
          .getRange(
            row,
            c.CV
          )
          .getNote()
      ),
    inspectionType:
      inspectionTypeFromNote_(
        sh
          .getRange(
            row,
            c.CV
          )
          .getNote()
      ),
    qcLevel:
      qcLevelFromNote_(
        sh
          .getRange(
            row,
            c.CV
          )
          .getNote()
      )
  };
}
/*********************************************************
 * ĐỌC NỘI DUNG KIỂM TRA TỪ NOTE Ô CV
 *
 * Web App lưu:
 * NỘI DUNG KIỂM TRA:
 * <nội dung>
 *********************************************************/
function cleanInspectionNote_(
  value
) {
  let text =
    String(
      value || ''
    )
      .replace(
        /\r/g,
        '\n'
      )
      .trim();
  if (!text) {
    return '';
  }
  text =
    text
      .replace(
        /^LOẠI KIỂM TRA:\s*[^\n]*\n?/i,
        ''
      )
      .replace(
        /^CẤP QC:\s*[123]\s*\n?/im,
        ''
      )
      .replace(
        /^NỘI DUNG KIỂM TRA:\s*/i,
        ''
      );
  const lines = [];
  text
    .split('\n')
    .forEach(function(rawLine) {
      // Tách theo | để xóa đúng segment "Yêu cầu rửa xe"
      // nhưng vẫn giữ phần kiểm tra nằm sau segment đó.
      const parts =
        String(rawLine || '')
          .split('|');
      const kept =
        parts
          .map(function(part) {
            return clean_(part);
          })
          .filter(function(part) {
            if (!part) return false;
            const key =
              normalizeText_(
                part
              );
            if (key.indexOf('YEU CAU RUA XE') >= 0) return false;
            if (/^(X|T|YES|NO|CO|KHONG)$/.test(key)) return false;
            return true;
          });
      const line =
        clean_(
          kept.join(' | ')
        );
      if (line) {
        lines.push(line);
      }
    });
  return uniqueArray_(
    lines
  ).join(
    '\n'
  );
}
function inspectionTypeFromNote_(
  value
) {
  const text =
    String(
      value || ''
    )
      .replace(
        /\r/g,
        ''
      )
      .trim();
  const match =
    text.match(
      /^LOẠI KIỂM TRA:\s*(Kiểm tra lại|Kiểm tra)\s*$/im
    );
  if (
    match
  ) {
    return clean_(
      match[1]
    );
  }
  if (
    /^NỘI DUNG KIỂM TRA:/im.test(
      text
    )
  ) {
    return 'Kiểm tra';
  }
  return '';
}
/*********************************************************
 * ĐỌC CẤP QC TỪ NOTE Ô CV
 *
 * Web App lưu:
 * CẤP QC: 1 / 2 / 3
 *
 * Rule nghiệp vụ:
 * - QC1 / QC2: không tạo/liên kết Theo dõi KT.
 * - QC3: tạo/liên kết Theo dõi KT.
 * - Người KT chỉ là thông tin phân công, không override Cấp QC.
 * - Kiểm tra lại giữ logic Theo dõi KT lại hiện hữu.
 *********************************************************/
function qcLevelFromNote_(value) {
  const text =
    String(value || '')
      .replace(/\r/g, '')
      .trim();
  const match =
    text.match(
      /^CẤP QC:\s*([123])\s*$/im
    );
  return match ? String(match[1]) : '';
}
/*********************************************************
 * BIỂN SỐ - CORE
 *********************************************************/
function plateAsText_(
  value
) {
  return String(
    normalizePlate_(
      value
    )
  );
}
function preparePlateCell_(
  sh,
  row
) {
  sh
    .getRange(
      row,
      2
    )
    .setNumberFormat(
      '@'
    );
}

function readPlateText_(
  sh,
  row
) {
  const cell =
    sh.getRange(
      row,
      2
    );
  return normalizePlate_(
    cell.getDisplayValue()
  );
}
/*********************************************************
 * EDIT BIỂN SỐ TRÊN ĐIỀU PHỐI
 *********************************************************/
function normalizePlateCellFromEvent_(
  sh,
  row,
  e
) {
  const cell =
    sh.getRange(
      row,
      2
    );
  cell
    .setNumberFormat(
      '@'
    );
  let raw =
    '';
  const singleCell =
    e &&
    e.range &&
    e.range.getNumRows() === 1 &&
    e.range.getNumColumns() === 1 &&
    e.range.getRow() === row &&
    e.range.getColumn() === 2;
  if (
    singleCell &&
    typeof e.value !==
      'undefined'
  ) {
    raw =
      String(
        e.value
      );
  } else {
    raw =
      cell.getDisplayValue();
  }
  const normalized =
    normalizePlate_(
      raw
    );
  if (!normalized) {
    return;
  }
  cell
    .setNumberFormat(
      '@'
    )
    .setValue(
      String(
        normalized
      )
    )
    .setNumberFormat(
      '@'
    );
}
/*********************************************************
 * EDIT BIỂN SỐ TRÊN MỌI SHEET
 *********************************************************/
function normalizeEditedPlateRange_(
  e
) {
  const sh =
    e.range.getSheet();
  const startRow =
    e.range.getRow();
  const rows =
    e.range.getNumRows();
  for (
    let i = 0;
    i < rows;
    i++
  ) {
    const row =
      startRow + i;
    if (
      row < plateStartRow_(
        sh.getName()
      )
    ) {
      continue;
    }
    const cell =
      sh.getRange(
        row,
        2
      );
    cell
      .setNumberFormat(
        '@'
      );
    let raw =
      cell.getDisplayValue();
    if (
      rows === 1 &&
      e.range.getNumColumns() === 1 &&
      e.range.getColumn() === 2 &&
      typeof e.value !==
        'undefined'
    ) {
      raw =
        String(
          e.value
        );
    }
    const normalized =
      normalizePlate_(
        raw
      );
    if (
      normalized
    ) {
      cell
        .setNumberFormat(
          '@'
        )
        .setValue(
          String(
            normalized
          )
        )
        .setNumberFormat(
          '@'
        );
    }
  }
}
/*********************************************************
 * KHÓA CỘT BIỂN SỐ TOÀN HỆ THỐNG
 *********************************************************/
function ensureAllPlateColumnsText_() {
  const ss =
    SpreadsheetApp.getActive();
  const configs = [
    [
      CFG.SHEET_DISPATCH,
      CFG.FIRST_DATA_ROW
    ],
    [
      CFG.SHEET_DATA_EM,
      2
    ],
    [
      CFG.SHEET_DATA_SCC,
      2
    ],
    [
      CFG.SHEET_KT,
      2
    ],
    [
      CFG.SHEET_KT_LAI,
      2
    ]
  ];
  configs.forEach(config => {
    const sh =
      ss.getSheetByName(
        config[0]
      );
    if (!sh) return;
    const start =
      config[1];
    const max =
      sh.getMaxRows();
    if (
      max < start
    ) {
      return;
    }
    sh
      .getRange(
        start,
        2,
        max - start + 1,
        1
      )
      .setNumberFormat(
        '@'
      );
  });
}
/*********************************************************
 * PLATE SHEET CONFIG
 *********************************************************/
function isProtectedPlateSheet_(
  sheetName
) {
  return [
    CFG.SHEET_DISPATCH,
    CFG.SHEET_DATA_EM,
    CFG.SHEET_DATA_SCC,
    CFG.SHEET_KT,
    CFG.SHEET_KT_LAI
  ]
    .indexOf(
      sheetName
    ) !== -1;
}
function plateStartRow_(
  sheetName
) {
  if (
    sheetName ===
    CFG.SHEET_DISPATCH
  ) {
    return CFG.FIRST_DATA_ROW;
  }
  return 2;
}/*********************************************************
 * V13.14 - HẸN GIAO ĐIỀU PHỐI -> CVDV
 *********************************************************/
/**
 * Menu thủ công:
 * Điều phối -> Đồng bộ Hẹn giao → CVDV
 */
function dongBoHenGiaoDongDangChon() {
  const ss =
    SpreadsheetApp.getActive();
  const sh =
    ss.getActiveSheet();
  const range =
    ss.getActiveRange();
  if (
    !sh ||
    sh.getName() !==
      CFG.SHEET_DISPATCH ||
    !range
  ) {
    SpreadsheetApp
      .getUi()
      .alert(
        'Hãy chọn một dòng trong sheet Điều phối.'
      );
    return;
  }
  const row =
    range.getRow();
  if (
    row <
    CFG.FIRST_DATA_ROW
  ) {
    SpreadsheetApp
      .getUi()
      .alert(
        'Hãy chọn một dòng dữ liệu từ dòng ' +
        CFG.FIRST_DATA_ROW +
        ' trở xuống.'
      );
    return;
  }
  const result =
    syncHenGiaoCVDVTuDieuPhoi_(
      sh,
      row,
      {
        source: 'MANUAL',
        manual: true
      }
    );
  const message =
    result && result.message
      ? result.message
      : 'Đã xử lý đồng bộ Hẹn giao.';
  try {
    ss.toast(
      message,
      result &&
      result.success
        ? '✅ HẸN GIAO CVDV'
        : '⚠️ HẸN GIAO CVDV',
      8
    );
  }
  catch (err) {}
  return result;
}
/**
 * Core sync 1 dòng Điều phối -> đúng 1 sheet CVDV.
 *
 * Rule:
 * - Không quét 9 sheet.
 * - Cột C Điều phối quyết định sheet CVDV.
 * - Trong sheet đích chỉ match Ngày tiếp nhận + BSX.
 * - 1 match  -> update G.
 * - 0 match  -> không tạo mới.
 * - >1 match -> BLOCK.
 * - O trống  -> không clear Hẹn giao cũ.
 */
function syncHenGiaoCVDVTuDieuPhoi_(
  dispatchSheet,
  row,
  options
) {
  options =
    options || {};
  if (
    !dispatchSheet ||
    dispatchSheet.getName() !==
      CFG.SHEET_DISPATCH
  ) {
    throw new Error(
      'Không đúng sheet Điều phối.'
    );
  }
  if (
    row <
    CFG.FIRST_DATA_ROW
  ) {
    return {
      success: false,
      skipped: true,
      code: 'ROW_INVALID',
      message: 'Dòng Điều phối không hợp lệ.'
    };
  }
  const ss =
    dispatchSheet.getParent();
  const values =
    dispatchSheet
      .getRange(
        row,
        1,
        1,
        CFG.DP.GIO_GIAO
      )
      .getValues()[0];
  const display =
    dispatchSheet
      .getRange(
        row,
        1,
        1,
        CFG.DP.GIO_GIAO
      )
      .getDisplayValues()[0];
  const receiveDate =
    values[
      CFG.DP.NGAY - 1
    ];
  const plate =
    normalizePlate_(
      display[
        CFG.DP.BIEN_SO - 1
      ]
    );
  const advisorRaw =
    clean_(
      display[
        CFG.DP.CVDV - 1
      ]
    );
  const deliveryDate =
    values[
      CFG.DP.NGAY_GIAO - 1
    ];
  const deliveryTime =
    values[
      CFG.DP.GIO_GIAO - 1
    ];
  const source =
    clean_(
      options.source || 'AUTO'
    );
  const baseLog = {
    row: row,
    receiveDate: receiveDate,
    plate: plate,
    advisor: advisorRaw,
    deliveryDate: deliveryDate,
    deliveryTime: deliveryTime,
    source: source
  };
  // ---------------------------------------------------
  // O trống -> KHÔNG clear G CVDV.
  // ---------------------------------------------------
  if (
    deliveryTime === '' ||
    deliveryTime === null
  ) {
    return {
      success: true,
      skipped: true,
      code: 'NO_DELIVERY_TIME',
      message:
        'Giờ giao đang trống nên không thay đổi Hẹn giao CVDV.'
    };
  }
  if (!plate) {
    ghiLogHenGiaoCVDV_(
      ss,
      Object.assign(
        {},
        baseLog,
        {
          result: 'INVALID_PLATE',
          detail:
            'Thiếu/không đọc được Biển số Điều phối.'
        }
      )
    );
    return {
      success: false,
      code: 'INVALID_PLATE',
      message: 'Thiếu Biển số.'
    };
  }
  if (
    !(receiveDate instanceof Date) ||
    isNaN(
      receiveDate.getTime()
    )
  ) {
    ghiLogHenGiaoCVDV_(
      ss,
      Object.assign(
        {},
        baseLog,
        {
          result: 'INVALID_RECEIVE_DATE',
          detail:
            'Ngày tiếp nhận Điều phối không hợp lệ.'
        }
      )
    );
    return {
      success: false,
      code: 'INVALID_RECEIVE_DATE',
      message:
        'Ngày tiếp nhận không hợp lệ.'
    };
  }
  // P056: direct fallback also respects exact Reception Không DV.
  const p056Gate = p056ReceptionGateReadOnly_(receiveDate, plate, ss);
  if (p056Gate.blocked) {
    console.warn('[P056_NO_SERVICE_PROMISE_BLOCK] ' + JSON.stringify(p056Gate));
    return {success:false,blocked:true,code:p056Gate.code,
      message:'Không đồng bộ Hẹn giao: ' + p056Gate.code};
  }
  const targetSheetName =
    mapAdvisorToCvdvSheet_(
      advisorRaw
    );
  if (!targetSheetName) {
    ghiLogHenGiaoCVDV_(
      ss,
      Object.assign(
        {},
        baseLog,
        {
          result: 'ADVISOR_INVALID',
          detail:
            'Không map được CVDV "' +
            advisorRaw +
            '".'
        }
      )
    );
    return {
      success: false,
      code: 'ADVISOR_INVALID',
      message:
        'Không map được CVDV "' +
        advisorRaw +
        '".'
    };
  }
  if (
    !(deliveryDate instanceof Date) ||
    isNaN(
      deliveryDate.getTime()
    )
  ) {
    ghiLogHenGiaoCVDV_(
      ss,
      Object.assign(
        {},
        baseLog,
        {
          targetSheet:
            targetSheetName,
          result:
            'INVALID_DELIVERY_DATE',
          detail:
            'Có Giờ giao nhưng Ngày giao không hợp lệ.'
        }
      )
    );
    return {
      success: false,
      code: 'INVALID_DELIVERY_DATE',
      message:
        'Ngày giao không hợp lệ.'
    };
  }
  const minutes =
    timeToMinutes_(
      deliveryTime
    );
  if (
    minutes === null ||
    !Number.isFinite(
      minutes
    )
  ) {
    ghiLogHenGiaoCVDV_(
      ss,
      Object.assign(
        {},
        baseLog,
        {
          targetSheet:
            targetSheetName,
          result:
            'INVALID_DELIVERY_TIME',
          detail:
            'Giờ giao không hợp lệ.'
        }
      )
    );
    return {
      success: false,
      code: 'INVALID_DELIVERY_TIME',
      message:
        'Giờ giao không hợp lệ.'
    };
  }
  const promiseDateTime =
    combineDeliveryPromiseDateTime_(
      deliveryDate,
      minutes
    );
  const cvdvSS =
    SpreadsheetApp.openById(
      DELIVERY_PROMISE_SYNC_CFG
        .CVDV_SPREADSHEET_ID
    );
  const targetSheet =
    cvdvSS.getSheetByName(
      targetSheetName
    );
  if (!targetSheet) {
    ghiLogHenGiaoCVDV_(
      ss,
      Object.assign(
        {},
        baseLog,
        {
          targetSheet:
            targetSheetName,
          result:
            'TARGET_SHEET_MISSING',
          detail:
            'Không tìm thấy sheet CVDV đích.'
        }
      )
    );
    return {
      success: false,
      code: 'TARGET_SHEET_MISSING',
      message:
        'Không tìm thấy sheet ' +
        targetSheetName +
        '.'
    };
  }
  const matches =
    findCvdvPromiseMatchesInOneSheet_(
      targetSheet,
      receiveDate,
      plate
    );
  if (
    matches.length === 0
  ) {
    ghiLogHenGiaoCVDV_(
      ss,
      Object.assign(
        {},
        baseLog,
        {
          targetSheet:
            targetSheetName,
          matchCount: 0,
          result:
            'NO_MATCH',
          detail:
            'Không tìm thấy Ngày + BSX trong sheet ' +
            targetSheetName +
            '. Không tạo record mới.'
        }
      )
    );
    return {
      success: false,
      code: 'NO_MATCH',
      message:
        'Không tìm thấy ' +
        plate +
        ' trong sheet ' +
        targetSheetName +
        ' đúng ngày tiếp nhận.'
    };
  }
  if (
    matches.length > 1
  ) {
    ghiLogHenGiaoCVDV_(
      ss,
      Object.assign(
        {},
        baseLog,
        {
          targetSheet:
            targetSheetName,
          matchCount:
            matches.length,
          result:
            'MULTI_MATCH',
          detail:
            'BLOCK: cùng Ngày + BSX xuất hiện tại dòng ' +
            matches.join(', ') +
            '.'
        }
      )
    );
    return {
      success: false,
      blocked: true,
      code: 'MULTI_MATCH',
      message:
        'Có ' +
        matches.length +
        ' record trùng Ngày + BSX trong sheet ' +
        targetSheetName +
        '.'
    };
  }
  const targetRow =
    matches[0];
  // Re-check key ngay trước khi ghi.
  const verify =
    targetSheet
      .getRange(
        targetRow,
        1,
        1,
        3
      )
      .getValues()[0];
  if (
    !sameDate_(
      verify[0],
      receiveDate
    ) ||
    normalizePlate_(
      verify[2]
    ) !== plate
  ) {
    ghiLogHenGiaoCVDV_(
      ss,
      Object.assign(
        {},
        baseLog,
        {
          targetSheet:
            targetSheetName,
          targetRow:
            targetRow,
          matchCount: 1,
          result:
            'VERIFY_CHANGED',
          detail:
            'Record đích thay đổi trước lúc ghi. Không cập nhật.'
        }
      )
    );
    return {
      success: false,
      blocked: true,
      code: 'VERIFY_CHANGED',
      message:
        'Record CVDV vừa thay đổi; chưa cập nhật.'
    };
  }
  const promiseCell =
    targetSheet.getRange(
      targetRow,
      mainQtCol_(targetSheet,DELIVERY_PROMISE_SYNC_CFG.CVDV_COL_PROMISE)
    );
  const oldPromise =
    promiseCell.getValue();
  // V13.14.1:
  // Chỉ đổi cách HIỂN THỊ, vẫn giữ đầy đủ DateTime trong giá trị ô.
  // Áp format trước cả nhánh NO_CHANGE để dòng đã sync ở V13.14.0
  // cũng được sửa hiển thị khi bấm đồng bộ lại.
  promiseCell
    .setNumberFormat(
      'dd/MM HH:mm'
    );
  if (
    samePromiseDateTime_(
      oldPromise,
      promiseDateTime
    )
  ) {
    return {
      success: true,
      noChange: true,
      code: 'NO_CHANGE',
      targetSheet:
        targetSheetName,
      targetRow:
        targetRow,
      message:
        'Hẹn giao CVDV đã đúng; đã chuẩn hóa hiển thị dd/MM HH:mm.'
    };
  }
  promiseCell
    .setValue(
      promiseDateTime
    )
    .setNumberFormat(
      'dd/MM HH:mm'
    );
  ghiLogHenGiaoCVDV_(
    ss,
    Object.assign(
      {},
      baseLog,
      {
        targetSheet:
          targetSheetName,
        targetRow:
          targetRow,
        matchCount: 1,
        result:
          'UPDATED',
        detail:
          'Đã cập nhật HẸN GIAO CVDV.'
      }
    )
  );
  return {
    success: true,
    updated: true,
    code: 'UPDATED',
    targetSheet:
      targetSheetName,
    targetRow:
      targetRow,
    promise:
      promiseDateTime,
    message:
      'Đã cập nhật ' +
      targetSheetName +
      '!G' +
      targetRow +
      ' = ' +
      Utilities.formatDate(
        promiseDateTime,
        ss.getSpreadsheetTimeZone(),
        'dd/MM HH:mm'
      )
  };
}
function syncHenGiaoCVDVTuXeChoGiao_(waitingSheet, row, options) {
  options = options || {};
  if (
    !waitingSheet ||
    waitingSheet.getName() !== CFG.SHEET_WAITING_DELIVERY
  ) {
    throw new Error('Không đúng sheet Xe chờ giao.');
  }
  if (row < CFG.WAITING_DELIVERY_FIRST_ROW) {
    return {
      success: false,
      skipped: true,
      code: 'ROW_INVALID',
      message: 'Dòng Xe chờ giao không hợp lệ.'
    };
  }

  const ss = waitingSheet.getParent();
  const values = waitingSheet
    .getRange(row, 1, 1, CFG.WD.GIO_GIAO)
    .getValues()[0];
  const display = waitingSheet
    .getRange(row, 1, 1, CFG.WD.GIO_GIAO)
    .getDisplayValues()[0];

  const receiveDate = values[CFG.WD.NGAY - 1];
  const plate = normalizePlate_(display[CFG.WD.BIEN_SO - 1]);
  const advisorRaw = clean_(display[CFG.WD.CVDV - 1]);
  const deliveryDate = values[CFG.WD.NGAY_GIAO - 1];
  const deliveryTime = values[CFG.WD.GIO_GIAO - 1];
  const source = clean_(options.source || 'WAITING_AUTO');
  const targetSheetName = mapAdvisorToCvdvSheet_(advisorRaw);

  if (
    !plate ||
    !(receiveDate instanceof Date) ||
    isNaN(receiveDate.getTime()) ||
    !targetSheetName
  ) {
    return {
      success: false,
      skipped: true,
      code: 'IDENTITY_INCOMPLETE',
      message: 'Thiếu Ngày / BSX / CVDV hợp lệ ở Xe chờ giao.'
    };
  }
  if (deliveryTime === '' || deliveryTime === null) {
    return {
      success: true,
      skipped: true,
      code: 'NO_DELIVERY_TIME',
      message: 'Giờ giao đang trống nên không thay đổi Hẹn giao CVDV.'
    };
  }
  // P056: direct waiting fallback may not create/modify CVDV promise for Không DV.
  const p056Gate = p056ReceptionGateReadOnly_(receiveDate, plate, ss);
  if (p056Gate.blocked) {
    console.warn('[P056_NO_SERVICE_WAITING_PROMISE_BLOCK] ' + JSON.stringify(p056Gate));
    return {success:false,blocked:true,code:p056Gate.code,
      message:'Không đồng bộ Hẹn giao từ Chờ giao: ' + p056Gate.code};
  }
  if (!(deliveryDate instanceof Date) || isNaN(deliveryDate.getTime())) {
    return {
      success: false,
      code: 'INVALID_DELIVERY_DATE',
      message: 'Ngày giao Xe chờ giao không hợp lệ.'
    };
  }
  const minutes = timeToMinutes_(deliveryTime);
  if (minutes === null || !Number.isFinite(minutes)) {
    return {
      success: false,
      code: 'INVALID_DELIVERY_TIME',
      message: 'Giờ giao Xe chờ giao không hợp lệ.'
    };
  }

  const promiseDateTime = combineDeliveryPromiseDateTime_(deliveryDate, minutes);
  const cvdvSS = SpreadsheetApp.openById(
    DELIVERY_PROMISE_SYNC_CFG.CVDV_SPREADSHEET_ID
  );
  const targetSheet = cvdvSS.getSheetByName(targetSheetName);
  if (!targetSheet) {
    return {
      success: false,
      code: 'TARGET_SHEET_MISSING',
      message: 'Không tìm thấy sheet CVDV ' + targetSheetName + '.'
    };
  }
  const matches = findCvdvPromiseMatchesInOneSheet_(
    targetSheet,
    receiveDate,
    plate
  );
  if (matches.length === 0) {
    return {
      success: false,
      code: 'NO_MATCH',
      message: 'Không tìm thấy Ngày + BSX ở sheet ' + targetSheetName + '.'
    };
  }
  if (matches.length > 1) {
    return {
      success: false,
      blocked: true,
      code: 'MULTI_MATCH',
      message: 'Có nhiều record cùng Ngày + BSX trong sheet ' + targetSheetName + '.'
    };
  }

  const targetRow = matches[0];
  const verify = targetSheet.getRange(targetRow, 1, 1, 3).getValues()[0];
  if (
    !sameDate_(verify[0], receiveDate) ||
    normalizePlate_(verify[2]) !== plate
  ) {
    return {
      success: false,
      blocked: true,
      code: 'VERIFY_CHANGED',
      message: 'Record CVDV thay đổi trước lúc ghi Hẹn giao.'
    };
  }

  const promiseCell = targetSheet.getRange(
    targetRow,
    mainQtCol_(targetSheet,DELIVERY_PROMISE_SYNC_CFG.CVDV_COL_PROMISE)
  );
  const oldPromise = promiseCell.getValue();
  promiseCell.setNumberFormat('dd/MM HH:mm');
  if (samePromiseDateTime_(oldPromise, promiseDateTime)) {
    return {
      success: true,
      noChange: true,
      code: 'NO_CHANGE',
      targetSheet: targetSheetName,
      targetRow: targetRow
    };
  }

  promiseCell
    .setValue(promiseDateTime)
    .setNumberFormat('dd/MM HH:mm');
  ghiLogHenGiaoCVDV_(ss, {
    row: row,
    receiveDate: receiveDate,
    plate: plate,
    advisor: advisorRaw,
    deliveryDate: deliveryDate,
    deliveryTime: deliveryTime,
    source: source,
    targetSheet: targetSheetName,
    targetRow: targetRow,
    matchCount: 1,
    result: 'UPDATED_FROM_WAITING',
    detail: 'Đã cập nhật HẸN GIAO CVDV từ snapshot Xe chờ giao.'
  });
  return {
    success: true,
    updated: true,
    code: 'UPDATED',
    targetSheet: targetSheetName,
    targetRow: targetRow,
    promise: promiseDateTime
  };
}

function mapAdvisorToCvdvSheet_(
  advisor
) {
  let key =
    normalizeKey_(
      advisor
    );
  // "CV Minh" -> "minh"
  key =
    key.replace(
      /^cv\s+/,
      ''
    );
  return (
    DELIVERY_PROMISE_SYNC_CFG
      .ADVISOR_SHEET_MAP[key] ||
    ''
  );
}
function findCvdvPromiseMatchesInOneSheet_(
  sheet,
  receiveDate,
  normalizedPlate
) {
  const firstRow =
    DELIVERY_PROMISE_SYNC_CFG
      .CVDV_FIRST_DATA_ROW;
  const lastRow =
    sheet.getLastRow();
  if (
    lastRow <
    firstRow
  ) {
    return [];
  }
  const values =
    sheet
      .getRange(
        firstRow,
        1,
        lastRow - firstRow + 1,
        3
      )
      .getValues();
  const matches = [];
  values.forEach(
    function(rowValues, index) {
      if (
        normalizePlate_(
          rowValues[
            DELIVERY_PROMISE_SYNC_CFG
              .CVDV_COL_PLATE - 1
          ]
        ) !==
        normalizedPlate
      ) {
        return;
      }
      if (
        !sameDate_(
          rowValues[
            DELIVERY_PROMISE_SYNC_CFG
              .CVDV_COL_DATE - 1
          ],
          receiveDate
        )
      ) {
        return;
      }
      matches.push(
        firstRow +
        index
      );
    }
  );
  return matches;
}
function combineDeliveryPromiseDateTime_(
  deliveryDate,
  minutes
) {
  const d =
    new Date(
      deliveryDate.getTime()
    );
  d.setHours(
    Math.floor(
      minutes / 60
    ),
    minutes % 60,
    0,
    0
  );
  return d;
}
function samePromiseDateTime_(
  a,
  b
) {
  if (
    !(a instanceof Date) ||
    isNaN(
      a.getTime()
    ) ||
    !(b instanceof Date) ||
    isNaN(
      b.getTime()
    )
  ) {
    return false;
  }
  return (
    a.getFullYear() ===
      b.getFullYear() &&
    a.getMonth() ===
      b.getMonth() &&
    a.getDate() ===
      b.getDate() &&
    a.getHours() ===
      b.getHours() &&
    a.getMinutes() ===
      b.getMinutes()
  );
}
function ghiLogHenGiaoCVDV_(
  ss,
  info
) {
  if (!ss) {
    return;
  }
  info =
    info || {};
  let sh =
    ss.getSheetByName(
      DELIVERY_PROMISE_SYNC_CFG
        .LOG_SHEET
    );
  if (!sh) {
    sh =
      ss.insertSheet(
        DELIVERY_PROMISE_SYNC_CFG
          .LOG_SHEET
      );
    sh
      .getRange(
        1,
        1,
        1,
        13
      )
      .setValues([[
        'Thời gian',
        'Nguồn',
        'Dòng Điều phối',
        'Ngày tiếp nhận',
        'Biển số',
        'CVDV',
        'Ngày giao',
        'Giờ giao',
        'Sheet CVDV',
        'Dòng CVDV',
        'Số match',
        'Kết quả',
        'Chi tiết'
      ]]);
    sh
      .getRange(
        1,
        1,
        1,
        13
      )
      .setFontWeight(
        'bold'
      );
    sh.setFrozenRows(1);
  }
  sh.appendRow([
    new Date(),
    clean_(info.source),
    Number(info.row || 0) || '',
    info.receiveDate || '',
    clean_(info.plate),
    clean_(info.advisor),
    info.deliveryDate || '',
    info.deliveryTime || '',
    clean_(info.targetSheet),
    Number(info.targetRow || 0) || '',
    typeof info.matchCount === 'number'
      ? info.matchCount
      : '',
    clean_(info.result),
    clean_(info.detail)
  ]);
  const last =
    sh.getLastRow();
  if (
    last >= 2
  ) {
    sh
      .getRange(
        last,
        1
      )
      .setNumberFormat(
        'dd/MM/yyyy HH:mm:ss'
      );
    sh
      .getRange(
        last,
        4
      )
      .setNumberFormat(
        'dd/MM/yyyy'
      );
    sh
      .getRange(
        last,
        7
      )
      .setNumberFormat(
        'dd/MM/yyyy'
      );
    sh
      .getRange(
        last,
        8
      )
      .setNumberFormat(
        'HH:mm'
      );
  }
}
/*********************************************************
 * V13.17.2 - CROSS-CHECK CVDV / LỄ TÂN + CLEAN VIEW AN TOÀN
 *********************************************************/

/**
 * P0.5.4: comparator ONLY. Never use this key as a persisted advisor name,
 * visitId, dispatch identity, grouping key for vehicle counts or a delivery
 * confirmation. Honda advisors are distinct people but one cross-source
 * advisory category. Toyota keys remain byte-for-byte compatible.
 */
function p054AdvisorMatchKey_(advisor) {
  const key = normalizeKey_(advisor).replace(/^cv\s+/, '');
  // Không tạo mảng tra cứu mới mỗi lần kiểm tra cố vấn.
  return key === 'lap' || key === 'bang' || key === 'tien' || key === 'honda'
    ? 'HONDA' : key;
}
function p054IsHondaAdvisor_(advisor) {
  return p054AdvisorMatchKey_(advisor) === 'HONDA';
}

/**
 * Manual, read-only Honda source review. Reception does NOT have a visitId.
 * Matching day + plate + HONDA is an ADVISOR comparison, never an automatic
 * physical-visit linkage. A prior-day reception does not create a new visit.
 * No Sheet.set*, notes, append, status change, journal writes or triggers.
 */
function p054HondaCrossCheckReadOnly_(sh, row, options) {
  // Giữ tham số options để tương thích lời gọi cũ; chưa dùng tại nhánh READ_ONLY.
  const mainSS = sh.getParent();
  const kind = sh.getName();
  const isDp = kind === CFG.SHEET_DISPATCH;
  const isWaiting = kind === CFG.SHEET_WAITING_DELIVERY;
  if (!isDp && !isWaiting) throw new Error('P054_SHEET_NOT_SUPPORTED');
  const width = isDp ? CFG.DP.VISIT_ID : CFG.WD.VISIT_ID;
  const raw = sh.getRange(row, 1, 1, width).getValues()[0];
  const shown = sh.getRange(row, 1, 1, width).getDisplayValues()[0];
  const plate = normalizePlate_(shown[1]);
  const advisor = clean_(shown[2]);
  const visitId = clean_(raw[width - 1]);
  const day = cvdvDateKey_(raw[0], mainSS);
  if (!p054IsHondaAdvisor_(advisor)) throw new Error('P054_NOT_HONDA');
  const result = {
    version:'P0.5.4', readOnly:true, sourceSheet:kind, sourceRow:row,
    plate:plate, originalAdvisor:advisor, advisorCompareKey:'HONDA',
    visitId:visitId, workDate:day, receptionRows:[], historicalRows:[],
    receptionResult:'NOT_CHECKED', cvdvResult:'NOT_CHECKED',
    cvdvRows:[], safeToLink:false, issuedVisitId:false,
    changedNames:false, changedCounts:false
  };
  if (!plate || !day || !visitId) {
    result.receptionResult='BLOCK_INCOMPLETE_SOURCE_IDENTITY';
    console.log('[P054_HONDA_READ_ONLY] '+JSON.stringify(result));
    return result;
  }
  try {
    const rSS=SpreadsheetApp.openById(CVDV_CROSSCHECK_CFG.RECEPTION_SPREADSHEET_ID);
    const reception=rSS.getSheetByName('Giờ ra vào 2026') ||
                    findReceptionSheetCross_(rSS);
    if (!reception) {
      result.receptionResult='RECEPTION_TAB_UNAVAILABLE';
    } else {
      const end=reception.getLastRow();
      const start=Math.max(2,end-6500+1);
      if (end>=start) {
        // B:H => B plate, E entry, F exit, G advisor, H notes.
        const rows=reception.getRange(start,2,end-start+1,7).getDisplayValues();
        rows.forEach(function(v,i) {
          if (normalizePlate_(v[0])!==plate) return;
          const entry=p049ReceptionStamp_(v[3]);
          if (!entry) return; // do not infer a missing date/time
          const sample={row:start+i,day:entry.day,
            advisor:clean_(v[5]),advisorKey:p054AdvisorMatchKey_(v[5]),
            noService:p056IsNoServiceType_(v[1]),
            hasGateOut:!!clean_(v[4])};
          if (entry.day===day) result.receptionRows.push(sample);
          else if (entry.day<day) result.historicalRows.push(sample);
        });
        result.historicalRows=result.historicalRows.slice(-12);
        if (result.receptionRows.some(function(x) { return x.noService; }))
          result.receptionResult='BLOCK_NO_SERVICE_RECEPTION';
        else if (result.receptionRows.length>1)
          result.receptionResult='BLOCK_MULTI_RECEPTION_SAME_DAY';
        else if (result.receptionRows.length===1)
          result.receptionResult=result.receptionRows[0].advisorKey==='HONDA'
            ? 'HONDA_ADVISOR_EQUIVALENT_ONLY' : 'BLOCK_RECEPTION_ADVISOR_CONFLICT';
        else if (result.historicalRows.length)
          result.receptionResult='PRIOR_RECEPTION_NEEDS_CONTINUOUS_STAY_PROOF';
        else result.receptionResult='NO_RECEPTION_IN_BOUNDED_WINDOW';
      } else result.receptionResult='RECEPTION_EMPTY';
    }
  } catch (err) {
    result.receptionResult='RECEPTION_READ_FAILED';
    result.receptionError=String(err.message||err).slice(0,240);
  }
  try {
    const cv=SpreadsheetApp.openById(DELIVERY_PROMISE_SYNC_CFG.CVDV_SPREADSHEET_ID);
    const honda=cv.getSheetByName('CV HONDA');
    if (!honda) result.cvdvResult='HONDA_SHEET_NOT_AVAILABLE';
    else {
      const end=honda.getLastRow();
      if (end>=3) {
        const rows=honda.getRange(3,1,end-2,5).getValues();
        const shown=honda.getRange(3,1,end-2,5).getDisplayValues();
        rows.forEach(function(v,i){
          if (normalizePlate_(shown[i][2])!==plate) return;
          const recordDay=cvdvDateKey_(v[0],cv);
          if (recordDay!==day) return;
          result.cvdvRows.push({row:i+3,day:recordDay,
                                status:clean_(shown[i][4])});
        });
      }
      result.cvdvResult=result.cvdvRows.length===0
        ? 'CVDV_NOT_YET_VERIFIED_OR_NOT_SYNCED'
        : result.cvdvRows.length===1
          ? 'CVDV_SINGLE_DATE_PLATE_RECORD' : 'BLOCK_CVDV_MULTIPLE_RECORDS';
    }
  } catch (err) {
    result.cvdvResult='CVDV_READ_FAILED';
    result.cvdvError=String(err.message||err).slice(0,240);
  }
  console.log('[P054_HONDA_READ_ONLY] '+JSON.stringify(result));
  return result;
}

/** P056 – no-service is a Reception classification, never a Dispatch work type.
 * Compare exact date+plate and fail closed if No DV, duplicate, source unavailable
 * or requested date is older than the bounded Reception window. READ ONLY.
 */
const P056_NO_SERVICE_CFG = Object.freeze({
  VERSION: 'P0.5.6_P056_NO_SERVICE_20261111',
  RECEPTION_TAB: 'Giờ ra vào 2026',
  LOOKBACK_ROWS: 3500
});
function p056IsNoServiceType_(type) {
  return normalizeKey_(type) === 'khong dv';
}
function p056ReceptionGateReadOnly_(receiveDate, plate, mainSS) {
  const identity = normalizePlate_(plate);
  if (!(receiveDate instanceof Date) || isNaN(receiveDate.getTime()) || !identity) {
    return { blocked: true, code: 'P056_IDENTITY_INCOMPLETE', rows: [] };
  }
  const day = cvdvDateKey_(receiveDate, mainSS);
  if (!day) return { blocked: true, code: 'P056_DAY_INVALID', rows: [] };
  try {
    const receptionSS = SpreadsheetApp.openById(CVDV_CROSSCHECK_CFG.RECEPTION_SPREADSHEET_ID);
    const sheet = findReceptionSheetCross_(receptionSS);
    if (!sheet || sheet.getName() !== P056_NO_SERVICE_CFG.RECEPTION_TAB) {
      return { blocked: true, code: 'P056_RECEPTION_TAB_UNAVAILABLE', rows: [] };
    }
    const end = sheet.getLastRow();
    const start = Math.max(2, end - P056_NO_SERVICE_CFG.LOOKBACK_ROWS + 1);
    if (end < start) return { blocked: true, code: 'P056_RECEPTION_EMPTY', rows: [] };
    // B:G: B plate, C type, D phone, E entry, F exit, G advisor.
    // Never write to this range; no fallback to inferred visits or completion.
    const shown = sheet.getRange(start, 2, end - start + 1, 6).getDisplayValues();
    let oldestDay = '';
    const matches = [];
    for (let i = 0; i < shown.length; i++) {
      const r = shown[i];
      const entry = p049ReceptionStamp_(r[3]);
      if (entry && (!oldestDay || entry.day < oldestDay)) oldestDay = entry.day;
      if (normalizePlate_(r[0]) !== identity) continue;
      if (!entry) {
        return { blocked: true, code: 'P056_PLATE_ENTRY_INVALID',
          day: day, plate: identity, rows: [start + i] };
      }
      if (entry.day === day) matches.push({row: start + i,
        plate: identity, day: day, type: String(r[1] || ''),
        noService: p056IsNoServiceType_(r[1])});
    }
    if (matches.some(function(x) { return x.noService; })) {
      return { blocked: true, code: 'P056_NO_SERVICE_RECEPTION_BLOCK',
        day: day, plate: identity, rows: matches.map(function(x) { return x.row; }) };
    }
    if (matches.length > 1) {
      return { blocked: true, code: 'P056_DUPLICATE_RECEPTION_VISIT',
        day: day, plate: identity, rows: matches.map(function(x) { return x.row; }) };
    }
    if (!oldestDay || day < oldestDay) {
      return { blocked: true, code: 'P056_DAY_OUTSIDE_RECEPTION_WINDOW',
        day: day, plate: identity, oldestDay: oldestDay, rows: [] };
    }
    return { blocked: false, code: matches.length ? 'P056_SERVICE_RECEPTION' : 'P056_NO_EXACT_RECEPTION',
      day: day, plate: identity, rows: matches.map(function(x) { return x.row; }),
      oldestDay: oldestDay };
  } catch (err) {
    return { blocked: true, code: 'P056_RECEPTION_READ_FAILED',
      day: day, plate: identity, error: String(err.message || err).slice(0,250), rows: [] };
  }
}
function p056NoServicePreflightReadOnly() {
  const ss = SpreadsheetApp.getActive();
  const receptionSS = SpreadsheetApp.openById(CVDV_CROSSCHECK_CFG.RECEPTION_SPREADSHEET_ID);
  const sheet = findReceptionSheetCross_(receptionSS);
  if (!sheet || sheet.getName() !== P056_NO_SERVICE_CFG.RECEPTION_TAB)
    throw new Error('P056_RECEPTION_TAB_UNAVAILABLE');
  const end = sheet.getLastRow();
  const start = Math.max(2, end - P056_NO_SERVICE_CFG.LOOKBACK_ROWS + 1);
  const rows = end >= start ? sheet.getRange(start,2,end-start+1,6).getDisplayValues() : [];
  // P056: read-only sentinels for the previously damaged Reception history.
  // These do not repair or modify any historical cell.
  const historyAnchors = [1000,5000,15000].map(function(row) {
    return {row:row,platePresent:end >= row &&
      !!normalizePlate_(sheet.getRange(row,2).getDisplayValue())};
  });
  const items = [];
  rows.forEach(function(r,i) {
    if (!p056IsNoServiceType_(r[1])) return;
    const e = p049ReceptionStamp_(r[3]);
    items.push({row:start+i, plate:normalizePlate_(r[0]), day:e?e.day:'', type:r[1]});
  });
  const dp = ss.getSheetByName(CFG.SHEET_DISPATCH);
  const wd = ss.getSheetByName(CFG.SHEET_WAITING_DELIVERY);
  const blockers = [];
  [[dp,CFG.DP.BIEN_SO,CFG.FIRST_DATA_ROW,'Điều phối'],
   [wd,CFG.WD.BIEN_SO,CFG.WAITING_DELIVERY_FIRST_ROW,'Xe chờ giao']].forEach(function(pair) {
    const sh=pair[0],col=pair[1],first=pair[2],name=pair[3];
    if (!sh) {blockers.push({sheet:name,code:'SHEET_MISSING'});return;}
    const last=sh.getLastRow();
    if (last<first)return;
    const sources=sh.getRange(first,1,last-first+1,2).getValues();
    const shown=sh.getRange(first,1,last-first+1,2).getDisplayValues();
    items.forEach(function(item){
      const matched=[];
      sources.forEach(function(v,i){
        if(normalizePlate_(shown[i][col-1])!==item.plate)return;
        if(cvdvDateKey_(v[0],ss)===item.day)matched.push(first+i);
      });
      if(matched.length)blockers.push({sheet:name,plate:item.plate,day:item.day,
        rows:matched,code:'SAME_DAY_PLATE_REQUIRES_REVIEW'});
    });
  });
  const out={version:P056_NO_SERVICE_CFG.VERSION,mode:'READ_ONLY',
    receptionRange:{start:start,end:end},nonServiceCount:items.length,items:items.slice(0,30),
    historyAnchors:historyAnchors,historyGuardPass:historyAnchors.every(function(x){return x.platePresent;}),
    dispatchPlateReview:blockers,sourceWrites:0,triggerChanges:0};
  console.log('[P056_NO_SERVICE_PREFLIGHT][READ_ONLY] '+JSON.stringify(out));
  return out;
}

function kiemTraCheoCvdvDongDangChon() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  const sheetName = sh ? sh.getName() : '';
  const isDispatch = sheetName === CFG.SHEET_DISPATCH;
  const isWaiting = sheetName === CFG.SHEET_WAITING_DELIVERY;
  if (!sh || (!isDispatch && !isWaiting)) {
    throw new Error('Hãy chọn một dòng trong sheet Điều phối hoặc Xe chờ giao.');
  }
  const row = sh.getActiveRange().getRow();
  const firstRow = isDispatch
    ? CFG.FIRST_DATA_ROW
    : CFG.WAITING_DELIVERY_FIRST_ROW;
  if (row < firstRow) {
    throw new Error('Hãy chọn một dòng dữ liệu trong ' + sheetName + '.');
  }

  // P0.5.4: Honda cross-check must never enter the legacy branch that
  // writes Reception/CVDV, promise cells or hides advisor-sheet rows.
  const selectedAdvisor=sh.getRange(row,3).getDisplayValue();
  if (p054IsHondaAdvisor_(selectedAdvisor)) {
    const honda=p054HondaCrossCheckReadOnly_(sh,row,{source:'MANUAL'});
    ss.toast('HONDA: '+honda.receptionResult+' | '+honda.cvdvResult,
             '🔎 HONDA – CHỈ KIỂM TRA',8);
    return {success:true,readOnly:true,crossCheck:honda,
            promise:null,hiddenBlankRows:0};
  }

  const result = crossCheckCvdvAdvisorForDispatchRow_(
    sh,
    row,
    {
      source: isDispatch ? 'MANUAL_DISPATCH' : 'MANUAL_WAITING',
      logNoRecord: true
    }
  );

  // P056: no partial side-effects after a no-service or unverified gate.
  if (result && result.blocked && /^P056_/.test(String(result.code || ''))) {
    ss.toast('Đã BLOCK: ' + result.code, '⚠️ LỄ TÂN KHÔNG DV', 8);
    return {success:false,readOnly:true,crossCheck:result,
      promise:null,hiddenBlankRows:0};
  }
  let promise = null;
  try {
    promise = isDispatch
      ? syncHenGiaoCVDVTuDieuPhoi_(
          sh,
          row,
          {
            source: 'CROSSCHECK_MANUAL_DISPATCH',
            manual: true
          }
        )
      : syncHenGiaoCVDVTuXeChoGiao_(
          sh,
          row,
          {
            source: 'CROSSCHECK_MANUAL_WAITING',
            manual: true
          }
        );
  }
  catch (err) {
    promise = {
      success: false,
      code: 'PROMISE_ERROR',
      message: err.message
    };
  }

  // V13.17.2: manual cross-check đồng thời dọn giao diện các dòng trống nội bộ
  // đã phát sinh từ các lần relocate trước đây. Chỉ HIDE, tuyệt đối không deleteRow.
  let hiddenBlankRows = 0;
  try {
    const cvdvSS = SpreadsheetApp.openById(
      DELIVERY_PROMISE_SYNC_CFG.CVDV_SPREADSHEET_ID
    );
    hiddenBlankRows = hideInternalBlankCvdvRowsCross_(cvdvSS);
  }
  catch (cleanupErr) {
    console.log('CVDV clean-view manual lỗi: ' + cleanupErr.message);
  }

  const parts = [];
  if (result.receptionUpdated) parts.push('Lễ tân đã sửa Cố vấn');
  if (result.cvdvMoved) {
    parts.push(
      'CVDV ' +
      result.cvdvFromSheet +
      ' → ' +
      result.cvdvTargetSheet
    );
  }
  if (!parts.length) {
    parts.push(
      result.blocked
        ? 'Đã BLOCK vì dữ liệu chưa an toàn'
        : 'Dữ liệu Cố vấn đã khớp / chưa có record cần chuyển'
    );
  }
  if (promise && promise.success && (promise.updated || promise.noChange)) {
    parts.push('Hẹn giao đã đồng bộ');
  }
  if (hiddenBlankRows > 0) {
    parts.push('Đã ẩn ' + hiddenBlankRows + ' dòng trống CVDV');
  }
  ss.toast(
    parts.join(' | '),
    result.blocked ? '⚠️ CROSS-CHECK CVDV' : '✅ CROSS-CHECK CVDV',
    9
  );
  return {
    success: !result.blocked,
    sourceSheet: sheetName,
    crossCheck: result,
    promise: promise,
    hiddenBlankRows: hiddenBlankRows
  };
}

function crossCheckCvdvAdvisorForDispatchRow_(dispatchSheet, row, options) {
  options = options || {};
  const sourceSheetName = dispatchSheet ? dispatchSheet.getName() : '';
  const isDispatch = sourceSheetName === CFG.SHEET_DISPATCH;
  const isWaiting = sourceSheetName === CFG.SHEET_WAITING_DELIVERY;
  const firstRow = isDispatch
    ? CFG.FIRST_DATA_ROW
    : CFG.WAITING_DELIVERY_FIRST_ROW;
  if (
    !dispatchSheet ||
    (!isDispatch && !isWaiting) ||
    row < firstRow
  ) {
    return {
      success: false,
      skipped: true,
      code: 'ROW_INVALID'
    };
  }
  const mainSS = dispatchSheet.getParent();
  const raw = dispatchSheet.getRange(row, 1, 1, 3).getValues()[0];
  const disp = dispatchSheet.getRange(row, 1, 1, 3).getDisplayValues()[0];
  const receiveDate = raw[CFG.DP.NGAY - 1];
  const plate = normalizePlate_(disp[CFG.DP.BIEN_SO - 1]);
  const mainAdvisor = clean_(disp[CFG.DP.CVDV - 1]);
  const targetSheetName = mapAdvisorToCvdvSheet_(mainAdvisor);
  const source = clean_(options.source || 'AUTO');
  const base = {
    source: source,
    sourceSheet: sourceSheetName,
    dispatchRow: row,
    receiveDate: receiveDate,
    plate: plate,
    mainAdvisor: mainAdvisor,
    cvdvTargetSheet: targetSheetName
  };
  if (!plate || !(receiveDate instanceof Date) || isNaN(receiveDate.getTime())) {
    return {
      success: false,
      skipped: true,
      code: 'IDENTITY_INCOMPLETE'
    };
  }
  // P056: direct and menu callers must never correct Reception advisor or
  // move CVDV when the EXACT arrival is classified as Không DV.
  const p056Gate = p056ReceptionGateReadOnly_(receiveDate, plate, mainSS);
  if (p056Gate.blocked) {
    console.warn('[P056_NO_SERVICE_CROSSCHECK_BLOCK] ' + JSON.stringify(p056Gate));
    return {success:false,blocked:true,code:p056Gate.code,
      receptionUpdated:false,cvdvMoved:false,sourceWrites:0,gate:p056Gate};
  }
  // P0.5.4: fail closed for direct callers as well as the menu wrapper.
  // Never route Honda through Toyota-only CVDV sheet relocation/update.
  if (p054IsHondaAdvisor_(mainAdvisor)) {
    return p054HondaCrossCheckReadOnly_(dispatchSheet,row,options);
  }
  if (!targetSheetName) {
    ghiLogCvdvCrossCheck_(
      mainSS,
      Object.assign({}, base, {
        result: 'ADVISOR_INVALID',
        detail: 'Không map được CVDV Điều phối "' + mainAdvisor + '".'
      })
    );
    return {
      success: false,
      blocked: true,
      code: 'ADVISOR_INVALID'
    };
  }

  const receptionSS = SpreadsheetApp.openById(
    CVDV_CROSSCHECK_CFG.RECEPTION_SPREADSHEET_ID
  );
  const receptionSheet = findReceptionSheetCross_(receptionSS);
  const cvdvSS = SpreadsheetApp.openById(
    DELIVERY_PROMISE_SYNC_CFG.CVDV_SPREADSHEET_ID
  );

  let receptionUpdated = false;
  let receptionBefore = '';
  let receptionRow = 0;
  let receptionAction = 'NO_MATCH';
  let blocked = false;
  const notes = [];

  if (receptionSheet) {
    const receptionMatches = findReceptionMatchesRecentCross_(
      receptionSheet,
      receiveDate,
      plate,
      mainSS
    );
    if (receptionMatches.length === 1) {
      const m = receptionMatches[0];
      receptionRow = m.row;
      receptionBefore = m.advisor;
      const receptionSheetName = mapAdvisorToCvdvSheet_(m.advisor);
      if (receptionSheetName !== targetSheetName) {
        // Re-check EXACT key ngay trước khi ghi.
        const verify = receptionSheet
          .getRange(m.row, CVDV_CROSSCHECK_CFG.RECEPTION_COL_PLATE, 1, 6)
          .getValues()[0];
        const verifyPlate = normalizePlate_(verify[0]);
        const verifyEntry = verify[
          CVDV_CROSSCHECK_CFG.RECEPTION_COL_ENTRY -
          CVDV_CROSSCHECK_CFG.RECEPTION_COL_PLATE
        ];
        if (
          verifyPlate === plate &&
          sameDate_(verifyEntry, receiveDate)
        ) {
          receptionSheet
            .getRange(m.row, CVDV_CROSSCHECK_CFG.RECEPTION_COL_ADVISOR)
            .setValue(targetSheetName.toUpperCase());
          receptionUpdated = true;
          receptionAction = 'UPDATED';
          notes.push(
            'Lễ tân ' +
            (m.advisor || '(trống)') +
            ' → ' +
            targetSheetName.toUpperCase()
          );
        }
        else {
          receptionAction = 'VERIFY_CHANGED';
          blocked = true;
          notes.push('Lễ tân thay đổi trước lúc sửa; không ghi.');
        }
      }
      else {
        receptionAction = 'ALREADY_CORRECT';
      }
    }
    else if (receptionMatches.length > 1) {
      receptionAction = 'MULTI_MATCH';
      blocked = true;
      notes.push(
        'Lễ tân có ' +
        receptionMatches.length +
        ' dòng cùng Ngày + BSX; không tự sửa.'
      );
    }
    else if (options.logNoRecord) {
      notes.push('Chưa tìm thấy record Lễ tân trong vùng dữ liệu gần nhất.');
    }
  }
  else {
    receptionAction = 'SHEET_MISSING';
    blocked = true;
    notes.push('Không tìm thấy sheet Giờ ra vào.');
  }

  let cvdvMoved = false;
  let cvdvFromSheet = '';
  let cvdvFromRow = 0;
  let cvdvTargetRow = 0;
  let cvdvAction = 'NO_MATCH';

  const targetSheet = cvdvSS.getSheetByName(targetSheetName);
  if (!targetSheet) {
    cvdvAction = 'TARGET_SHEET_MISSING';
    blocked = true;
    notes.push('Không tìm thấy sheet CVDV đích ' + targetSheetName + '.');
  }
  else {
    const targetMatches = findCvdvPromiseMatchesInOneSheet_(
      targetSheet,
      receiveDate,
      plate
    );
    if (targetMatches.length === 1) {
      cvdvAction = 'ALREADY_CORRECT';
      cvdvTargetRow = targetMatches[0];
    }
    else if (targetMatches.length > 1) {
      cvdvAction = 'MULTI_TARGET';
      blocked = true;
      notes.push(
        'Sheet đúng ' + targetSheetName +
        ' có ' + targetMatches.length +
        ' record cùng Ngày + BSX; không tự xử lý.'
      );
    }
    else {
      const globalMatches = [];
      getCvdvAdvisorSheetNamesCross_().forEach(function(sheetName) {
        if (sheetName === targetSheetName) return;
        const sh = cvdvSS.getSheetByName(sheetName);
        if (!sh) return;
        const rows = findCvdvPromiseMatchesInOneSheet_(sh, receiveDate, plate);
        rows.forEach(function(r) {
          globalMatches.push({
            sheetName: sheetName,
            row: r
          });
        });
      });
      if (globalMatches.length === 1) {
        cvdvFromSheet = globalMatches[0].sheetName;
        cvdvFromRow = globalMatches[0].row;
        const moved = relocateCvdvRecordCross_(
          cvdvSS,
          globalMatches[0],
          targetSheetName,
          receiveDate,
          plate
        );
        if (moved.success) {
          cvdvMoved = true;
          cvdvAction = 'MOVED';
          cvdvTargetRow = moved.targetRow;
          notes.push(
            'CVDV ' +
            cvdvFromSheet +
            '!dòng ' +
            cvdvFromRow +
            ' → ' +
            targetSheetName +
            '!dòng ' +
            cvdvTargetRow
          );
        }
        else {
          cvdvAction = moved.code || 'MOVE_BLOCKED';
          blocked = true;
          notes.push(moved.message || 'Không thể chuyển record CVDV.');
        }
      }
      else if (globalMatches.length > 1) {
        cvdvAction = 'MULTI_MATCH';
        blocked = true;
        notes.push(
          'CVDV có ' +
          globalMatches.length +
          ' record cùng Ngày + BSX ở nhiều sheet; không tự chuyển.'
        );
      }
      else if (options.logNoRecord) {
        notes.push('Chưa có record CVDV cùng Ngày + BSX để chuyển.');
      }
    }
  }

  const changed = receptionUpdated || cvdvMoved;
  if (
    changed ||
    blocked ||
    options.logNoRecord
  ) {
    ghiLogCvdvCrossCheck_(
      mainSS,
      Object.assign({}, base, {
        receptionAdvisorBefore: receptionBefore,
        receptionRow: receptionRow,
        cvdvFromSheet: cvdvFromSheet,
        cvdvFromRow: cvdvFromRow,
        cvdvTargetRow: cvdvTargetRow,
        receptionAction: receptionAction,
        cvdvAction: cvdvAction,
        result:
          blocked
            ? (changed ? 'PARTIAL_FIXED_BLOCKED' : 'BLOCKED')
            : (changed ? 'AUTO_FIXED' : 'NO_CHANGE'),
        detail: notes.join(' | ')
      })
    );
  }

  return {
    success: !blocked,
    blocked: blocked,
    receptionUpdated: receptionUpdated,
    receptionAction: receptionAction,
    receptionRow: receptionRow,
    cvdvMoved: cvdvMoved,
    cvdvAction: cvdvAction,
    cvdvFromSheet: cvdvFromSheet,
    cvdvFromRow: cvdvFromRow,
    cvdvTargetSheet: targetSheetName,
    cvdvTargetRow: cvdvTargetRow,
    plate: plate,
    mainAdvisor: mainAdvisor,
    detail: notes.join(' | ')
  };
}



function findReceptionSheetCross_(ss) {
  // P056: prefer exact Production tab. Legacy setting is not the 2026 tab.
  let sh = ss.getSheetByName(P056_NO_SERVICE_CFG.RECEPTION_TAB) ||
           ss.getSheetByName(CVDV_CROSSCHECK_CFG.RECEPTION_SHEET_NAME);
  if (sh) return sh;
  const wanted = 'gio ra vao';
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (normalizeKey_(sheets[i].getName()) === wanted) return sheets[i];
  }
  return null;
}

function findReceptionMatchesRecentCross_(sheet, receiveDate, plate, mainSS) {
  const last = sheet.getLastRow();
  if (last < CVDV_CROSSCHECK_CFG.RECEPTION_FIRST_DATA_ROW) return [];
  const start = Math.max(
    CVDV_CROSSCHECK_CFG.RECEPTION_FIRST_DATA_ROW,
    last - CVDV_CROSSCHECK_CFG.RECEPTION_LOOKBACK_ROWS + 1
  );
  const values = sheet
    .getRange(start, CVDV_CROSSCHECK_CFG.RECEPTION_COL_PLATE, last - start + 1, 6)
    .getValues();
  const out = [];
  for (let i = 0; i < values.length; i++) {
    const rowPlate = normalizePlate_(values[i][0]);
    const entry = values[i][
      CVDV_CROSSCHECK_CFG.RECEPTION_COL_ENTRY -
      CVDV_CROSSCHECK_CFG.RECEPTION_COL_PLATE
    ];
    if (rowPlate !== plate || !sameDate_(entry, receiveDate)) continue;
    out.push({
      row: start + i,
      noService: p056IsNoServiceType_(values[i][1]),
      advisor: clean_(values[i][
        CVDV_CROSSCHECK_CFG.RECEPTION_COL_ADVISOR -
        CVDV_CROSSCHECK_CFG.RECEPTION_COL_PLATE
      ])
    });
  }
  return out;
}



function getCvdvAdvisorSheetNamesCross_() {
  return Array.from(
    new Set(
      Object.keys(DELIVERY_PROMISE_SYNC_CFG.ADVISOR_SHEET_MAP)
        .map(function(key) {
          return DELIVERY_PROMISE_SYNC_CFG.ADVISOR_SHEET_MAP[key];
        })
    )
  );
}



function relocateCvdvRecordCross_(cvdvSS, match, targetSheetName, receiveDate, plate) {
  const sourceSheet = cvdvSS.getSheetByName(match.sheetName);
  const targetSheet = cvdvSS.getSheetByName(targetSheetName);
  if (!sourceSheet || !targetSheet) {
    return {
      success: false,
      code: 'SHEET_MISSING',
      message: 'Thiếu sheet nguồn hoặc sheet đích CVDV.'
    };
  }
  if (sourceSheet.getName() === targetSheet.getName()) {
    return {
      success: true,
      noChange: true,
      targetRow: match.row
    };
  }

  const verifySource = sourceSheet.getRange(match.row, 1, 1, 3).getValues()[0];
  if (
    !sameDate_(verifySource[0], receiveDate) ||
    normalizePlate_(verifySource[2]) !== plate
  ) {
    return {
      success: false,
      code: 'SOURCE_CHANGED',
      message: 'Record CVDV nguồn thay đổi trước lúc chuyển.'
    };
  }

  const targetMatches = findCvdvPromiseMatchesInOneSheet_(
    targetSheet,
    receiveDate,
    plate
  );
  if (targetMatches.length > 0) {
    return {
      success: false,
      code: 'TARGET_ALREADY_EXISTS',
      message: 'Đích đã xuất hiện record cùng Ngày + BSX; dừng để tránh duplicate.'
    };
  }

  const sourceRange = sourceSheet.getRange(
    match.row,
    1,
    1,
    mainQtWidth_(sourceSheet)
  );
  const sourceValues = sourceRange.getValues()[0];
  const sourceNotes = sourceRange.getNotes()[0];
  const qt=mainQtNew_(sourceSheet);
  if(qt!==mainQtNew_(targetSheet))return {success:false,code:'QT_MIXED_SCHEMA',
    message:'Không chuyển CVDV giữa 2 sheet có cấu trúc khác nhau.'};
  const width=qt?11:10;
  if(qt&&sourceValues[5]&&['Chưa QT','Đã QT'].indexOf(String(sourceValues[5]))<0)
    return {success:false,code:'QT_INVALID',message:'Giá trị QT nguồn không hợp lệ.'};

  // V13.18.1: chèn đúng block ngày thay vì append cuối sheet.
  const slot = findDateBlockSlotCvdvCross_(
    targetSheet,
    receiveDate,
    cvdvSS
  );
  const targetRow = slot.row;

  try {
    ensureCvdvRowExistsCross_(targetSheet, targetRow);
    const targetDisplay = targetSheet
      .getRange(targetRow, 1, 1, width)
      .getDisplayValues()[0];
    if (!slot.inserted && !isReusableCvdvRowCross_(targetDisplay)) {
      throw new Error('Dòng đích không còn trống an toàn.');
    }

    applyCvdvFormatLockCross_(
      targetSheet,
      targetRow,
      slot.templateRow
    );

    // Chỉ chuyển dữ liệu nghiệp vụ; không kéo format nguồn sang.
    if(!slot.inserted){
      var dstBefore=targetSheet.getRange(targetRow,1,1,width);
      var dstVals=dstBefore.getValues()[0],dstNotes=dstBefore.getNotes()[0];
      if(!mainQtBusiness_(dstVals)||
         (qt?[0,1,2,3,4,5,7,9,10]:[0,1,2,3,4,6,8,9]).some(function(i){
            return !!String(dstNotes[i]||'').trim();}))
        throw Error('[MAIN_QT] Target slot occupied with data/note');
    }
    mainQtEnsureDropdown_(targetSheet,targetRow);
    targetSheet.getRange(targetRow, 1, 1, 5).setValues([[ 
      sourceValues[0], sourceValues[1], sourceValues[2], sourceValues[3], sourceValues[4]
    ]]);

    // F/H legacy are calculated, never paste F array or H calculated values.
    // In QT layout, F is the manual QT field; formulas move to G/I.
    if(qt && sourceValues[5]!==''&&sourceValues[5]!=null)
      targetSheet.getRange(targetRow,6).setValue(sourceValues[5]);
    targetSheet.getRange(targetRow,qt?8:7).setValue(sourceValues[qt?7:6]);
    targetSheet.getRange(targetRow,qt?10:9,1,2).setValues([[
      sourceValues[qt?9:8],sourceValues[qt?10:9]
    ]]);
    targetSheet.getRange(targetRow,1,1,width).setNotes([sourceNotes]);
    mainQtEnsureDropdown_(targetSheet,targetRow);

    // FORMAT LOCK lần cuối sau khi ghi.
    applyCvdvFormatLockCross_(targetSheet, targetRow, slot.templateRow);
    try { targetSheet.showRows(targetRow); } catch (err) {}
    SpreadsheetApp.flush();
  } catch (writeErr) {
    // Avoid deleting rows written by another script or partial accepted data.
    console.warn('[MAIN_QT] target transfer error, preserved for audit: '+String(writeErr));
    return {
      success: false,
      code: 'TARGET_WRITE_FAILED',
      message: 'Không ghi được CVDV đích: ' + writeErr.message
    };
  }

  const verifyTarget = targetSheet.getRange(targetRow, 1, 1, 3).getValues()[0];
  if (
    !sameDate_(verifyTarget[0], receiveDate) ||
    normalizePlate_(verifyTarget[2]) !== plate
  ) {
    // Không xóa nguồn nếu đích chưa xác minh được.
    console.warn('[MAIN_QT] Target verification failed; keep both records for audit');
    return {
      success: false,
      code: 'TARGET_VERIFY_FAILED',
      message: 'Copy sang sheet đích không xác minh được; nguồn được giữ nguyên.'
    };
  }

  // Verify *all* transferred business fields and unchanged source BEFORE clear.
  if(!mainQtValidateRow_(sourceSheet,targetSheet,match.row,targetRow,sourceValues))
    return {success:false,code:'QT_BUSINESS_VERIFY_FAILED',
      message:'Chưa xác minh toàn bộ A:E, QT, Hẹn giao và ghi chú; giữ nguyên nguồn.'};
  // Only clear business fields; G/I are formulas in 11-column schema.
  sourceSheet.getRange(match.row,1,1,5).clearContent();
  if(qt)sourceSheet.getRange(match.row,6).clearContent();
  sourceSheet.getRange(match.row,qt?8:7).clearContent();
  sourceSheet.getRange(match.row,qt?10:9,1,2).clearContent();
  sourceSheet.getRange(match.row,1,1,width).clearNote();

  // STT là thứ tự trong ngày theo từng CVDV, nên re-index 2 sheet sau khi chuyển.
  const dateKey = cvdvDateKey_(receiveDate, cvdvSS);
  reindexCvdvSttDateCross_(sourceSheet, dateKey, cvdvSS);
  reindexCvdvSttDateCross_(targetSheet, dateKey, cvdvSS);
  SpreadsheetApp.flush();

  const sourceRowHidden = hideClearedCvdvSourceRowCross_(
    sourceSheet,
    match.row
  );
  try { targetSheet.showRows(targetRow); } catch (err) {}

  return {
    success: true,
    targetRow: targetRow,
    sourceSheet: match.sheetName,
    sourceRow: match.row,
    targetSheet: targetSheetName,
    sourceRowHidden: sourceRowHidden,
    insertedInDateBlock: !!slot.inserted,
    formatLocked: true
  };
}

function findDateBlockSlotCvdvCross_(sheet, receiveDate, cvdvSS) {
  const firstRow = CVDV_CROSSCHECK_CFG.CVDV_FIRST_DATA_ROW;
  const targetKey = cvdvDateKey_(receiveDate, cvdvSS);
  const lastRow = sheet.getLastRow();

  let lastSame = 0;
  let firstLater = 0;
  let lastReal = 0;

  if (lastRow >= firstRow) {
    const values = sheet.getRange(firstRow, 1, lastRow - firstRow + 1, 3).getValues();
    for (let i = 0; i < values.length; i++) {
      const rowNum = firstRow + i;
      const plate = normalizePlate_(values[i][2]);
      if (!plate) continue;
      const key = cvdvDateKey_(values[i][0], cvdvSS);
      if (!key) continue;
      lastReal = rowNum;
      if (key === targetKey) lastSame = rowNum;
      if (!firstLater && key > targetKey) firstLater = rowNum;
    }
  }

  if (lastSame) {
    const candidate = lastSame + 1;
    if (candidate <= sheet.getMaxRows()) {
      const display = sheet
        .getRange(candidate, 1, 1, mainQtWidth_(sheet))
        .getDisplayValues()[0];
      if (isReusableCvdvRowCross_(display)) {
        return { row: candidate, inserted: false, templateRow: lastSame };
      }
    }
    sheet.insertRowsAfter(lastSame, 1);
    return { row: lastSame + 1, inserted: true, templateRow: lastSame };
  }

  if (firstLater) {
    sheet.insertRowsBefore(firstLater, 1);
    return {
      row: firstLater,
      inserted: true,
      templateRow: firstLater + 1
    };
  }

  const row = findReusableCvdvRowCross_(sheet);
  ensureCvdvRowExistsCross_(sheet, row);
  return {
    row: row,
    inserted: false,
    templateRow: lastReal || Math.max(firstRow, row - 1)
  };
}

/**
 * Format lock mềm: người dùng vẫn sửa nghiệp vụ bình thường,
 * nhưng mọi record do correction tạo/chuyển luôn theo style sheet đích.
 */
function applyCvdvFormatLockCross_(sheet, row, templateRow) {
  const width=mainQtWidth_(sheet);
  const target = sheet.getRange(row, 1, 1, width);

  if (
    templateRow &&
    templateRow >= CVDV_CROSSCHECK_CFG.CVDV_FIRST_DATA_ROW &&
    templateRow <= sheet.getMaxRows() &&
    templateRow !== row
  ) {
    const template = sheet.getRange(templateRow, 1, 1, width);
    try {
      template.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    } catch (err) {}
    try {
      template.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    } catch (err) {}
    try {
      sheet.setRowHeight(row, sheet.getRowHeight(templateRow));
    } catch (err) {}
  }

  target.setFontFamily('Arial');
  sheet.getRange(row, 1).setNumberFormat('dd/MM/yy');
  sheet.getRange(row, 3).setNumberFormat('@');
  sheet.getRange(row,mainQtCol_(sheet,7)).setNumberFormat('dd/MM HH:mm');
  sheet.getRange(row,mainQtCol_(sheet,8)).setNumberFormat('[h]:mm:ss');
}



function hideClearedCvdvSourceRowCross_(sheet, row) {
  if (!sheet || row < CVDV_CROSSCHECK_CFG.CVDV_FIRST_DATA_ROW) return false;
  const display = sheet
    .getRange(row, 1, 1, mainQtWidth_(sheet))
    .getDisplayValues()[0];
  if (!isReusableCvdvRowCross_(display)) return false;
  try {
    sheet.hideRows(row);
    return true;
  }
  catch (err) {
    console.log(
      'Không ẩn được dòng nguồn CVDV ' + sheet.getName() + '!' + row + ': ' + err.message
    );
    return false;
  }
}

function hideInternalBlankCvdvRowsCross_(cvdvSS) {
  if (!cvdvSS) return 0;
  let hidden = 0;
  getCvdvAdvisorSheetNamesCross_().forEach(function(sheetName) {
    const sh = cvdvSS.getSheetByName(sheetName);
    if (!sh) return;
    const first = CVDV_CROSSCHECK_CFG.CVDV_FIRST_DATA_ROW;
    const last = sh.getLastRow();
    if (last < first) return;

    const values = sh
      .getRange(first, 1, last - first + 1, mainQtWidth_(sh))
      .getDisplayValues();

    // Chỉ xét tới record có BSX cuối cùng. Các dòng nhập trống ở cuối sheet vẫn để hiện.
    let lastBusinessOffset = -1;
    for (let i = 0; i < values.length; i++) {
      if (clean_(values[i][2])) lastBusinessOffset = i;
    }
    if (lastBusinessOffset < 0) return;

    let runStart = -1;
    let runCount = 0;
    const flushRun = function() {
      if (runStart < 0 || runCount <= 0) return;
      try {
        sh.hideRows(runStart, runCount);
        hidden += runCount;
      }
      catch (err) {
        console.log(
          'Không ẩn được dòng trống nội bộ ' + sh.getName() + '!' +
          runStart + ':' + (runStart + runCount - 1) + ': ' + err.message
        );
      }
      runStart = -1;
      runCount = 0;
    };

    for (let i = 0; i <= lastBusinessOffset; i++) {
      const rowNumber = first + i;
      const isBlankBusinessRow = isReusableCvdvRowCross_(values[i]);
      if (isBlankBusinessRow) {
        if (runStart < 0) runStart = rowNumber;
        runCount++;
      }
      else {
        flushRun();
      }
    }
    flushRun();
  });
  return hidden;
}

function findReusableCvdvRowCross_(sheet) {
  const first = CVDV_CROSSCHECK_CFG.CVDV_FIRST_DATA_ROW;
  const last = Math.max(sheet.getLastRow(), first);
  const values = sheet
    .getRange(first, 1, last - first + 1, mainQtWidth_(sheet))
    .getDisplayValues();
  let lastPlateOffset = -1;
  for (let i = 0; i < values.length; i++) {
    if (clean_(values[i][2])) lastPlateOffset = i;
  }
  const startOffset = Math.max(0, lastPlateOffset + 1);
  for (let i = startOffset; i < values.length; i++) {
    if (isReusableCvdvRowCross_(values[i])) return first + i;
  }
  return last + 1;
}

function isReusableCvdvRowCross_(row) {
  if (!row || clean_(row[2])) return false;
  // Ignore only formula columns: F/H old, G/I new.
  return mainQtBusiness_(row);
}

function ensureCvdvRowExistsCross_(sheet, row) {
  if (row <= sheet.getMaxRows()) return;
  sheet.insertRowsAfter(
    sheet.getMaxRows(),
    row - sheet.getMaxRows()
  );
}

function reindexCvdvSttDateCross_(sheet, dateKey, mainSS) {
  if (!dateKey) return;
  const first = CVDV_CROSSCHECK_CFG.CVDV_FIRST_DATA_ROW;
  const last = sheet.getLastRow();
  if (last < first) return;
  const values = sheet.getRange(first, 1, last - first + 1, 3).getValues();
  let seq = 0;
  for (let i = 0; i < values.length; i++) {
    const rowDateKey = cvdvDateKey_(values[i][0], mainSS);
    const plate = normalizePlate_(values[i][2]);
    if (!plate || rowDateKey !== dateKey) continue;
    seq++;
    if (Number(values[i][1]) !== seq) {
      sheet.getRange(first + i, 2).setValue(seq);
    }
  }
}

function ghiLogCvdvCrossCheck_(ss, info) {
  if (!ss) return;
  info = info || {};
  let sh = ss.getSheetByName(CVDV_CROSSCHECK_CFG.LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CVDV_CROSSCHECK_CFG.LOG_SHEET);
    sh.getRange(1, 1, 1, 15).setValues([[
      'Thời gian',
      'Nguồn',
      'Dòng Điều phối',
      'Ngày tiếp nhận',
      'Biển số',
      'CVDV Main',
      'CVDV Lễ tân trước',
      'Dòng Lễ tân',
      'Sheet CVDV trước',
      'Dòng CVDV trước',
      'Sheet CVDV đúng',
      'Dòng CVDV mới',
      'Xử lý Lễ tân',
      'Xử lý CVDV',
      'Kết quả / Chi tiết'
    ]]);
    sh.getRange(1, 1, 1, 15).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  sh.appendRow([
    new Date(),
    clean_(info.source),
    Number(info.dispatchRow || 0) || '',
    info.receiveDate || '',
    clean_(info.plate),
    clean_(info.mainAdvisor),
    clean_(info.receptionAdvisorBefore),
    Number(info.receptionRow || 0) || '',
    clean_(info.cvdvFromSheet),
    Number(info.cvdvFromRow || 0) || '',
    clean_(info.cvdvTargetSheet),
    Number(info.cvdvTargetRow || 0) || '',
    clean_(info.receptionAction),
    clean_(info.cvdvAction),
    clean_(info.result) + (info.detail ? ' | ' + clean_(info.detail) : '')
  ]);
  const last = sh.getLastRow();
  sh.getRange(last, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  sh.getRange(last, 4).setNumberFormat('dd/MM/yyyy');
}

/*********************************************************
 * GIỜ NHẬN / GIỜ GIAO
 *
 * 9    -> 09:00
 * 14   -> 14:00
 * 930  -> 09:30
 * 1430 -> 14:30
 *********************************************************/
function normalizeEditedTimeCells_(sh, row, startCol, endCol, e) {
  [CFG.DP.GIO_NHAN, CFG.DP.GIO_GIAO].forEach(col => {
    if (startCol > col || endCol < col) return;
    const cell = sh.getRange(row, col);
    const single = e && e.range && e.range.getNumRows() === 1 && e.range.getNumColumns() === 1 && e.range.getRow() === row && e.range.getColumn() === col;
    const raw = single && typeof e.value !== 'undefined' ? String(e.value).trim() : cell.getValue();
    if (raw === '') return;
    let parsed;
    if (raw instanceof Date) parsed = isValidTimeValue_(raw) ? raw : null;
    else if (typeof raw === 'number' && raw >= 0 && raw < 1) parsed = new Date(1899, 11, 30, 0, Math.round(raw * 1440));
    else {
      const text = String(raw).trim();
      const h = text.match(/^(\d{1,2})$/);
      parsed = h && validTime_(+h[1], 0) ? new Date(1899, 11, 30, +h[1], 0) : parseGioWeb_(text.replace(/h/ig, ':'));
    }
    if (parsed) setTimeCell_(cell, parsed.getHours(), parsed.getMinutes());
    else sh.getParent().toast('Giờ không hợp lệ: ' + raw + '. Phải sửa trước khi Hoàn thành.', 'GIỜ KHÔNG HỢP LỆ', 8);
  });
}
function validTime_(
  hh,
  mm
) {
  return (
    Number.isFinite(hh) &&
    Number.isFinite(mm) &&
    hh >= 0 &&
    hh <= 23 &&
    mm >= 0 &&
    mm <= 59
  );
}
function setTimeCell_(
  cell,
  hh,
  mm
) {
  const serial =
    (
      hh * 60 +
      mm
    ) /
    1440;
  cell
    .setValue(
      serial
    )
    .setNumberFormat(
      'HH:mm'
    );
}
/*********************************************************
 * STATUS STAMP
 *********************************************************/
/** Change detection uses last committed status W/X, not just edit-column. */
function p043StatusReallyChanged_(sh, row, statusCol, lastCol, edited, retry) {
  if (!edited || retry) return false;
  const current = normalizeKey_(sh.getRange(row, statusCol).getDisplayValue());
  const committed = normalizeKey_(sh.getRange(row, lastCol).getDisplayValue());
  return current !== committed;
}
function updateStatusStamp_(
  sh,
  row,
  statusCol,
  stampCol,
  edited,
  observedEventAt
) {
  const status =
    clean_(
      sh
        .getRange(
          row,
          statusCol
        )
        .getDisplayValue()
    );
  const stampCell =
    sh.getRange(
      row,
      stampCol
    );
  if (
    edited
  ) {
    if (
      status
    ) {
      // P0.4.4: the journal and direct commit MUST share one time.
      // Otherwise worker may reject the older journal with
      // P043_QUEUED_EVENT_ORDER_CONFLICT after partial direct writes.
      const stamp = observedEventAt ? new Date(observedEventAt) : new Date();
      if (isNaN(stamp.getTime()) || stamp.getTime() > Date.now() + 60000) {
        throw new Error('P044_DIRECT_EVENT_TIME_INVALID: ' + statusCol);
      }
      stampCell
        .setValue(stamp)
        .setNumberFormat('dd/MM/yyyy HH:mm:ss');
    } else {
      stampCell
        .clearContent();
    }
  } else if (
    status &&
    !isComplete_(status) &&
    !stampCell.getValue()
  ) {
    stampCell
      .setValue(
        new Date()
      )
      .setNumberFormat(
        'dd/MM/yyyy HH:mm:ss'
      );
  }
}
/*********************************************************
 * RESTORE STATUS
 *********************************************************/
function restoreStatusAfterValidationFail_(sh, row, col, e) {
  if (!e || !e.range || !editTouchesColumn_(e.range, col)) return;
  const previousCol = col === CFG.DP.TD_EM ? CFG.DP.LAST_EM_STATUS : CFG.DP.LAST_SCC_STATUS;
  let previous = clean_(sh.getRange(row, previousCol).getValue());
  const single = e.range.getNumRows() === 1 && e.range.getNumColumns() === 1 && e.range.getColumn() === col;
  if (!previous && single && e.oldValue !== undefined) previous = e.oldValue;
  sh.getRange(row, col).setValue(previous);
}
/*********************************************************
 * WARNING
 *********************************************************/
function showCompletionWarning_(plate, side, missing, ss) {
  (ss || laySpreadsheetWeb_()).toast('Xe ' + plate + ' / ' + side + ': ' + missing.join(', '), 'Chưa thể Hoàn thành', 10);
}
/*********************************************************
 * REMOVE DATA OPEN NẾU ĐỔI LUỒNG
 *********************************************************/
function removeOpenDataRowIfStale_(
  sheetName,
  rec
) {
  const sh =
    spreadsheetForRecord_(rec)
      .getSheetByName(
        sheetName
      );
  if (!sh) return;
  const row =
    findDataRow_(
      sh,
      rec
    );
  if (!row) return;
  const isEM =
    sheetName ===
    CFG.SHEET_DATA_EM;
  const completedCol =
    isEM
      ? 12
      : 13;
  const statusCol =
    isEM
      ? 13
      : 14;
  const completed =
    sh
      .getRange(
        row,
        completedCol
      )
      .getValue();
  const status =
    sh
      .getRange(
        row,
        statusCol
      )
      .getDisplayValue();
  if (
    !completed &&
    !isComplete_(status)
  ) {
    // P0.4.2: không xóa OPEN trong một lượt có xác nhận hoàn thành
    // ở bất kỳ nhánh nào; tránh Data biến mất khi đổi phân công tạm thời.
    if (isComplete_(rec.emStatus) || isComplete_(rec.sccStatus)) {
      console.log('[P042_STALE_WORK_PRESERVED] ' + sheetName + ' / ' + rec.visitId);
      return;
    }
    sh.deleteRow(row);
  }
}
/*********************************************************
 * READY TO CLEAR
 *********************************************************/
function isDispatchReadyToClear_(
  rec,
  route
) {
  const emDone =
    !route.em.active ||
    isComplete_(
      rec.emStatus
    );
  const sccDone =
    !route.scc.active ||
    isComplete_(
      rec.sccStatus
    );
  return (
    (
      route.em.active ||
      route.scc.active
    ) &&
    emDone &&
    sccDone
  );
}
/*********************************************************
 * CLEAR + COMPACT
 *********************************************************/
function dispatchRowHasAnyValueForCompact_(displayRow) {
  if (!displayRow) return false;
  for (let i = 0; i < displayRow.length; i++) {
    if (clean_(displayRow[i])) return true;
  }
  return false;
}

function inspectDispatchForSafeCompact_(sh) {
  const start = CFG.FIRST_DATA_ROW;
  const last = sh.getLastRow();
  const width = CFG.DP.LAST_SCC_STATUS;
  if (last < start) {
    return {
      ok: true,
      start: start,
      last: last,
      display: [],
      activeVisitIds: [],
      activePlates: []
    };
  }

  const display = sh
    .getRange(start, 1, last - start + 1, width)
    .getDisplayValues();

  const seenIds = Object.create(null);
  const seenPlates = Object.create(null);
  const activeVisitIds = [];
  const activePlates = [];

  for (let i = 0; i < display.length; i++) {
    const r = display[i];
    if (!dispatchRowHasAnyValueForCompact_(r)) continue;

    const row = start + i;
    const plate = normalizePlate_(r[CFG.DP.BIEN_SO - 1]);
    const visitId = clean_(r[CFG.DP.VISIT_ID - 1]);

    // Fail closed: row có dữ liệu nhưng thiếu identity thì không tự dồn.
    if (!plate || !visitId) {
      return {
        ok: false,
        blocked: true,
        reason: 'PARTIAL_IDENTITY_ROW',
        row: row,
        plate: plate || '',
        visitId: visitId || ''
      };
    }

    if (seenIds[visitId]) {
      return {
        ok: false,
        blocked: true,
        reason: 'DUPLICATE_VISIT_ID',
        row: row,
        visitId: visitId,
        previousRow: seenIds[visitId]
      };
    }
    if (seenPlates[plate]) {
      return {
        ok: false,
        blocked: true,
        reason: 'DUPLICATE_ACTIVE_PLATE',
        row: row,
        plate: plate,
        previousRow: seenPlates[plate]
      };
    }

    seenIds[visitId] = row;
    seenPlates[plate] = row;
    activeVisitIds.push(visitId);
    activePlates.push(plate);
  }

  return {
    ok: true,
    start: start,
    last: last,
    display: display,
    activeVisitIds: activeVisitIds,
    activePlates: activePlates
  };
}

function safeCompactDispatchRowsUnderLock_(sh, deadlineMs) {
  // V13.19.4 SAFE ROW COMPACT
  // Dùng moveRows() của Spreadsheet service để di chuyển NGUYÊN HÀNG vật lý.
  // Không setValues / setDataValidations / snapshot block A:X.
  // Vì vậy dropdown chip, màu, format, note và hidden columns đi cùng record.
  if (!sh || sh.getName() !== CFG.SHEET_DISPATCH) {
    return {packedRows: 0, packedBlocks: 0, blocked: true, reason: 'INVALID_SHEET'};
  }
  if (isWebAppConfirmPriorityActive_()) {
    return {packedRows: 0, packedBlocks: 0, yieldedForWebApp: true};
  }

  const startedAt = Date.now();
  const localDeadline = Math.min(
    deadlineMs || Number.MAX_SAFE_INTEGER,
    startedAt + Number(CONCURRENCY_CFG.SAFE_ROW_COMPACT_MAX_MS || 3000)
  );

  const before = inspectDispatchForSafeCompact_(sh);
  if (!before.ok) {
    console.log('[WARN][SafeRowCompactBlocked] ' + JSON.stringify(before));
    return {
      packedRows: 0,
      packedBlocks: 0,
      blocked: true,
      reason: before.reason || 'PREFLIGHT_FAILED',
      details: before
    };
  }

  const expectedOrder = before.activeVisitIds.slice();
  let packedRows = 0;
  let packedBlocks = 0;

  while (Date.now() < localDeadline) {
    if (isWebAppConfirmPriorityActive_()) {
      return {
        packedRows: packedRows,
        packedBlocks: packedBlocks,
        yieldedForWebApp: true
      };
    }

    const last = sh.getLastRow();
    if (last < CFG.FIRST_DATA_ROW) break;

    const width = CFG.DP.LAST_SCC_STATUS;
    const display = sh
      .getRange(
        CFG.FIRST_DATA_ROW,
        1,
        last - CFG.FIRST_DATA_ROW + 1,
        width
      )
      .getDisplayValues();

    let gapIndex = -1;
    for (let i = 0; i < display.length; i++) {
      if (!dispatchRowHasAnyValueForCompact_(display[i])) {
        gapIndex = i;
        break;
      }
    }
    if (gapIndex < 0) break;

    let sourceIndex = -1;
    for (let i = gapIndex + 1; i < display.length; i++) {
      if (dispatchRowHasAnyValueForCompact_(display[i])) {
        sourceIndex = i;
        break;
      }
    }
    // Không còn data dưới gap => data đã liền, các row trống đều ở cuối.
    if (sourceIndex < 0) break;

    let sourceEndIndex = sourceIndex;
    while (
      sourceEndIndex + 1 < display.length &&
      dispatchRowHasAnyValueForCompact_(display[sourceEndIndex + 1])
    ) {
      sourceEndIndex++;
    }

    const destinationRow = CFG.FIRST_DATA_ROW + gapIndex;
    const sourceRow = CFG.FIRST_DATA_ROW + sourceIndex;
    const blockRows = sourceEndIndex - sourceIndex + 1;

    if (isWebAppConfirmPriorityActive_()) {
      return {
        packedRows: packedRows,
        packedBlocks: packedBlocks,
        yieldedForWebApp: true
      };
    }

    // moveRows chọn nguyên hàng theo rowSpec. destinationIndex dùng tọa độ
    // TRƯỚC khi move; vì sourceRow > destinationRow nên block sẽ nằm đúng gap.
    sh.moveRows(
      sh.getRange(sourceRow, 1, blockRows, 1),
      destinationRow
    );
    SpreadsheetApp.flush();
    packedRows += blockRows;
    packedBlocks++;
  }

  const after = inspectDispatchForSafeCompact_(sh);
  if (!after.ok) {
    console.log('[ERROR][SafeRowCompactPostCheck] ' + JSON.stringify(after));
    return {
      packedRows: packedRows,
      packedBlocks: packedBlocks,
      blocked: true,
      reason: 'POSTCHECK_FAILED',
      details: after
    };
  }

  const sameOrder =
    expectedOrder.length === after.activeVisitIds.length &&
    expectedOrder.every(function(id, idx) {
      return id === after.activeVisitIds[idx];
    });

  if (!sameOrder) {
    // moveRows không xóa dữ liệu; nếu verify order fail thì dừng và báo rõ,
    // KHÔNG rollback bằng setValues để tránh lặp lại lỗi V13.19.2.1.
    console.log(
      '[ERROR][SafeRowCompactOrderMismatch] ' +
      JSON.stringify({before: expectedOrder, after: after.activeVisitIds})
    );
    return {
      packedRows: packedRows,
      packedBlocks: packedBlocks,
      blocked: true,
      reason: 'ORDER_VERIFY_FAILED'
    };
  }

  return {
    packedRows: packedRows,
    packedBlocks: packedBlocks,
    blocked: false,
    elapsedMs: Date.now() - startedAt
  };
}

function clearDispatchOwnedRowOnly_(sh, row, expectedVisitId) {
  const last = sh.getLastRow();
  const width = CFG.DP.LAST_SCC_STATUS;
  if (row < CFG.FIRST_DATA_ROW || row > last) return false;

  const expectedId = clean_(expectedVisitId || '');
  if (!expectedId) {
    throw new Error('Không clear Điều phối: thiếu expected visitId.');
  }

  const currentId = clean_(sh.getRange(row, CFG.DP.VISIT_ID).getValue());
  if (currentId !== expectedId) {
    throw new Error('Không clear Điều phối: ownership row/visitId đã thay đổi.');
  }

  const ownerRows = resolveDispatchRowsByVisitId_(sh, expectedId);
  if (ownerRows.length !== 1 || ownerRows[0] !== row) {
    throw new Error(
      'Không clear Điều phối: visitId không duy nhất hoặc không còn ở dòng dự kiến.'
    );
  }

  if (isWebAppConfirmPriorityActive_()) {
    const err = new Error('Nhường WebApp Priority; chưa clear Điều phối.');
    err.code = 'WEBAPP_PRIORITY_YIELD';
    throw err;
  }

  const verifyId = clean_(sh.getRange(row, CFG.DP.VISIT_ID).getValue());
  if (verifyId !== expectedId) {
    throw new Error('Không clear Điều phối: visitId thay đổi ngay trước mutation.');
  }

  const rowRange = sh.getRange(row, 1, 1, width);
  // Chỉ xóa content + note. Format / dropdown / chip màu của row vẫn nguyên.
  rowRange.clearContent();
  rowRange.clearNote();
  SpreadsheetApp.flush();

  const idAfter = clean_(sh.getRange(row, CFG.DP.VISIT_ID).getValue());
  const plateAfter = clean_(readPlateText_(sh, row));
  if (idAfter || plateAfter) {
    throw new Error(
      'Clear Điều phối chưa sạch hoàn toàn tại dòng ' + row +
      '; cần kiểm tra thủ công.'
    );
  }
  return true;
}

function clearAndCompactDispatchRow_(sh, row, expectedVisitId) {
  // V13.19.4:
  // - clear đúng owner row trước.
  // - sau đó compact layout bằng moveRows() nguyên hàng.
  // Nếu compact bị BLOCK/yield, dữ liệu vẫn an toàn; chỉ còn gap và worker sau sẽ dồn tiếp.
  const cleared = clearDispatchOwnedRowOnly_(sh, row, expectedVisitId);
  if (!cleared) return false;

  const pack = safeCompactDispatchRowsUnderLock_(sh);
  console.log(
    '[PERF][SafeRowCompact] ' +
    JSON.stringify({
      packedRows: Number(pack.packedRows || 0),
      packedBlocks: Number(pack.packedBlocks || 0),
      blocked: !!pack.blocked,
      reason: clean_(pack.reason || ''),
      yieldedForWebApp: !!pack.yieldedForWebApp,
      elapsedMs: Number(pack.elapsedMs || 0)
    })
  );
  return true;
}
/*********************************************************
 * MENU DỌN XE
 *********************************************************/
function donDongDaHoanThanh() {
  return dongBoLaiDongBiKet();
}
/*********************************************************
 * GHI GIỜ HIỆN TẠI
 *********************************************************/
function ghiGioHienTai() {
  const ss =
    SpreadsheetApp.getActive();
  const sh =
    ss.getActiveSheet();
  const range =
    ss.getActiveRange();
  if (!range) return;
  const col =
    range.getColumn();
  const name =
    sh.getName();
  const allowed =
    (
      name ===
      CFG.SHEET_DISPATCH &&
      (
        col ===
        CFG.DP.GIO_NHAN ||
        col ===
        CFG.DP.GIO_GIAO
      )
    )
    ||
    (
      (
        name ===
        CFG.SHEET_KT ||
        name ===
        CFG.SHEET_KT_LAI
      ) &&
      (
        col === 7 ||
        col === 8
      )
    );
  if (!allowed) {
    SpreadsheetApp
      .getUi()
      .alert(
        'Chọn ô Giờ nhận / Giờ giao / Giờ bắt đầu / Giờ kết thúc.'
      );
    return;
  }
  const now =
    new Date();
  setTimeCell_(
    range,
    now.getHours(),
    now.getMinutes()
  );
}
/*********************************************************
 * ẨN DATA NGÀY CŨ
 *********************************************************/
function anDuLieuCu() {
  const ss =
    SpreadsheetApp.getActive();
  [
    CFG.SHEET_DATA_EM,
    CFG.SHEET_DATA_SCC
  ]
    .forEach(name => {
      const sh =
        ss.getSheetByName(
          name
        );
      if (
        !sh ||
        sh.getLastRow() < 2
      ) {
        return;
      }
      const last =
        sh.getLastRow();
      sh.showRows(
        2,
        last - 1
      );
      const values =
        sh
          .getRange(
            2,
            1,
            last - 1,
            1
          )
          .getValues();
      const today =
        todayOnly_()
          .getTime();
      let start =
        null;
      let count =
        0;
      for (
        let i = 0;
        i < values.length;
        i++
      ) {
        const value =
          values[i][0];
        const isToday =
          value instanceof Date &&
          stripTime_(
            value
          )
          .getTime() ===
          today;
        if (
          !isToday
        ) {
          if (
            start === null
          ) {
            start =
              i + 2;
          }
          count++;
        } else if (
          start !== null
        ) {
          sh.hideRows(
            start,
            count
          );
          start =
            null;
          count =
            0;
        }
      }
      if (
        start !== null
      ) {
        sh.hideRows(
          start,
          count
        );
      }
    });
}
/*********************************************************
 * HIỆN TOÀN BỘ DATA
 *********************************************************/
function hienToanBoData() {
  const ss =
    SpreadsheetApp.getActive();
  [
    CFG.SHEET_DATA_EM,
    CFG.SHEET_DATA_SCC
  ]
    .forEach(name => {
      const sh =
        ss.getSheetByName(
          name
        );
      if (!sh) return;
      if (
        sh.getMaxRows() > 1
      ) {
        sh.showRows(
          2,
          sh.getMaxRows() - 1
        );
      }
    });
}
/*********************************************************
 * FORMAT DATA
 *********************************************************/
function formatDataEMRow_(
  sh,
  row
) {
  sh
    .getRange(
      row,
      1
    )
    .setNumberFormat(
      'dd/MM/yyyy'
    );
  preparePlateCell_(
    sh,
    row
  );
  sh
    .getRange(
      row,
      9
    )
    .setNumberFormat(
      'HH:mm'
    );
  sh
    .getRange(
      row,
      10
    )
    .setNumberFormat(
      'dd/MM/yyyy'
    );
  sh
    .getRange(
      row,
      11,
      1,
      2
    )
    .setNumberFormat(
      'HH:mm'
    );
}
function formatDataSCCRow_(
  sh,
  row
) {
  sh.getRange(row, CFG.DATA_SCC.NGAY).setNumberFormat('dd/MM/yyyy');
  preparePlateCell_(sh, row);
  sh.getRange(row, CFG.DATA_SCC.GIO_NHAN).setNumberFormat('HH:mm');
  sh.getRange(row, CFG.DATA_SCC.NGAY_GIAO).setNumberFormat('dd/MM/yyyy');
  sh.getRange(row, CFG.DATA_SCC.GIO_GIAO, 1, 2).setNumberFormat('HH:mm');
}
function formatDispatchTimeCells_(
  sh,
  row
) {
  sh
    .getRange(
      row,
      CFG.DP.GIO_NHAN
    )
    .setNumberFormat(
      'HH:mm'
    );
  sh
    .getRange(
      row,
      CFG.DP.GIO_GIAO
    )
    .setNumberFormat(
      'HH:mm'
    );
}

/*********************************************************
 * HELPERS
 *********************************************************/

function editTouchesColumn_(
  range,
  col
) {
  return (
    range.getColumn() <= col &&
    range.getLastColumn() >= col
  );
}
function normalizePlate_(
  value
) {
  return String(
    value == null
      ? ''
      : value
  )
    .toUpperCase()
    .replace(
      /[^A-Z0-9]/g,
      ''
    );
}
function normalizeText_(value) {
  return normalizeKey_(value).toUpperCase();
}
function clean_(
  value
) {
  return String(
    value == null
      ? ''
      : value
  )
    .trim();
}
function normalizeKey_(
  value
) {
  return String(
    value || ''
  )
    .toLowerCase()
    .normalize(
      'NFD'
    )
    .replace(
      /[\u0300-\u036f]/g,
      ''
    )
    .replace(
      /đ/g,
      'd'
    )
    .replace(
      /\s+/g,
      ' '
    )
    .trim();
}
function splitMulti_(
  value
) {
  if (!value) {
    return [];
  }
  let text =
    String(
      value
    );
  // Bảo vệ token BD++ trước khi tách dấu cộng.
  const TOKEN_BD_PLUS =
    '__BD_PLUS_PLUS__';
  text =
    text.replace(
      /\bBD\s*\+\s*\+/gi,
      TOKEN_BD_PLUS
    );
  const items =
    text
      // Chỉ dấu + có khoảng trắng hai bên mới là dấu phân cách.
      // Ví dụ: "BD + Kiểm tra".
      // "BD++" đã được bảo vệ phía trên.
      .split(
        /\s+\+\s+|[,;\n]+/
      )
      .map(
        item =>
          clean_(
            String(
              item || ''
            ).replace(
              new RegExp(
                TOKEN_BD_PLUS,
                'g'
              ),
              'BD++'
            )
          )
      )
      .filter(
        Boolean
      );
  return uniqueArray_(
    items
  );
}
function pushUnique_(
  arr,
  value
) {
  if (
    arr.indexOf(
      value
    ) === -1
  ) {
    arr.push(
      value
    );
  }
}
function uniqueArray_(
  arr
) {
  return Array.from(
    new Set(
      arr
    )
  );
}
function valueExists_(
  value
) {
  return (
    value !== '' &&
    value !== null &&
    typeof value !==
      'undefined'
  );
}
function isComplete_(
  value
) {
  return (
    normalizeKey_(
      value
    ) ===
    'hoan thanh'
  );
}
function todayOnly_() {
  return stripTime_(
    new Date()
  );
}
function stripTime_(
  value
) {
  const d =
    new Date(
      value
    );
  d.setHours(
    0,
    0,
    0,
    0
  );
  return d;
}
function sameDate_(
  a,
  b
) {
  if (
    !(a instanceof Date) ||
    !(b instanceof Date)
  ) {
    return false;
  }
  return (
    stripTime_(a)
      .getTime() ===
    stripTime_(b)
      .getTime()
  );
}
function sameTime_(
  a,
  b
) {
  const aa =
    timeToMinutes_(
      a
    );
  const bb =
    timeToMinutes_(
      b
    );
  if (
    aa === null ||
    bb === null
  ) {
    return false;
  }
  return (
    aa === bb
  );
}
function timeToMinutes_(
  value
) {
  if (
    value instanceof Date
  ) {
    return (
      value.getHours() *
      60 +
      value.getMinutes()
    );
  }
  if (
    typeof value ===
    'number'
  ) {
    const fraction =
      (
        value % 1 +
        1
      ) % 1;
    return (
      Math.round(
        fraction *
        1440
      ) % 1440
    );
  }
  const text =
    String(
      value || ''
    );
  const match =
    text.match(
      /^(\d{1,2}):(\d{2})$/
    );
  if (
    !match
  ) {
    return null;
  }
  const hh =
    Number(
      match[1]
    );
  const mm =
    Number(
      match[2]
    );
  if (
    !validTime_(
      hh,
      mm
    )
  ) {
    return null;
  }
  return (
    hh * 60 +
    mm
  );
}
function spreadsheetForRecord_(rec) {
  if (rec && rec.ss) return rec.ss;
  return laySpreadsheetWeb_();
}
function visitSchema_() {
  return [
    {name: CFG.SHEET_DISPATCH, start: 6, header: 5, col: CFG.DP.VISIT_ID, fields: ['visitId', 'lastEMStatus', 'lastSCCStatus']},
    {name: CFG.SHEET_DATA_EM, start: 2, header: 1, col: 15, fields: ['visitId']},
    {name: CFG.SHEET_DATA_SCC, start: 2, header: 1, col: CFG.DATA_SCC.VISIT_ID, fields: ['visitId']},
    {name: CFG.SHEET_KT, start: 2, header: 1, col: 11, fields: ['visitId', 'trackingSide']},
    {name: CFG.SHEET_KT_LAI, start: 2, header: 1, col: 12, fields: ['visitId', 'trackingSide']}
  ];
}
function visitColumn_(name) {
  const spec = visitSchema_().find(x => x.name === name);
  if (!spec) throw new Error('Sheet không hỗ trợ visitId: ' + name);
  return spec.col;
}
function assertVisitSetup_(ss) {
  const configured = PropertiesService.getScriptProperties().getProperty('VISIT_SCHEMA_' + ss.getId());
  if (configured !== '4') throw new Error('VISIT schema v4 chưa được cấu hình. Dùng MAINTENANCE_ARCHIVE V13.16 để migration trước khi chạy Production.');
}
function allVisitIdSpecs_(ss) {
  const specs = repairVisitSchema_(ss).map(function(spec) {
    return {
      name: spec.name,
      start: spec.start,
      plateCol: 2,
      idCol: spec.col
    };
  });
  specs.push({
    name: CFG.SHEET_WAITING_DELIVERY,
    start: CFG.WAITING_DELIVERY_FIRST_ROW,
    plateCol: CFG.WD.BIEN_SO,
    idCol: CFG.WD.VISIT_ID
  });
  return specs;
}
function visitIdOccurrences_(ss, visitId) {
  visitId = clean_(visitId);
  if (!visitId) return [];
  const out = [];
  allVisitIdSpecs_(ss).forEach(function(spec) {
    const sh = ss.getSheetByName(spec.name);
    if (!sh) return;
    const last = sh.getLastRow();
    if (last < spec.start) return;
    const count = last - spec.start + 1;
    const maxCol = Math.max(spec.plateCol, spec.idCol);
    const values = sh.getRange(spec.start, 1, count, maxCol).getValues();
    const display = sh.getRange(spec.start, 1, count, maxCol).getDisplayValues();
    for (let i = 0; i < count; i++) {
      const id = clean_(values[i][spec.idCol - 1]);
      if (id !== visitId) continue;
      out.push({
        sheet: spec.name,
        row: spec.start + i,
        plate: normalizePlate_(display[i][spec.plateCol - 1]),
        idCol: spec.idCol,
        visitId: id
      });
    }
  });
  return out;
}
function generateUniqueVisitId_(ss) {
  for (let attempt = 0; attempt < 20; attempt++) {
    const id = Utilities.getUuid();
    if (!visitIdOccurrences_(ss, id).length) return id;
  }
  throw new Error('Không tạo được visitId duy nhất sau nhiều lần thử.');
}
/**
 * P0.5.8 – identity guard, FAIL CLOSED on ambiguous missing visitId.
 * New blank Dispatch entry: can be assigned a UUID as before.
 * Existing work/history without V: never silently generate a new UUID.
 * This check does not edit Data, finish times, waiting rows or queue.
 */
function p058GuardMissingDispatchVisitIdBeforeNewAllocation_(sh, row, plate) {
  if(!sh || sh.getName()!==CFG.SHEET_DISPATCH) return true;
  const own=sh.getRange(row,1,1,CFG.DP.LAST_SCC_STATUS).getDisplayValues()[0];
  const operational=[5,6,7,8,9,10,11,12,16,17,18,20,21,23,24];
  // These fields mean a business visit has already progressed, or was pasted.
  // Refuse to reinterpret it as a freshly photographed vehicle.
  const hasWork=operational.some(function(col){return !!clean_(own[col-1]);});
  if(hasWork){
    throw new Error('P058_EXISTING_VISIT_ID_MISSING: Điều phối dòng '+row+
      ', biển '+normalizePlate_(plate)+
      '. Có dữ liệu công việc nhưng V trống. Không cấp ID mới; kiểm tra log/Data để khôi phục đúng lượt.');
  }
  const date=clean_(own[CFG.DP.NGAY-1]);
  const advisor=normalizeKey_(own[CFG.DP.CVDV-1]);
  const target=normalizePlate_(plate);
  if(!date || !target) return true;
  const ss=sh.getParent(),found=[];
  [CFG.SHEET_DATA_EM,CFG.SHEET_DATA_SCC].forEach(function(name){
    const data=ss.getSheetByName(name);
    if(!data) throw new Error('P058_SOURCE_SHEET_MISSING: '+name);
    const last=data.getLastRow();
    if(last<2)return;
    // Only identity A:C and visitId O; bounded, no edits.
    const identity=data.getRange(2,1,last-1,3).getDisplayValues();
    const ids=data.getRange(2,15,last-1,1).getDisplayValues();
    for(let i=0;i<identity.length;i++){
      if(normalizePlate_(identity[i][1])!==target ||
         clean_(identity[i][0])!==date)continue;
      if(advisor && normalizeKey_(identity[i][2])!==advisor)continue;
      const oldId=clean_(ids[i][0]);
      if(oldId)found.push({source:name,row:i+2,visitId:oldId});
    }
  });
  if(found.length){
    const unique=Array.from(new Set(found.map(function(x){return x.visitId;})));
    throw new Error('P058_POSSIBLE_OLD_VISIT_ID: Điều phối dòng '+row+
      ', biển '+target+', ngày '+date+'. Có '+unique.length+
      ' visitId lịch sử cùng định danh. Dừng tạo UUID, xác minh lượt/giờ tiếp nhận; không tự tái dùng ID cũ.');
  }
  return true;
}

function ensureVisitId_(sh, row) {
  const plate = readPlateText_(sh, row);
  if (!plate) return '';
  const ss = sh.getParent();
  assertVisitSetup_(ss);
  const col = visitColumn_(sh.getName());
  const cell = sh.getRange(row, col);
  let id = clean_(cell.getValue());
  if (!id) {
    // P0.4: không cấp visitId mới cho xe còn hiện diện trong Điều phối
    // hoặc Xe chờ giao dưới một lượt khác (kể cả từ paste trực tiếp).
    if (sh.getName() === CFG.SHEET_DISPATCH) {
      p04AssertPhysicalVisitOwnership_(sh, row, plate);
      p058GuardMissingDispatchVisitIdBeforeNewAllocation_(sh, row, plate);
    }
    id = generateUniqueVisitId_(ss);
    cell.setNumberFormat('@').setValue(id);
    return id;
  }
  // Các sheet phụ chỉ mang visitId đã được cấp từ lượt xe.
  if (sh.getName() !== CFG.SHEET_DISPATCH) return id;
  const targetPlate = normalizePlate_(plate);
  const occurrences = visitIdOccurrences_(ss, id);
  const otherDispatchRows = occurrences.filter(function(item) {
    return item.sheet === CFG.SHEET_DISPATCH && item.row !== row;
  });
  const differentPlate = occurrences.filter(function(item) {
    return item.plate && item.plate !== targetPlate;
  });
  // FAIL-SAFE:
  // Không tự "heal" hàng loạt theo biển số + visitId vì có thể gộp nhầm
  // hai lượt khác nhau. Khi ownership mơ hồ, dừng sync để kiểm tra.
  if (otherDispatchRows.length || differentPlate.length) {
    throw new Error(
      'visitId bị trùng hoặc ownership không rõ tại Điều phối dòng ' +
      row +
      ' (' +
      targetPlate +
      '). Không tự sửa hàng loạt. Hãy chạy audit/repair visitId trước khi đồng bộ.'
    );
  }
  return id;
}
/**
 * P0.4.2 - XÁC NHẬN CHUNG EM/SCC, KHÔNG ĐỒNG NHẤT VỚI GIỜ KẾT THÚC THỰC TẾ.
 *
 * Data thiếu khi user bấm Hoàn thành: tạo bản ghi OPEN đúng visitId trước,
 * sau đó đồng bộ hoàn tất với mốc xác nhận đã ghi ở Điều phối. Tuyệt đối
 * không tự chữa lượt cũ chỉ từ một trạng thái đã Hoàn thành trên sheet.
 */
function p042RestoreQueuedConfirmation_(sh, row, statusCol, stampCol, e) {
  if (!e || !e.isRetryWorker) return;
  const evidence = p043ConfirmationEvidenceMap_(e)[String(statusCol)];
  if (!evidence || !isComplete_(sh.getRange(row, statusCol).getDisplayValue())) return;
  const stamp = new Date(evidence);
  if (isNaN(stamp.getTime()) || stamp.getTime() > Date.now() + 60000) {
    throw new Error('P043_QUEUED_EVENT_TIME_INVALID: ' + statusCol);
  }
  const latestCommittedCol = statusCol === CFG.DP.TD_EM ?
    CFG.DP.LAST_EM_STATUS : CFG.DP.LAST_SCC_STATUS;
  // A completed, already committed status must not be restamped by a
  // stale queued event or a repeat click.
  if (isComplete_(sh.getRange(row, latestCommittedCol).getDisplayValue())) return;
  const cell = sh.getRange(row, stampCol);
  const previous = cell.getValue();
  if (previous instanceof Date && !isNaN(previous.getTime()) &&
      previous.getTime() > stamp.getTime() + 1000) {
    throw new Error('P043_QUEUED_EVENT_ORDER_CONFLICT: ' + statusCol +
      '; do not overwrite a newer stored timestamp.');
  }
  // If T/U already holds the start/doing status time, the completion
  // event MUST replace it. P0.4.2 only restored an empty T/U cell.
  if (!(previous instanceof Date) || previous.getTime() !== stamp.getTime()) {
    cell.setValue(stamp).setNumberFormat('dd/MM/yyyy HH:mm:ss');
    console.log('[P043_RESTORED_CONFIRMATION] ' + row + ' / ' + statusCol);
  }
}

function p042ConfirmationStamp_(rec, isEM) {
  return isEM ? rec.emConfirmationStamp : rec.sccConfirmationStamp;
}
function p042CompletionStamp_(rec, isEM, requireStamp) {
  const stamp = p042ConfirmationStamp_(rec, isEM);
  if (stamp instanceof Date && !isNaN(stamp.getTime()) &&
      stamp.getTime() <= Date.now() + 60000 &&
      (!(rec.date instanceof Date) ||
       stamp.getTime() >= new Date(rec.date.getFullYear(), rec.date.getMonth(), rec.date.getDate()).getTime())) {
    return stamp;
  }
  if (requireStamp) {
    throw new Error('P042_CONFIRMATION_STAMP_MISSING_OR_INVALID: ' +
      (isEM ? 'EM' : 'SCC') + ' / ' + normalizePlate_(rec.plate) + ' / ' + rec.visitId);
  }
  return null;
}
// Mốc hoàn thành dùng cho báo cáo = mốc Điều phối xác nhận. Khi một
// công việc Kiểm tra đã thực sự bắt đầu SAU mốc xác nhận riêng của nó,
// chỉ được dùng mốc xác nhận chung muộn hơn nếu hai nhánh đều hoàn thành.
// Không tự lùi về giờ worker hay đoán giờ lái thử.

/**
 * P0.4.9.1 - Chứng cứ QC3 từ nguồn Data EM/SCC CHỈ khi có xác nhận
 * nghiệp vụ cụ thể. Đọc và đối chiếu, không sửa Data, không tạo timestamp.
 *
 * Ràng buộc: đúng 01 Data visitId + biển số, status Hoàn thành, giờ HH:mm
 * có ghi chú xác nhận của người vận hành chứa cùng xe/visitId/ngày/giờ.
 * Không suy diễn T/U từ trạng thái cũ hoặc lấy thời gian Worker.
 * Chỉ dùng cho phiên KT bắt đầu sau mốc P/Q, cùng ngày sửa chữa.
 */
function p0491VerifiedQcCompletionFromData_(rec, isEM, began) {
  if (!(began instanceof Date) || isNaN(began.getTime())) return null;
  if (!(rec.date instanceof Date) || isNaN(rec.date.getTime())) return null;
  var ss = spreadsheetForRecord_(rec);
  var tz = ss.getSpreadsheetTimeZone();
  var day = Utilities.formatDate(rec.date, tz, 'dd/MM/yyyy');
  if (Utilities.formatDate(began, tz, 'dd/MM/yyyy') !== day) return null;
  var name = isEM ? CFG.SHEET_DATA_EM : CFG.SHEET_DATA_SCC;
  var data = ss.getSheetByName(name);
  if (!data) return null;
  var visitId = clean_(rec.visitId), plate = normalizePlate_(rec.plate);
  if (!visitId || !plate) return null;
  // Không chỉ dùng findDataRow_: phải xác minh visitId là duy nhất.
  var last = data.getLastRow();
  if (last < 2) return null;
  var idColumn = isEM ? 15 : CFG.DATA_SCC.VISIT_ID;
  var ids = data.getRange(2, idColumn, last - 1, 1).getDisplayValues();
  var hits = [];
  ids.forEach(function(v,i) { if (clean_(v[0]) === visitId) hits.push(i + 2); });
  if (hits.length !== 1) return null;
  var row = hits[0];
  if (normalizePlate_(data.getRange(row, 2).getDisplayValue()) !== plate) return null;
  if (!isComplete_(data.getRange(row, 13).getDisplayValue())) return null;
  var cell = data.getRange(row, 12);
  var timeText = clean_(cell.getDisplayValue());
  var match = /^(\d{1,2}):(\d{2})$/.exec(timeText);
  if (!match) return null;
  var hour=Number(match[1]), min=Number(match[2]);
  if (hour > 23 || min > 59) return null;
  var raw=cell.getValue();
  if (!(raw instanceof Date) || isNaN(raw.getTime())) return null;
  var note = String(cell.getNote() || '');
  var statusNote = String(data.getRange(row, 13).getNote() || '');
  // Cần ghi chú nguồn trên chính mốc giờ; trạng thái không đủ làm bằng chứng.
  if (!/xác nhận nghiệp vụ từ người vận hành|nguồn xác nhận\s*:\s*người vận hành/i.test(note)) return null;
  if (note.indexOf(visitId) === -1 || note.indexOf(plate) === -1) return null;
  if (note.indexOf(day) === -1 || note.indexOf(timeText) === -1) return null;
  if (!/hoàn thành/i.test(note)) return null;
  if (statusNote && statusNote.indexOf(visitId) === -1) return null;
  var effective = new Date(began.getFullYear(),began.getMonth(),began.getDate(),hour,min,0,0);
  if (effective.getTime() <= began.getTime() || effective.getTime() > Date.now() + 60000) return null;
  // Chặn lệch ngày khi chuyển đổi timezone/number format.
  if (Utilities.formatDate(effective,tz,'dd/MM/yyyy') !== day) return null;
  effective.p0491Proof = {sheet:name, row:row, time:timeText, visitId:visitId, plate:plate};
  return effective;
}

function p042ResolvedCompletionStamp_(rec, isEM) {
  const stamp = p042CompletionStamp_(rec, isEM, true);
  const side = isEM ? 'EM' : 'SCC';
  const route = classifyWork_(rec);
  const targets = trackingTargets_(rec, route).filter(function(t) { return t.side === side; });
  const ss = spreadsheetForRecord_(rec);
  for (let i = 0; i < targets.length; i++) {
    const t = targets[i];
    const sh = ss.getSheetByName(t.sheet);
    if (!sh) throw new Error('P042_TRACKING_SHEET_MISSING: ' + t.sheet);
    const row = findTrackingRow_(sh, rec, t.repeat, side);
    if (!row) continue;
    const began = sh.getRange(row, 7).getValue();
    if (!(began instanceof Date) || isNaN(began.getTime()) ||
        stamp.getTime() >= began.getTime()) continue;
    // P0.4.9.1: giờ thực tế đã xác nhận tại Data có thể muộn hơn mốc
    // thay trạng thái T/U cũ. Chỉ dùng khi provenance khớp mọi chốt.
    const sourceBound = p0491VerifiedQcCompletionFromData_(rec, isEM, began);
    if (sourceBound) {
      console.log('[P0491_QC_SOURCE_BOUND] ' + JSON.stringify(sourceBound.p0491Proof));
      return sourceBound;
    }
    const otherStatus = isEM ? rec.sccStatus : rec.emStatus;
    const otherStamp = p042CompletionStamp_(rec, !isEM, false);
    if (isComplete_(otherStatus) && otherStamp &&
        otherStamp.getTime() >= began.getTime()) {
      console.log('[P042_SHARED_CONFIRMATION_AFTER_KT] ' + rec.plate + ' / ' + rec.visitId);
      return otherStamp;
    }
    throw new Error('P042_CONFIRMATION_BEFORE_INSPECTION_START: ' + side +
      ' / ' + rec.plate + ' / ' + rec.visitId +
      '. Phải xác nhận lại sau giờ bắt đầu KT, không tự tạo giờ kết thúc.');
  }
  return stamp;
}
/**
 * P0.4.3: Data may exist as OPEN while P/Q displays Complete.
 * If last committed W/X is not Complete, T/U may still belong to
 * the preceding 'Đang làm' event. Never auto-infer completion.
 */
function p043GuardPendingCompletedStatus_(rec, route, e) {
  const ss = spreadsheetForRecord_(rec);
  [{active:route.em.active, status:rec.emStatus,
    committed:rec.lastEMStatus, col:CFG.DP.TD_EM, sheet:CFG.SHEET_DATA_EM},
   {active:route.scc.active, status:rec.sccStatus,
    committed:rec.lastSCCStatus, col:CFG.DP.TD_SCC, sheet:CFG.SHEET_DATA_SCC}]
  .forEach(function(t) {
    if (!t.active || !isComplete_(t.status)) return;
    // A committed W/X alone is not proof if Data source is still OPEN.
    const stored = ss.getSheetByName(t.sheet);
    if (!stored) throw new Error('P043_DATA_SHEET_MISSING: ' + t.sheet);
    const row = findDataRow_(stored, rec);
    if (row && isComplete_(stored.getRange(row, 13).getDisplayValue())) return;
    const directClick = !!(e && !e.isRetryWorker && !e.isRecovery &&
      !e.isReconciliation && editTouchesColumn_(e.range, t.col));
    const queuedClick = p043HasConfirmationEvidence_(e, t.col);
    if (!directClick && !queuedClick) {
      const err = new Error('P043_PENDING_COMPLETION_EVENT_REQUIRED: ' + t.sheet +
        ' / ' + normalizePlate_(rec.plate) + ' / ' + rec.visitId +
        '. Mốc T/U có thể thuộc trạng thái cũ; không tự gán giờ Hoàn thành.');
      err.code = 'P043_PENDING_COMPLETION_EVENT_REQUIRED';
      throw err;
    }
  });
}
function p042EnsureWorkRowsBeforeCompletion_(rec, route, e) {
  const ss = spreadsheetForRecord_(rec);
  const configs = [
    {isEM: true, active: route.em.active, status: rec.emStatus,
     name: CFG.SHEET_DATA_EM, work: route.em.workText},
    {isEM: false, active: route.scc.active, status: rec.sccStatus,
     name: CFG.SHEET_DATA_SCC, work: route.scc.workText}
  ];
  const missingCompleted = [];
  configs.forEach(function(item) {
    if (!item.active) return;
    const data = ss.getSheetByName(item.name);
    if (!data) throw new Error('P042_WORK_SHEET_MISSING: ' + item.name);
    const existing = findDataRow_(data, rec);
    if (existing) {
      const existingPlate = normalizePlate_(data.getRange(existing, 2).getDisplayValue());
      if (existingPlate !== normalizePlate_(rec.plate)) {
        throw new Error('P042_STORED_PLATE_MISMATCH: ' + item.name + ' / ' + rec.visitId);
      }
    } else if (isComplete_(item.status)) {
      // Mốc T/U xác nhận được ghi trong quá trình onEdit và giữ qua worker.
      // Lượt lịch sử thiếu Data, chỉ có status Hoàn thành nhưng không có
      // sự kiện mới KHÔNG được auto bootstrap hay bịa giờ thực hiện.
      if (!e || e.isReconciliation || e.isRecovery ||
          (!p043HasConfirmationEvidence_(e, item.isEM ? CFG.DP.TD_EM : CFG.DP.TD_SCC) &&
           (e.isRetryWorker ||
            !editTouchesColumn_(e.range, item.isEM ? CFG.DP.TD_EM : CFG.DP.TD_SCC)))) {
        throw new Error('P04_COMPLETION_WITHOUT_STORED_WORK: ' + item.name +
          ' / ' + normalizePlate_(rec.plate) + ' / ' + rec.visitId +
          '. Cần sự kiện xác nhận Hoàn thành, không tự tạo Data lịch sử.');
      }
      p042ResolvedCompletionStamp_(rec, item.isEM);
      missingCompleted.push(item);
    }
  });
  // Preflight TOÀN BỘ hai luồng trước khi ghi một dòng mới, tránh nửa chừng.
  missingCompleted.forEach(function(item) {
    const problems = item.isEM ? validateEMCompletion_(rec, route) :
      validateSCCCompletion_(rec, route);
    if (problems.length) throw new Error('P042_COMPLETION_VALIDATION_FAILED: ' +
      item.name + ' / ' + problems.join(', '));
  });
  missingCompleted.forEach(function(item) {
    // Mở bản ghi trước, sau đó sync chuẩn chuyển Hoàn thành.
    // Không tạo trực tiếp một Data 'Hoàn thành' mới.
    const openRec = Object.assign({}, rec, item.isEM ?
      {emStatus: 'Đang làm'} : {sccStatus: 'Đang làm'});
    syncDataForVisit_(openRec, item.work, item.isEM);
    const data = ss.getSheetByName(item.name);
    const row = findDataRow_(data, rec);
    if (!row) throw new Error('P042_BOOTSTRAP_VERIFY_FAILED: ' + item.name);
    data.getRange(row, 12).setNote(
      'P0.4.2: Data OPEN được tạo từ sự kiện xác nhận Hoàn thành của Điều phối.' +
      ' Thời gian ghi vào cột này là mốc XÁC NHẬN, không phải giờ kết thúc' +
      ' thực tế riêng của hạng mục. visitId=' + rec.visitId
    );
    console.log('[P042_BOOTSTRAP_PENDING] ' + item.name + ' / ' + rec.plate +
      ' / ' + rec.visitId);
  });
}

function syncDataForVisit_(rec, workText, isEM) {
  const name = isEM ? CFG.SHEET_DATA_EM : CFG.SHEET_DATA_SCC;
  const sh = spreadsheetForRecord_(rec).getSheetByName(name);
  if (!sh) throw new Error('Không tìm thấy ' + name);
  // P0.5.5: live, read-only owner assertion precedes ALL Data writes.
  // Enforces same visitId + plate + advisor; never auto-fixes collisions.
  p055AssertDataWriteOwner_(spreadsheetForRecord_(rec), rec, name);
  const foundRow = findDataRow_(sh, rec);
  if (!foundRow) p055AssertOtherDataOwnerBeforeInsert_(spreadsheetForRecord_(rec), rec, name);
  const status = isEM ? rec.emStatus : rec.sccStatus;
  // P0.4: Không tạo mới một công việc đã Hoàn thành nếu Data gốc
  // chưa hề có. Nếu tự tạo và gắn new Date() sẽ ngụy tạo thời gian hoàn thành.
  // Giữ nguyên Điều phối để đối chiếu biên bản/công việc thật.
  if (!foundRow && isComplete_(status)) {
    throw new Error('P04_COMPLETION_WITHOUT_STORED_WORK: ' +
      name + ' / ' + normalizePlate_(rec.plate) + ' / ' + clean_(rec.visitId) +
      '. Không tạo Data Hoàn thành không có bằng chứng.');
  }
  const row = foundRow || Math.max(sh.getLastRow() + 1, 2);
  // P0.4.5: safely extend Data EM/SCC before inserting the next record.
  if (!foundRow && row > sh.getMaxRows()) {
    sh.insertRowsAfter(sh.getMaxRows(), Math.max(25, row - sh.getMaxRows()));
    console.log('[P045_DATA_CAPACITY_EXTENDED] ' + name + ' | new max rows=' + sh.getMaxRows());
  }
  const completedCol = isEM ? 12 : CFG.DATA_SCC.GIO_HOAN_THANH;
  let completed = sh.getRange(row, completedCol).getValue();
  if (isComplete_(status)) {
    // Chỉ mốc XÁC NHẬN đã lưu ở Điều phối; không tạo giờ kết thúc thực tế
    // bằng thời điểm worker xử lý. Dòng cũ đã có mốc thì giữ nguyên.
    completed = completed || p042ResolvedCompletionStamp_(rec, isEM);
  } else completed = '';
  const common = [rec.date, plateAsText_(rec.plate), rec.advisor, rec.model];
  const values = isEM
    ? common.concat([
        rec.emGroup, workText, rec.bdLevel, rec.gtgt, rec.receiveTime,
        rec.deliveryDate, rec.deliveryTime, completed, status, rec.note, rec.visitId
      ])
    : common.concat([
        rec.inspector, workText, rec.itemCount, rec.ktv1,
        rec.receiveTime, rec.deliveryDate, rec.deliveryTime, completed,
        status, rec.note, rec.visitId
      ]);
  preparePlateCell_(sh, row);
  sh.getRange(row, 1, 1, values.length).setValues([values.map(v => v == null ? '' : v)]);
  if (isEM) formatDataEMRow_(sh, row);
  else formatDataSCCRow_(sh, row);
  invalidateHistoryCache_(spreadsheetForRecord_(rec), rec.plate);
  // P0.5.7: clear receipt ONLY AFTER a Completed Data row has been written.
  if (isComplete_(status)) p057ClearCompletedReceiptAfterDataCommit_(rec,isEM);
}
function isValidTimeValue_(value) {
  if (value instanceof Date) return !isNaN(value.getTime());
  if (typeof value === 'number') return Number.isFinite(value) && value >= 0 && value < 1;
  const m = String(value == null ? '' : value).trim().match(/^(\d{1,2}):(\d{2})$/);
  return !!m && validTime_(Number(m[1]), Number(m[2]));
}
function trackingTargets_(rec, route) {
  const targets = [];
  // Theo dõi KT thường chỉ được kích hoạt bởi CẤP QC = 3.
  // Người KT / Nhóm EM / Nhóm SCC không được phép override rule này.
  // Riêng Kiểm tra lại giữ nguyên luồng Theo dõi KT lại hiện hữu.
  const qcLevel =
    clean_(rec && rec.qcLevel);
  const allowNormalInspectionTracking =
    qcLevel === '3';
  ['SCC', 'EM'].forEach(side => {
    const r = side === 'EM' ? route.em : route.scc;
    if (!r.active) return;
    if (
      allowNormalInspectionTracking &&
      r.items.includes('Kiểm tra')
    ) {
      targets.push({
        sheet: CFG.SHEET_KT,
        side: side,
        repeat: false
      });
    }
    if (r.items.includes('Kiểm tra lại')) {
      targets.push({
        sheet: CFG.SHEET_KT_LAI,
        side: side,
        repeat: true
      });
    }
  });
  return targets;
}
function trackingIsRunning_(status, side) {
  return normalizeKey_(status) === 'dang lam';
}
function validateTrackingBeforeClear_(rec, route) {
  const errors = [];
  trackingTargets_(rec, route).forEach(t => {
    const sh = spreadsheetForRecord_(rec).getSheetByName(t.sheet);
    if (!sh) { errors.push('Thiếu ' + t.sheet); return; }
    const row = findTrackingRow_(sh, rec, t.repeat, t.side);
    if (!row) { errors.push('Thiếu lượt ' + t.sheet + '/' + t.side); return; }
    if (rec.inspectionText && !clean_(sh.getRange(row, 5).getValue())) errors.push('Thiếu nội dung KT ' + t.side);
    if (!clean_(sh.getRange(row, 6).getValue())) errors.push('Thiếu người thực hiện KT ' + t.side);
    const status = t.side === 'EM' ? rec.emStatus : rec.sccStatus;
    if (isComplete_(status) && !sh.getRange(row, 8).getValue()) errors.push('Thiếu giờ kết thúc KT ' + t.side);
  });
  return errors;
}
function dongBoLaiDongDangChon() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Mở Google Sheets và chọn dòng cần đồng bộ.');
  const range = ss.getActiveRange();
  if (!range || range.getSheet().getName() !== CFG.SHEET_DISPATCH) throw new Error('Chọn dòng trên Điều phối.');
  xuLyChinhSua({range: range.getSheet().getRange(range.getRow(), CFG.DP.GHI_CHU, range.getNumRows(), 1), source: ss});
}
function syncVisitIdentity_(rec) {
  const ss = spreadsheetForRecord_(rec), updates = [];
  [CFG.SHEET_DATA_EM, CFG.SHEET_DATA_SCC].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) throw new Error('Không tìm thấy ' + name);
    const row = findDataRow_(sh, rec);
    if (row) updates.push({sh: sh, row: row, timeCol: name === CFG.SHEET_DATA_EM ? 9 : CFG.DATA_SCC.GIO_NHAN});
  });
  updates.forEach(u => {
    u.sh.getRange(u.row, 1, 1, 4).setValues([[rec.date, plateAsText_(rec.plate), rec.advisor, rec.model]]);
    u.sh.getRange(u.row, u.timeCol, 1, 3).setValues([[rec.receiveTime, rec.deliveryDate, rec.deliveryTime]]);
  });
}
/*********************************************************
 * WEB APP INSERT - PREFLIGHT VÀ ROLLBACK THEO VISIT ID
 *********************************************************/
function validateWebAppInsertPreflight_(ss) {
  if (!ss) throw new Error('Không xác định được Spreadsheet để ghi Web App.');
  assertVisitSetup_(ss);
  const errors = [];
  visitSchema_().forEach(spec => {
    const sh = ss.getSheetByName(spec.name);
    if (!sh) { errors.push('Thiếu sheet ' + spec.name); return; }
    if (spec.col + spec.fields.length - 1 > sh.getMaxColumns()) {
      errors.push(spec.name + ': thiếu cột metadata v3.');
      return;
    }
    const headers = sh.getRange(spec.header, spec.col, 1, spec.fields.length).getValues()[0];
    spec.fields.forEach((field, i) => {
      if (clean_(headers[i]) !== field) errors.push(spec.name + ': thiếu header metadata ' + field + '.');
    });
  });
  if (errors.length) throw new Error('Preflight Web App thất bại: ' + [...new Set(errors)].join(' | '));
}
function rollbackVisitFromWebApp_(ss, dispatchSheet, row, visitId) {
  visitId = clean_(visitId);
  if (!ss || !dispatchSheet || !visitId || row < CFG.FIRST_DATA_ROW) {
    throw new Error('Thiếu thông tin rollback Web App.');
  }
  if (dispatchSheet.getParent().getId() !== ss.getId() ||
      dispatchSheet.getName() !== CFG.SHEET_DISPATCH) {
    throw new Error('Rollback không đúng Spreadsheet/sheet Điều phối.');
  }
  const dispatchId = clean_(dispatchSheet.getRange(row, visitColumn_(CFG.SHEET_DISPATCH)).getValue());
  if (dispatchId !== visitId) {
    throw new Error('Không rollback vì visitId dòng Điều phối không khớp lượt vừa tạo.');
  }
  visitSchema_().filter(spec => spec.name !== CFG.SHEET_DISPATCH).forEach(spec => {
    const sh = ss.getSheetByName(spec.name);
    if (!sh || sh.getLastRow() < spec.start) return;
    const ids = sh.getRange(spec.start, spec.col, sh.getLastRow() - spec.start + 1, 1).getValues();
    const rows = [];
    ids.forEach((value, index) => {
      if (clean_(value[0]) === visitId) rows.push(spec.start + index);
    });
    rows.sort((a, b) => b - a).forEach(targetRow => sh.deleteRow(targetRow));
  });
  dispatchSheet.getRange(row, 1, 1, CFG.DP.GHI_CHU).clearContent().clearNote();
  dispatchSheet.getRange(row, visitColumn_(CFG.SHEET_DISPATCH), 1, 3).clearContent();
  SpreadsheetApp.flush();
  return true;
}
function inspectionRules_(rec) {
  const keys = splitMulti_(rec && rec.workRaw).map(normalizeKey_);
  const typeKey = normalizeKey_(rec && rec.inspectionType);
  const hasRepeat =
    typeKey === 'kiem tra lai' ||
    keys.indexOf('kiem tra lai') >= 0;
  const hasInspection =
    hasRepeat ||
    typeKey === 'kiem tra' ||
    keys.indexOf('kiem tra') >= 0;
  const hasPerformer = !!(
    clean_(rec && rec.inspector) ||
    clean_(rec && rec.ktv1) ||
    clean_(rec && rec.ktv2) ||
    clean_(rec && rec.emGroup)
  );
  return {
    hasInspection: hasInspection,
    hasRepeat: hasRepeat,
    hasPerformer: hasPerformer
  };
}
function addInspectionAssignmentMissing_(rec, missing) {
  const rules = inspectionRules_(rec);
  if (rules.hasRepeat && !clean_(rec.inspector)) {
    missing.push('Người KT');
  } else if (rules.hasInspection && !rules.hasPerformer) {
    missing.push('Người/nhóm thực hiện kiểm tra');
  }
}
/*********************************************************
 * P0.3 - IDENTITY EDIT GUARD
 * - chỉ chạy khi người dùng sửa cột B hoặc V của Điều phối
 * - không tự sửa Data EM/SCC, không gộp visitId
 * - trường hợp nhập nhầm 1 ô: chỉ phục hồi theo e.oldValue khi
 *   hai nguồn Data EM/SCC đều chứng thực cùng một biển số.
 * - multi-cell hoặc dữ liệu mâu thuẫn: FAIL CLOSED, không đồng bộ.
 *********************************************************/
function p03StoredPlateForVisit_(ss, visitId) {
  visitId = clean_(visitId);
  if (!visitId) return '';
  const plates = [];
  [CFG.SHEET_DATA_EM, CFG.SHEET_DATA_SCC].forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) throw new Error('P03_IDENTITY_DATA_SHEET_MISSING: ' + name);
    const last = sh.getLastRow();
    if (last < 2) return;
    const ids = sh.getRange(2, visitColumn_(name), last - 1, 1).getDisplayValues();
    const matchRows = [];
    ids.forEach(function(v, idx) {
      if (clean_(v[0]) === visitId) matchRows.push(idx + 2);
    });
    if (matchRows.length > 1) {
      throw new Error('P03_IDENTITY_DUPLICATE_STORED_VISIT: ' + name + '/' + visitId);
    }
    if (matchRows.length === 1) {
      const plate = normalizePlate_(sh.getRange(matchRows[0], 2).getDisplayValue());
      if (!plate) throw new Error('P03_IDENTITY_SOURCE_PLATE_MISSING: ' + name);
      plates.push(plate);
    }
  });
  if (!plates.length) return '';
  if (plates.some(function(p) { return p !== plates[0]; })) {
    throw new Error('P03_IDENTITY_CONFLICT_BETWEEN_EM_SCC: ' + visitId);
  }
  return plates[0];
}
function guardEditedDispatchIdentityP03_(sh, e) {
  const range = e && e.range;
  if (!sh || !range) return;
  const left = range.getColumn();
  const right = range.getLastColumn();
  const plateEdited = left <= CFG.DP.BIEN_SO && right >= CFG.DP.BIEN_SO;
  const idEdited = left <= CFG.DP.VISIT_ID && right >= CFG.DP.VISIT_ID;
  if (!plateEdited && !idEdited) return;
  const isSingle = range.getNumRows() === 1 && range.getNumColumns() === 1;
  const first = Math.max(CFG.FIRST_DATA_ROW, range.getRow());
  const last = range.getLastRow();
  for (let row = first; row <= last; row++) {
    const idCell = sh.getRange(row, CFG.DP.VISIT_ID);
    const id = clean_(idCell.getValue());
    const currentPlate = normalizePlate_(readPlateText_(sh, row));
    // P0.5.8.1: check an attempted V edit BEFORE skipping a blank ID/plate.
    // A single-cell deletion/change has e.oldValue and can be restored.
    // Multi-cell paste lacks oldValue: fail closed and leave an explicit error;
    // protect column V against manual paste in Production to prevent that case.
    if (idEdited) {
      const oldId = isSingle ? clean_(e.oldValue || '') : '';
      if (oldId && oldId !== id) {
        // A delayed onEdit must not restore a previous car's ID to a moved row.
        const previousOwners = visitIdOccurrences_(sh.getParent(), oldId);
        if (previousOwners.some(function(owner) {
          return (owner.sheet === CFG.SHEET_DISPATCH && owner.row !== row) ||
            (currentPlate && owner.plate && owner.plate !== currentPlate);
        })) {
          throw new Error('P058_OLD_ID_OWNER_CONFLICT: dòng ' + row +
            ', visitId=' + oldId + '. Giữ nguyên, cần xác minh lượt.');
        }
        idCell.setNumberFormat('@').setValue(oldId);
        SpreadsheetApp.flush();
        console.warn('[P058_ID_RESTORED_FROM_EDIT] row=' + row);
        throw new Error('P058_VISIT_ID_RESTORED: đã phục hồi visitId cũ tại Điều phối!' + row);
      }
      console.error('[P058_ID_EDIT_BLOCKED] row=' + row +
        ', currentId=' + id + ', oldValueAvailable=' + !!oldId +
        ', range=' + range.getA1Notation());
      throw new Error('P058_VISIT_ID_EDIT_BLOCKED: dòng ' + row +
        '. Không được sửa/xóa/dán cột V; nếu V trống sau paste cần đối soát nguồn, không cấp UUID mới.');
    }
    if (!currentPlate || !id) continue;
    const sourcePlate = p03StoredPlateForVisit_(sh.getParent(), id);
    if (sourcePlate && sourcePlate !== currentPlate) {
      const oldPlate = isSingle ? normalizePlate_(e.oldValue || '') : '';
      if (oldPlate && oldPlate === sourcePlate) {
        sh.getRange(row, CFG.DP.BIEN_SO).setNumberFormat('@').setValue(sourcePlate);
        console.log('[P03][PlateRestored] row=' + row + ' plate=' + sourcePlate);
        throw new Error('P03_PLATE_EDIT_ROLLED_BACK: đã phục hồi biển số ' + sourcePlate + ' tại dòng ' + row);
      }
      throw new Error('P03_PLATE_OWNER_CONFLICT: dòng ' + row +
        ' mang visitId của biển số ' + sourcePlate +
        ', không đồng bộ/không đoán biển số khi thay đổi nhiều ô.');
    }
  }
}

/*********************************************************
 * CHẶN METADATA BỊ COPY/FILL SANG VISIT KHÁC
 *********************************************************/
function sanitizeCopiedVisitMetadata_(sh, editedRange) {
  if (!sh || sh.getName() !== CFG.SHEET_DISPATCH || !editedRange) return [];
  const first = Math.max(editedRange.getRow(), CFG.FIRST_DATA_ROW);
  const last = editedRange.getRow() + editedRange.getNumRows() - 1;
  if (last < first || sh.getLastRow() < CFG.FIRST_DATA_ROW) return [];
  const visitCol = visitColumn_(CFG.SHEET_DISPATCH);
  const count = sh.getLastRow() - CFG.FIRST_DATA_ROW + 1;
  const allIds = sh.getRange(CFG.FIRST_DATA_ROW, visitCol, count, 1)
    .getValues()
    .map(r => clean_(r[0]));
  for (let row = first; row <= last; row++) {
    if (!readPlateText_(sh, row)) continue;
    const id = clean_(sh.getRange(row, visitCol).getValue());
    if (!id) continue;
    const matches = [];
    allIds.forEach((value, index) => {
      if (value === id) matches.push(CFG.FIRST_DATA_ROW + index);
    });
    if (matches.length > 1) {
      throw new Error(
        'visitId ' + id +
        ' xuất hiện ở nhiều dòng Điều phối (' +
        matches.join(', ') +
        '). Không tự reset metadata vì ownership không rõ.'
      );
    }
  }
  return [];
}
/*********************************************************
 * AUTO MODEL TỪ LỊCH SỬ DATA EM / DATA SCC
 *********************************************************/
function loadModelHistory_(ss) {
  const records = [];
  [
    {name: CFG.SHEET_DATA_EM, timeCol: 9, idCol: 15},
    {name: CFG.SHEET_DATA_SCC, timeCol: CFG.DATA_SCC.GIO_NHAN, idCol: CFG.DATA_SCC.VISIT_ID}
  ].forEach(spec => {
    const dataSheet = ss.getSheetByName(spec.name);
    if (!dataSheet) return;
    const last = dataSheet.getLastRow();
    if (last < 2) return;
    const values = dataSheet.getRange(2, 1, last - 1, spec.idCol).getValues();
    values.forEach((historyRow, index) => {
      const plate = normalizePlate_(historyRow[1]);
      const date = historyRow[0];
      if (!plate || !(date instanceof Date) || isNaN(date.getTime())) return;
      const minutes = timeToMinutes_(historyRow[spec.timeCol - 1]);
      records.push({
        sheet: spec.name,
        row: index + 2,
        plate: plate,
        dateValue: new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime(),
        minutes: minutes === null ? -1 : minutes,
        advisor: clean_(historyRow[2]),
        model: clean_(historyRow[3]),
        visitId: clean_(historyRow[spec.idCol - 1])
      });
    });
  });
  return records;
}
function historyCacheKey_(ss, plate) {
  return 'VISIT_HISTORY_V2_' + ss.getId() + '_' + normalizePlate_(plate);
}
function getHistoryFromCache_(ss, plate) {
  try {
    const raw = CacheService.getScriptCache().get(historyCacheKey_(ss, plate));
    if (raw === null) return null;
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch (err) {
    return null;
  }
}
function setHistoryCache_(ss, plate, result) {
  try {
    CacheService.getScriptCache().put(
      historyCacheKey_(ss, plate),
      JSON.stringify(result || {advisor: '', model: '', found: false}),
      21600
    );
  } catch (err) {}
}
function invalidateHistoryCache_(ss, plate) {
  plate = normalizePlate_(plate);
  if (!plate) return;
  try {
    CacheService.getScriptCache().remove(historyCacheKey_(ss, plate));
  } catch (err) {}
}
function latestAdvisorModelFromHistory_(history, plate) {
  const visits = {};
  history.filter(item => item.plate === plate).forEach(item => {
    const key = item.visitId || [item.dateValue, item.minutes, item.plate].join('|');
    if (!visits[key]) visits[key] = {dateValue: item.dateValue, minutes: item.minutes, advisors: [], models: []};
    if (item.advisor && visits[key].advisors.indexOf(item.advisor) < 0) visits[key].advisors.push(item.advisor);
    if (item.model && visits[key].models.indexOf(item.model) < 0) visits[key].models.push(item.model);
  });
  const ordered = Object.keys(visits).map(key => visits[key]).sort((a, b) =>
    b.dateValue - a.dateValue || b.minutes - a.minutes);
  let advisor = '', model = '', advisorResolved = false, modelResolved = false;
  for (let i = 0; i < ordered.length && (!advisorResolved || !modelResolved); i++) {
    if (!advisorResolved && ordered[i].advisors.length) {
      const normalizedAdvisors = uniqueArray_(ordered[i].advisors.map(normalizeKey_));
      advisor = normalizedAdvisors.length === 1 ? ordered[i].advisors[0] : '';
      advisorResolved = true;
    }
    if (!modelResolved && ordered[i].models.length) {
      const normalizedModels = uniqueArray_(ordered[i].models.map(normalizeKey_));
      model = normalizedModels.length === 1 ? ordered[i].models[0] : '';
      modelResolved = true;
    }
  }
  return {advisor: advisor, model: model, found: ordered.length > 0};
}
function invalidateHistoryCacheForEdit_(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet(), name = sh.getName();
  if (name !== CFG.SHEET_DATA_EM && name !== CFG.SHEET_DATA_SCC) return;
  const timeCol = name === CFG.SHEET_DATA_EM ? 9 : CFG.DATA_SCC.GIO_NHAN;
  const idCol = name === CFG.SHEET_DATA_EM ? 15 : CFG.DATA_SCC.VISIT_ID;
  if (![1, 2, 3, 4, timeCol, idCol].some(col => editTouchesColumn_(e.range, col))) return;
  const first = Math.max(2, e.range.getRow());
  const last = e.range.getRow() + e.range.getNumRows() - 1;
  if (last >= first) {
    sh.getRange(first, 2, last - first + 1, 1).getDisplayValues().forEach(value =>
      invalidateHistoryCache_(sh.getParent(), value[0]));
  }
  if (e.range.getNumRows() === 1 && e.range.getNumColumns() === 1 &&
      e.range.getColumn() === 2 && valueExists_(e.oldValue)) {
    invalidateHistoryCache_(sh.getParent(), e.oldValue);
  }
}
function autoFillAdvisorModelFromHistory_(sh, row, historyCache) {
  if (!sh || sh.getName() !== CFG.SHEET_DISPATCH || row < CFG.FIRST_DATA_ROW) return {};
  const advisorCell = sh.getRange(row, CFG.DP.CVDV);
  const modelCell = sh.getRange(row, CFG.DP.MODEL);
  const needsAdvisor = !clean_(advisorCell.getDisplayValue());
  const needsModel = !clean_(modelCell.getDisplayValue());
  if (!needsAdvisor && !needsModel) return {};
  const plate = readPlateText_(sh, row);
  if (!/^[0-9A-Z]{6,12}$/.test(plate)) return {};
  const buildAllowed = cell => {
    const byKey = {};
    layDropdownTuCell_(cell).forEach(value => {
    const key = normalizeKey_(value);
      if (!byKey[key]) byKey[key] = [];
      byKey[key].push(value);
    });
    return byKey;
  };
  const advisorAllowed = needsAdvisor ? buildAllowed(advisorCell) : {};
  const modelAllowed = needsModel ? buildAllowed(modelCell) : {};
  const ss = sh.getParent();
  let result = getHistoryFromCache_(ss, plate);
  if (result === null) {
    const holder = historyCache && !Array.isArray(historyCache) ? historyCache : {};
    if (!Array.isArray(holder.records)) holder.records = loadModelHistory_(ss);
    result = latestAdvisorModelFromHistory_(holder.records, plate);
    setHistoryCache_(ss, plate, result);
  }
  const applied = {};
  if (needsAdvisor && result.advisor) {
    const matches = uniqueArray_(advisorAllowed[normalizeKey_(result.advisor)] || []);
    if (matches.length === 1) { advisorCell.setValue(matches[0]); applied.advisor = matches[0]; }
  }
  if (needsModel && result.model) {
    const matches = uniqueArray_(modelAllowed[normalizeKey_(result.model)] || []);
    if (matches.length === 1) { modelCell.setValue(matches[0]); applied.model = matches[0]; }
  }
  return applied;
}
/*********************************************************
 * XE CHỜ GIAO - GHI TRƯỚC, XÁC MINH, RỒI MỚI DỌN ĐIỀU PHỐI
 *********************************************************/
function waitingDeliverySheet_(ss) {
  const sh = ss && ss.getSheetByName(CFG.SHEET_WAITING_DELIVERY);
  if (!sh) throw new Error('Không tìm thấy sheet ' + CFG.SHEET_WAITING_DELIVERY + '.');
  const expected = ['Ngày', 'Biển số', 'CVDV', 'Model xe', 'Ngày giao', 'Giờ giao', 'Trạng thái', 'visitId'];
  const actual = sh.getRange(5, 1, 1, expected.length).getDisplayValues()[0].map(clean_);
  expected.forEach((header, i) => {
    if (actual[i] !== header) throw new Error(CFG.SHEET_WAITING_DELIVERY + ': header ' + (i + 1) + ' phải là ' + header + '.');
  });
  return sh;
}
function isReadyForWaitingDelivery_(rec, route) {
  return isDispatchReadyToClear_(rec, route) && normalizeKey_(rec.washStatus) === 'hoan thanh';
}
/**
 * V13.20.3-P0.1 - fail-closed P0 gate, chung cho mọi đường chuyển.
 * Kiểm tra TOÀN BỘ bản ghi Data EM / Data SCC có cùng visitId,
 * kể cả khi route của Điều phối không còn nhận diện luồng đó.
 * Không sửa trạng thái, không xóa dòng, không đụng visitId lịch sử.
 */
function assertVisitWorkCompletedBeforeWaiting_(rec, route) {
  const id = clean_(rec && rec.visitId);
  const plate = normalizePlate_(rec && rec.plate);
  if (!id || !plate) {
    throw new Error('P0_WAITING_IDENTITY_INVALID: thiếu visitId hoặc biển số.');
  }
  if (!route || !route.em || !route.scc) {
    throw new Error('P0_WAITING_ROUTE_UNKNOWN: không xác định được luồng công việc.');
  }
  const ss = spreadsheetForRecord_(rec);
  if (!ss) throw new Error('P0_WAITING_SPREADSHEET_MISSING.');
  const tasks = [
    {name: CFG.SHEET_DATA_EM, required: !!route.em.active, label: 'EM', idCol: 15},
    {name: CFG.SHEET_DATA_SCC, required: !!route.scc.active, label: 'SCC', idCol: CFG.DATA_SCC.VISIT_ID}
  ];
  tasks.forEach(function(task) {
    const sh = ss.getSheetByName(task.name);
    if (!sh) throw new Error('P0_WAITING_SHEET_MISSING: ' + task.name);
    const last = sh.getLastRow();
    const hits = [];
    if (last >= 2) {
      const ids = sh.getRange(2, task.idCol, last - 1, 1).getDisplayValues();
      ids.forEach(function(row, idx) {
        if (clean_(row[0]) === id) hits.push(idx + 2);
      });
    }
    if (hits.length > 1) {
      throw new Error('P0_WAITING_DUPLICATE_VISIT: Data ' + task.label + ' có ' + hits.length + ' dòng cùng visitId.');
    }
    if (hits.length === 0) {
      if (task.required) {
        throw new Error('P0_WAITING_REQUIRED_DATA_MISSING: Data ' + task.label + ' chưa có visitId này.');
      }
      return;
    }
    const cell = sh.getRange(hits[0], 2, 1, 12).getDisplayValues()[0];
    // B=plate, M=tiến độ cuối. Chỉ đọc, không ghi.
    const sourcePlate = normalizePlate_(cell[0]);
    const storedStatus = cell[11];
    if (sourcePlate !== plate) {
      throw new Error('P0_WAITING_PLATE_MISMATCH: Data ' + task.label + ' lệch biển số cho visitId.');
    }
    if (!isComplete_(storedStatus)) {
      throw new Error('P0_WAITING_WORK_NOT_DONE: Data ' + task.label +
        ' còn trạng thái "' + clean_(storedStatus) + '"; giữ xe trong Điều phối.');
    }
  });
}
function waitingDeliveryDataLastRow_(sh) {
  const start = CFG.WAITING_DELIVERY_FIRST_ROW;
  const count = sh.getMaxRows() - start + 1;
  if (count <= 0) return start - 1;
  const identity = sh.getRange(start, 1, count, CFG.WD.VISIT_ID).getDisplayValues();
  for (let i = identity.length - 1; i >= 0; i--) {
    if (identity[i].some(value => clean_(value))) return start + i;
  }
  return start - 1;
}
function waitingDeliveryRowsByVisitId_(sh, visitId) {
  const last = waitingDeliveryDataLastRow_(sh);
  if (last < CFG.WAITING_DELIVERY_FIRST_ROW) return [];
  const ids = sh.getRange(CFG.WAITING_DELIVERY_FIRST_ROW, CFG.WD.VISIT_ID,
    last - CFG.WAITING_DELIVERY_FIRST_ROW + 1, 1).getValues();
  const rows = [];
  ids.forEach((value, i) => {
    if (clean_(value[0]) === visitId) rows.push(CFG.WAITING_DELIVERY_FIRST_ROW + i);
  });
  return rows;
}
function waitingDeliveryOpenRowsByPlate_(sh, plate) {
  const normalizedPlate = normalizePlate_(plate);
  if (!normalizedPlate) return [];
  const last = waitingDeliveryDataLastRow_(sh);
  if (last < CFG.WAITING_DELIVERY_FIRST_ROW) return [];
  const count = last - CFG.WAITING_DELIVERY_FIRST_ROW + 1;
  const display = sh
    .getRange(
      CFG.WAITING_DELIVERY_FIRST_ROW,
      1,
      count,
      CFG.WD.VISIT_ID
    )
    .getDisplayValues();
  const rows = [];
  display.forEach(function(values, index) {
    const rowPlate =
      normalizePlate_(
        values[CFG.WD.BIEN_SO - 1]
      );
    const status =
      normalizeKey_(
        values[CFG.WD.TRANG_THAI - 1]
      );
    if (
      rowPlate === normalizedPlate &&
      status === 'cho giao'
    ) {
      rows.push(
        CFG.WAITING_DELIVERY_FIRST_ROW + index
      );
    }
  });
  return rows;
}
function waitingDeliveryRowMatches_(sh, row, rec) {
  const value = sh.getRange(row, 1, 1, CFG.WD.VISIT_ID).getValues()[0];
  const display = sh.getRange(row, 1, 1, CFG.WD.VISIT_ID).getDisplayValues()[0];
  return normalizePlate_(display[CFG.WD.BIEN_SO - 1]) === normalizePlate_(rec.plate) &&
    normalizeKey_(display[CFG.WD.CVDV - 1]) === normalizeKey_(rec.advisor) &&
    normalizeKey_(display[CFG.WD.MODEL - 1]) === normalizeKey_(rec.model) &&
    sameDate_(value[CFG.WD.NGAY - 1], rec.date) &&
    sameDate_(value[CFG.WD.NGAY_GIAO - 1], rec.deliveryDate) &&
    sameTime_(value[CFG.WD.GIO_GIAO - 1], rec.deliveryTime) &&
    ['cho giao', 'hoan thanh'].indexOf(normalizeKey_(display[CFG.WD.TRANG_THAI - 1])) >= 0 &&
    clean_(value[CFG.WD.VISIT_ID - 1]) === rec.visitId;
}
/**
 * Xe chờ giao vẫn có thể giữ dropdown Model để chọn nhanh,
 * nhưng KHÔNG được chặn các dòng xe ngoài danh sách Toyota.
 *
 * Ví dụ thực tế: BMW, Honda, Mazda... vẫn phải chuyển sang
 * "Xe chờ giao" bình thường.
 */
/**
 * Đồng bộ dropdown CVDV + Model của "Xe chờ giao"
 * từ đúng dropdown mẫu ở "Điều phối".
 *
 * Nguồn:
 * - Điều phối!C6 -> Xe chờ giao cột C
 * - Điều phối!D6 -> Xe chờ giao cột D
 *
 * Như vậy sau này chỉ cần chỉnh dropdown ở Điều phối.
 */
function prepareWaitingDeliveryDropdowns_(ss, shWaiting, row) {
  const shDispatch =
    ss.getSheetByName(
      CFG.SHEET_DISPATCH
    );
  if (!shDispatch) {
    throw new Error(
      'Không tìm thấy sheet "' +
      CFG.SHEET_DISPATCH +
      '" để đồng bộ dropdown Xe chờ giao.'
    );
  }
  const sourceCvdvRule =
    shDispatch
      .getRange(
        CFG.FIRST_DATA_ROW,
        CFG.DP.CVDV
      )
      .getDataValidation();
  const sourceModelRule =
    shDispatch
      .getRange(
        CFG.FIRST_DATA_ROW,
        CFG.DP.MODEL
      )
      .getDataValidation();
  const targetCvdv =
    shWaiting.getRange(
      row,
      CFG.WD.CVDV
    );
  const targetModel =
    shWaiting.getRange(
      row,
      CFG.WD.MODEL
    );
  if (sourceCvdvRule) {
    targetCvdv.setDataValidation(
      sourceCvdvRule.copy().build()
    );
  }
  else {
    targetCvdv.clearDataValidations();
  }
  if (sourceModelRule) {
    targetModel.setDataValidation(
      sourceModelRule.copy().build()
    );
  }
  else {
    targetModel.clearDataValidations();
  }
}
/**
 * Đồng bộ dropdown cho toàn bộ vùng Xe chờ giao.
 * Hàm này chạy khi mở file và có thể chạy thủ công từ menu.
 */
function dongBoDropdownXeChoGiao() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error(
      'Không xác định được Spreadsheet hiện tại.'
    );
  }
  const shWaiting =
    ss.getSheetByName(
      CFG.SHEET_WAITING_DELIVERY
    );
  if (!shWaiting) {
    throw new Error(
      'Không tìm thấy sheet "' +
      CFG.SHEET_WAITING_DELIVERY +
      '".'
    );
  }
  const shDispatch =
    ss.getSheetByName(
      CFG.SHEET_DISPATCH
    );
  if (!shDispatch) {
    throw new Error(
      'Không tìm thấy sheet "' +
      CFG.SHEET_DISPATCH +
      '".'
    );
  }
  const sourceCvdvRule =
    shDispatch
      .getRange(
        CFG.FIRST_DATA_ROW,
        CFG.DP.CVDV
      )
      .getDataValidation();
  const sourceModelRule =
    shDispatch
      .getRange(
        CFG.FIRST_DATA_ROW,
        CFG.DP.MODEL
      )
      .getDataValidation();
  const rowCount =
    Math.max(
      shWaiting.getMaxRows() -
        CFG.WAITING_DELIVERY_FIRST_ROW +
        1,
      1
    );
  const cvdvRange =
    shWaiting.getRange(
      CFG.WAITING_DELIVERY_FIRST_ROW,
      CFG.WD.CVDV,
      rowCount,
      1
    );
  const modelRange =
    shWaiting.getRange(
      CFG.WAITING_DELIVERY_FIRST_ROW,
      CFG.WD.MODEL,
      rowCount,
      1
    );
  if (sourceCvdvRule) {
    cvdvRange.setDataValidation(
      sourceCvdvRule.copy().build()
    );
  }
  else {
    cvdvRange.clearDataValidations();
  }
  if (sourceModelRule) {
    modelRange.setDataValidation(
      sourceModelRule.copy().build()
    );
  }
  else {
    modelRange.clearDataValidations();
  }
  SpreadsheetApp.flush();
  try {
    ss.toast(
      'CVDV và Model của Xe chờ giao đã lấy dropdown từ Điều phối.',
      '✅ ĐỒNG BỘ DROPDOWN',
      5
    );
  }
  catch (err) {
    // Không cần xử lý khi chạy ngoài UI.
  }
}
function waitingAdvisorKey_(value) {
  return normalizeKey_(value).replace(/^cv\s+/, '');
}
function isHondaWaitingAdvisor_(value) {
  const key = waitingAdvisorKey_(value);
  return WAITING_CANONICAL_CFG.HONDA_ADVISORS.indexOf(key) >= 0;
}
function waitingDateKeyToDate_(dateKey) {
  const match = String(dateKey || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  // 12:00 tránh lệch ngày do timezone/DST khi format lại trong Sheet.
  const value = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12, 0, 0, 0);
  return isNaN(value.getTime()) ? null : value;
}
function waitingDateKeyDiffDays_(olderKey, newerKey) {
  const a = String(olderKey || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const b = String(newerKey || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!a || !b) return NaN;
  const ams = Date.UTC(Number(a[1]), Number(a[2]) - 1, Number(a[3]));
  const bms = Date.UTC(Number(b[1]), Number(b[2]) - 1, Number(b[3]));
  return Math.round((bms - ams) / 86400000);
}
function canonicalWaitingDateFromDataCvdv_(ss, rec) {
  if (!ss || !rec || isHondaWaitingAdvisor_(rec.advisor)) {
    return rec && rec.date;
  }
  if (!(rec.date instanceof Date) || isNaN(rec.date.getTime())) {
    return rec.date;
  }
  const targetDateKey = cvdvDateKey_(rec.date, ss);
  const targetPlate = normalizePlate_(rec.plate);
  const targetAdvisor = waitingAdvisorKey_(rec.advisor);
  if (!targetDateKey || !targetPlate || !targetAdvisor) return rec.date;

  const source = ss.getSheetByName(CFG.SHEET_CVDV_DATA);
  if (!source) return rec.date;
  const first = CFG.CVDV_DATA_FIRST_ROW;
  const last = source.getLastRow();
  if (last < first) return rec.date;

  const width = Math.max(
    CFG.CVDV_DATA.NGAY,
    CFG.CVDV_DATA.BIEN_SO,
    CFG.CVDV_DATA.TINH_TRANG,
    CFG.CVDV_DATA.CVDV
  );
  const rowCount = last - first + 1;
  const values = source.getRange(first, 1, rowCount, width).getValues();
  const display = source.getRange(first, 1, rowCount, width).getDisplayValues();
  const groups = {};

  for (let i = 0; i < rowCount; i++) {
    if (normalizePlate_(display[i][CFG.CVDV_DATA.BIEN_SO - 1]) !== targetPlate) continue;
    if (waitingAdvisorKey_(display[i][CFG.CVDV_DATA.CVDV - 1]) !== targetAdvisor) continue;

    const sourceDate = values[i][CFG.CVDV_DATA.NGAY - 1];
    const sourceDateKey = cvdvDateKey_(sourceDate, ss);
    if (!sourceDateKey || sourceDateKey > targetDateKey) continue;
    const diffDays = waitingDateKeyDiffDays_(sourceDateKey, targetDateKey);
    if (
      !isFinite(diffDays) ||
      diffDays < 0 ||
      diffDays > WAITING_CANONICAL_CFG.CANONICAL_LOOKBACK_DAYS
    ) continue;

    if (!groups[sourceDateKey]) {
      groups[sourceDateKey] = {
        dateKey: sourceDateKey,
        date: sourceDate instanceof Date && !isNaN(sourceDate.getTime())
          ? new Date(sourceDate.getTime())
          : waitingDateKeyToDate_(sourceDateKey),
        hasInShop: false
      };
    }
    if (normalizeKey_(display[i][CFG.CVDV_DATA.TINH_TRANG - 1]) === 'trong xuong') {
      groups[sourceDateKey].hasInShop = true;
    }
  }
  const dateKeys = Object.keys(groups).sort();
  if (!dateKeys.length) return rec.date;
  // Chỉ map về ngày CVDV nếu NGÀY MỚI NHẤT cùng BSX + cùng CVDV vẫn Trong xưởng.
  // Nếu ngày mới nhất đã Đã ra cổng, không quay ngược về một record Trong xưởng cũ hơn.
  const latest = groups[dateKeys[dateKeys.length - 1]];
  return latest && latest.hasInShop && latest.date ? latest.date : rec.date;
}
function canonicalizeWaitingDeliveryRecord_(rec) {
  if (!rec) return rec;
  const ss = spreadsheetForRecord_(rec);
  const canonicalDate = canonicalWaitingDateFromDataCvdv_(ss, rec);
  if (
    !(canonicalDate instanceof Date) ||
    isNaN(canonicalDate.getTime()) ||
    sameDate_(canonicalDate, rec.date)
  ) {
    return rec;
  }
  const copy = Object.assign({}, rec);
  copy.date = canonicalDate;
  return copy;
}
function cvdvEvidenceGroupMatchesAdvisor_(group, advisor) {
  const target = waitingAdvisorKey_(advisor);
  if (!target || !group || !group.records || !group.records.length) return false;
  const advisorKeys = uniqueArray_(
    group.records
      .map(function(rec) { return waitingAdvisorKey_(rec.advisor); })
      .filter(function(key) { return !!key; })
  );
  return advisorKeys.length === 1 && advisorKeys[0] === target;
}
function repairWaitingDatesFromCvdv_(waiting, evidence, ss) {
  const out = {checked: 0, repaired: 0, blocked: 0};
  const start = CFG.WAITING_DELIVERY_FIRST_ROW;
  const last = waitingDeliveryDataLastRow_(waiting);
  if (last < start) return out;

  const count = last - start + 1;
  const range = waiting.getRange(start, 1, count, CFG.WD.VISIT_ID);
  const values = range.getValues();
  const display = range.getDisplayValues();
  const visitCount = {};
  values.forEach(function(row) {
    const id = clean_(row[CFG.WD.VISIT_ID - 1]);
    if (id) visitCount[id] = (visitCount[id] || 0) + 1;
  });

  for (let i = 0; i < count; i++) {
    if (normalizeKey_(display[i][CFG.WD.TRANG_THAI - 1]) !== 'cho giao') continue;
    if (isHondaWaitingAdvisor_(display[i][CFG.WD.CVDV - 1])) continue;
    const visitId = clean_(values[i][CFG.WD.VISIT_ID - 1]);
    if (!visitId || visitCount[visitId] !== 1) {
      out.blocked++;
      continue;
    }
    const waitingDateKey = cvdvDateKey_(values[i][CFG.WD.NGAY - 1], ss);
    const plate = normalizePlate_(display[i][CFG.WD.BIEN_SO - 1]);
    const advisor = clean_(display[i][CFG.WD.CVDV - 1]);
    if (!waitingDateKey || !plate || !advisor) continue;
    out.checked++;

    const exact = evidence && evidence.exact
      ? evidence.exact[waitingDateKey + '|' + plate]
      : null;
    if (exact) continue;

    const plateIndex = evidence && evidence.byPlateDates
      ? evidence.byPlateDates[plate]
      : null;
    if (!plateIndex || !plateIndex.dateKeys || !plateIndex.dateKeys.length) continue;

    const candidateKeys = plateIndex.dateKeys
      .filter(function(dateKey) {
        if (dateKey >= waitingDateKey) return false;
        const diff = waitingDateKeyDiffDays_(dateKey, waitingDateKey);
        return isFinite(diff) && diff >= 0 && diff <= WAITING_CANONICAL_CFG.CANONICAL_LOOKBACK_DAYS;
      })
      .filter(function(dateKey) {
        return cvdvEvidenceGroupMatchesAdvisor_(plateIndex.groups[dateKey], advisor);
      });
    if (!candidateKeys.length) continue;

    // Chỉ xét NGÀY MỚI NHẤT cùng BSX + cùng CVDV trong cửa sổ.
    // Nếu ngày mới nhất đã Đã ra cổng thì giữ BLOCK, không lùi về record Trong xưởng cũ.
    const canonicalKey = candidateKeys[candidateKeys.length - 1];
    const canonicalGroup = plateIndex.groups[canonicalKey];
    if (!canonicalGroup || !canonicalGroup.hasInShop) continue;
    const canonicalDate = waitingDateKeyToDate_(canonicalKey);
    if (!canonicalDate) continue;

    const row = start + i;
    const liveValues = waiting.getRange(row, 1, 1, CFG.WD.VISIT_ID).getValues()[0];
    const liveDisplay = waiting.getRange(row, 1, 1, CFG.WD.VISIT_ID).getDisplayValues()[0];
    if (
      clean_(liveValues[CFG.WD.VISIT_ID - 1]) !== visitId ||
      normalizeKey_(liveDisplay[CFG.WD.TRANG_THAI - 1]) !== 'cho giao' ||
      normalizePlate_(liveDisplay[CFG.WD.BIEN_SO - 1]) !== plate ||
      waitingAdvisorKey_(liveDisplay[CFG.WD.CVDV - 1]) !== waitingAdvisorKey_(advisor) ||
      cvdvDateKey_(liveValues[CFG.WD.NGAY - 1], ss) !== waitingDateKey
    ) {
      out.blocked++;
      continue;
    }
    waiting.getRange(row, CFG.WD.NGAY).setValue(canonicalDate).setNumberFormat('dd/MM/yyyy');
    out.repaired++;
  }
  if (out.repaired) SpreadsheetApp.flush();
  return out;
}
function autoClearHondaWaiting_(waiting, ss) {
  // TPMH 12/11/2026: TAT TU THU HOI HONDA THEO NGAY / SAU 15:00.
  // Xe Honda (CV Lap / CV Tien) CHI ra khoi Cho giao khi co bang chung
  // DA RA CONG dung luot qua luong CVDV/Le tan/Camera hien hanh, hoac
  // duoc quan tri xac minh va dong luot an toan.
  // "Hoan thanh EM", ngay tiep nhan va gio 15:00 KHONG PHAI chung cu ra cong.
  // Keep legacy return shape for dongBoTrangThaiCVDV_ and existing logs.
  // Intentionally NO read/write to waiting; NEVER compact any waiting rows.
  return {
    checked: 0,
    eligible: 0,
    completed: 0,
    blockedMissingVisitId: 0,
    blockedDuplicateVisitId: 0
  };
}

/*********************************************************
 * P0.3 - CHẶN HAI DÒNG ĐIỀU PHỐI CÙNG BIỂN SỐ KHÁC visitId
 * Read only, kiểm tra trước khi append/update Chờ giao.
 *********************************************************/
function assertNoOtherActiveDispatchPlateP03_(ss, rec) {
  const sh = ss.getSheetByName(CFG.SHEET_DISPATCH);
  if (!sh) throw new Error('P03_DISPATCH_SHEET_MISSING');
  const plate = normalizePlate_(rec && rec.plate);
  const id = clean_(rec && rec.visitId);
  if (!plate || !id) throw new Error('P03_DISPATCH_IDENTITY_INVALID');
  const start = CFG.FIRST_DATA_ROW;
  const last = sh.getLastRow();
  if (last < start) return true;
  // A:V chứ không dùng toàn 24 cột; chỉ đọc, không ghi.
  const values = sh.getRange(start, 1, last - start + 1, CFG.DP.VISIT_ID).getDisplayValues();
  const matches = [];
  values.forEach(function(r,i) {
    if (normalizePlate_(r[CFG.DP.BIEN_SO-1]) === plate) {
      matches.push({row:start+i, visitId:clean_(r[CFG.DP.VISIT_ID-1])});
    }
  });
  if (matches.length > 1 || matches.some(function(m) {return m.visitId !== id;})) {
    throw new Error('P03_MULTIPLE_ACTIVE_VISITS: biển số ' + plate +
      ' xuất hiện ở ' + matches.map(function(m) {return m.row + ':' + m.visitId;}).join(', ') +
      '. Không tự chuyển sang Xe chờ giao.');
  }
  return true;
}

function appendWaitingDelivery_(rec) {
  rec = canonicalizeWaitingDeliveryRecord_(rec);
  const sh = waitingDeliverySheet_(spreadsheetForRecord_(rec));
  const visitId = clean_(rec.visitId);
  if (!visitId) {
    throw new Error(
      'Thiếu visitId; không thể chuyển sang Xe chờ giao.'
    );
  }
  // P0.3: không chuyển xe nếu 2 lượt Điều phối cùng biển số
  // nhưng khác visitId; không thể đoán lượt nào là lần ra/vào.
  assertNoOtherActiveDispatchPlateP03_(spreadsheetForRecord_(rec), rec);
  // P0: chốt ngay trước mọi thao tác append/update dữ liệu chờ giao.
  // Chặn cả Data EM/SCC ngoài route khi cùng visitId chưa Hoàn thành.
  assertVisitWorkCompletedBeforeWaiting_(rec, classifyWork_(rec));
  const existing =
    waitingDeliveryRowsByVisitId_(
      sh,
      visitId
    );
  if (existing.length > 1) {
    throw new Error(
      'Xe chờ giao có duplicate visitId ' +
      visitId +
      '.'
    );
  }
  /*********************************************************
   * ĐÃ CÓ CÙNG visitId
   *
   * Trường hợp thường gặp:
   * - xe đã từng được chuyển sang Xe chờ giao;
   * - sau đó CVDV / Model / Ngày giao / Giờ giao ở Điều phối
   *   được chỉnh lại;
   * - code cũ xem đó là "mâu thuẫn" và chặn.
   *
   * Với cùng visitId, đây vẫn là CÙNG MỘT LƯỢT XE.
   * Vì vậy:
   * - nếu biển số vẫn cùng xe -> cập nhật dữ liệu mới;
   * - nếu biển số khác -> giữ chặn vì có khả năng visitId lỗi.
   *********************************************************/
  if (existing.length === 1) {
    const row =
      existing[0];
    const currentDisplay =
      sh
        .getRange(
          row,
          1,
          1,
          CFG.WD.VISIT_ID
        )
        .getDisplayValues()[0];
    const currentPlate =
      normalizePlate_(
        currentDisplay[
          CFG.WD.BIEN_SO - 1
        ]
      );
    const incomingPlate =
      normalizePlate_(
        rec.plate
      );
    if (
      currentPlate &&
      incomingPlate &&
      currentPlate !== incomingPlate
    ) {
      throw new Error(
        'Xe chờ giao đã có visitId ' +
        visitId +
        ' nhưng biển số khác (' +
        currentDisplay[
          CFG.WD.BIEN_SO - 1
        ] +
        ' ≠ ' +
        rec.plate +
        '); giữ nguyên Điều phối.'
      );
    }
    const currentStatus =
      normalizeKey_(
        currentDisplay[
          CFG.WD.TRANG_THAI - 1
        ]
      );
    const keepStatus =
      currentStatus === 'hoan thanh'
        ? 'Hoàn thành'
        : 'Chờ giao';
    const values = [[
      rec.date,
      plateAsText_(
        rec.plate
      ),
      rec.advisor,
      rec.model,
      rec.deliveryDate,
      rec.deliveryTime,
      keepStatus,
      visitId
    ]];
    preparePlateCell_(
      sh,
      row
    );
    prepareWaitingDeliveryDropdowns_(
      spreadsheetForRecord_(rec),
      sh,
      row
    );
    sh
      .getRange(
        row,
        1,
        1,
        CFG.WD.VISIT_ID
      )
      .setValues(
        values
      );
    sh
      .getRange(
        row,
        CFG.WD.GIO_GIAO
      )
      .setNumberFormat(
        'HH:mm'
      );
    sh
      .getRange(
        row,
        CFG.WD.VISIT_ID
      )
      .setNumberFormat(
        '@'
      );
    SpreadsheetApp.flush();
    if (
      !waitingDeliveryRowMatches_(
        sh,
        row,
        rec
      )
    ) {
      throw new Error(
        'Đã cập nhật Xe chờ giao nhưng không xác minh được dữ liệu cho visitId ' +
        visitId +
        '.'
      );
    }
    return row;
  }
  /*********************************************************
   * CHƯA CÓ visitId NHƯNG BIỂN SỐ ĐANG "CHỜ GIAO"
   *
   * Nghiệp vụ:
   * - Một xe vật lý đang ở đại lý chỉ được tính 1 lượt Xe chờ giao.
   * - Nếu phát sinh RO/visitId mới trước khi xe ra cổng,
   *   cập nhật chính dòng Chờ giao hiện tại thay vì append thêm.
   * - Nếu lượt cũ đã Hoàn thành/Đã ra cổng thì dòng cũ đã được
   *   compact; lúc đó visit mới sẽ tạo dòng mới bình thường.
   *********************************************************/
  const openRowsSamePlate =
    waitingDeliveryOpenRowsByPlate_(
      sh,
      rec.plate
    );
  if (openRowsSamePlate.length > 1) {
    throw new Error(
      'Xe chờ giao đang có ' +
      openRowsSamePlate.length +
      ' dòng Chờ giao cùng biển số ' +
      rec.plate +
      '. Hãy chạy gomTrungXeChoGiaoDangCho() một lần rồi thử lại.'
    );
  }
  if (openRowsSamePlate.length === 1) {
    const occupiedRow = openRowsSamePlate[0];
    const previousVisitId = clean_(sh.getRange(occupiedRow, CFG.WD.VISIT_ID).getValue());
    // P0.3: Tuyệt đối không đổi owner một lượt đã công bố Chờ giao.
    // Hai visitId cùng biển số có thể là nhiều RO của một lần nằm xưởng
    // hoặc lần tái nhập xưởng. Không có chứng cứ ra cổng, không tự đoán.
    throw new Error(
      'P03_PHYSICAL_VISIT_CONFLICT: Xe ' + normalizePlate_(rec.plate) +
      ' còn Chờ giao ở dòng ' + occupiedRow +
      ' (visitId=' + previousVisitId +
      '), nhưng đang chuyển visitId=' + visitId +
      '. Không ghi đè lịch sử/visitId. Kiểm tra một lượt nằm xưởng trước khi xử lý.'
    );
  }
  /*********************************************************
   * CHƯA CÓ visitId -> TẠO DÒNG MỚI
   *********************************************************/
  const row =
    Math.max(
      waitingDeliveryDataLastRow_(
        sh
      ) + 1,
      CFG.WAITING_DELIVERY_FIRST_ROW
    );
  const values = [[
    rec.date,
    plateAsText_(
      rec.plate
    ),
    rec.advisor,
    rec.model,
    rec.deliveryDate,
    rec.deliveryTime,
    'Chờ giao',
    visitId
  ]];
  try {
    preparePlateCell_(
      sh,
      row
    );
    prepareWaitingDeliveryDropdowns_(
      spreadsheetForRecord_(rec),
      sh,
      row
    );
    sh
      .getRange(
        row,
        1,
        1,
        CFG.WD.VISIT_ID
      )
      .setValues(
        values
      );
    sh
      .getRange(
        row,
        CFG.WD.GIO_GIAO
      )
      .setNumberFormat(
        'HH:mm'
      );
    sh
      .getRange(
        row,
        CFG.WD.VISIT_ID
      )
      .setNumberFormat(
        '@'
      );
    SpreadsheetApp.flush();
    if (
      !waitingDeliveryRowMatches_(
        sh,
        row,
        rec
      )
    ) {
      throw new Error(
        'Không xác minh được dữ liệu vừa ghi vào Xe chờ giao.'
      );
    }
  }
  catch (err) {
    const writtenId =
      clean_(
        sh
          .getRange(
            row,
            CFG.WD.VISIT_ID
          )
          .getValue()
      );
    if (
      !writtenId ||
      writtenId === visitId
    ) {
      sh
        .getRange(
          row,
          1,
          1,
          CFG.WD.VISIT_ID
        )
        .clearContent()
        .clearNote();
    }
    throw new Error(
      'Ghi Xe chờ giao thất bại; Điều phối được giữ nguyên: ' +
      err.message
    );
  }
  // Lượt mới bình thường chưa đổi ownership: cho phép CVDV auto-sync.
  clearWaitingVisitCvdvUnsafe_(
    visitId
  );
  return row;
}
function moveDispatchToWaitingDelivery_(sh, row, rec) {
  const visitId = clean_(rec && rec.visitId);
  if (!visitId) {
    throw new Error(
      'Không thể chuyển Xe chờ giao khi thiếu visitId.'
    );
  }
  const ownerRows =
    resolveDispatchRowsByVisitId_(
      sh,
      visitId
    );
  if (
    ownerRows.length !== 1 ||
    ownerRows[0] !== row
  ) {
    throw new Error(
      'Không thể chuyển Xe chờ giao: ownership Điều phối không duy nhất.'
    );
  }
  // P0: re-read sau các bước sync, ngay trước khi chuyển chờ giao.
  SpreadsheetApp.flush();
  const current = readDispatchRow_(sh, row);
  if (clean_(current.visitId) !== visitId ||
      normalizePlate_(current.plate) !== normalizePlate_(rec.plate)) {
    throw new Error('P0_WAITING_OWNER_CHANGED: visitId/biển số vừa thay đổi.');
  }
  const route = classifyWork_(current);
  if (!isReadyForWaitingDelivery_(current, route)) {
    throw new Error('P0_WAITING_ROUTE_NOT_READY: chưa đủ điều kiện chuyển chờ giao.');
  }
  const missing = validateStoredDataBeforeClear_(current, route);
  if (missing.length) {
    throw new Error('P0_WAITING_SOURCE_INVALID: ' + missing.join(' | '));
  }
  // appendWaitingDelivery_ tự kiểm tra ALL Data EM/SCC theo visitId.
  appendWaitingDelivery_(current);
  // Xác minh lại ownership trước compact.
  const ownerRowsAfterAppend =
    resolveDispatchRowsByVisitId_(
      sh,
      visitId
    );
  if (
    ownerRowsAfterAppend.length !== 1 ||
    ownerRowsAfterAppend[0] !== row
  ) {
    throw new Error(
      'Ownership Điều phối thay đổi sau khi ghi Xe chờ giao; không compact.'
    );
  }
  clearAndCompactDispatchRow_(
    sh,
    row,
    visitId
  );
  return true;
}
function waitingDeliveryCompactSnapshot_(range) {
  return {
    values: range.getValues(),
    notes: range.getNotes(),
    validations: range.getDataValidations()
  };
}
function restoreWaitingDeliveryCompactSnapshot_(range, snapshot) {
  // Tháo validation trước khi restore values vì reject-input có thể chặn chính snapshot cũ.
  range.clearDataValidations();
  range.setValues(snapshot.values);
  range.setNotes(snapshot.notes);
  range.setDataValidations(snapshot.validations);
}
function compactWaitingDeliverySnapshot_(snapshot, offset, width) {
  const values = snapshot.values.map(function(r) { return r.slice(); });
  const notes = snapshot.notes.map(function(r) { return r.slice(); });
  const validations = snapshot.validations.map(function(r) { return r.slice(); });
  values.splice(offset, 1);
  notes.splice(offset, 1);
  validations.splice(offset, 1);
  values.push(Array(width).fill(''));
  notes.push(Array(width).fill(''));
  validations.push(Array(width).fill(null));
  return {values: values, notes: notes, validations: validations};
}
function applyWaitingDeliveryCompact_(range, compacted) {
  // Validation reject-input (đặc biệt Model/CVDV legacy) không được chặn compact.
  range.clearDataValidations();
  range.setValues(compacted.values);
  range.setNotes(compacted.notes);
  range.setDataValidations(compacted.validations);
}
function clearAndCompactWaitingDeliveryRow_(sh, row, expectedVisitId) {
  const start = CFG.WAITING_DELIVERY_FIRST_ROW;
  const last = waitingDeliveryDataLastRow_(sh);
  const width = CFG.WD.VISIT_ID;
  if (row < start || row > last) return false;
  if (expectedVisitId) {
    const currentId = clean_(sh.getRange(row, CFG.WD.VISIT_ID).getValue());
    if (currentId !== clean_(expectedVisitId)) {
      throw new Error('Không compact Xe chờ giao: ownership row/visitId đã thay đổi.');
    }
    const ownerRows = waitingDeliveryRowsByVisitId_(sh, expectedVisitId);
    if (ownerRows.length !== 1 || ownerRows[0] !== row) {
      throw new Error('Không compact Xe chờ giao: visitId không duy nhất hoặc không còn ở dòng dự kiến.');
    }
  }
  const range = sh.getRange(start, 1, last - start + 1, width);
  const snapshot = waitingDeliveryCompactSnapshot_(range);
  try {
    const compacted = compactWaitingDeliverySnapshot_(snapshot, row - start, width);
    applyWaitingDeliveryCompact_(range, compacted);
    SpreadsheetApp.flush();
    const count = sh.getMaxRows() - start + 1;
    if (count > 0) sh.getRange(start, CFG.WD.BIEN_SO, count, 1).setNumberFormat('@');
    if (expectedVisitId) clearWaitingVisitCvdvUnsafe_(expectedVisitId);
    return true;
  } catch (err) {
    try {
      restoreWaitingDeliveryCompactSnapshot_(range, snapshot);
      SpreadsheetApp.flush();
    } catch (restoreErr) {
      throw new Error(
        'Compact Xe chờ giao lỗi: ' + err.message +
        ' | Khôi phục snapshot lỗi: ' + restoreErr.message
      );
    }
    throw new Error(
      'Compact Xe chờ giao lỗi; đã khôi phục snapshot dữ liệu + validation: ' + err.message
    );
  }
}



function handleWaitingDeliveryEdit_(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== CFG.SHEET_WAITING_DELIVERY ||
      !editTouchesColumn_(e.range, CFG.WD.TRANG_THAI)) return;
  waitingDeliverySheet_(sh.getParent());
  const first = Math.max(e.range.getRow(), CFG.WAITING_DELIVERY_FIRST_ROW);
  const last = e.range.getRow() + e.range.getNumRows() - 1;
  for (let row = last; row >= first; row--) {
    if (isComplete_(sh.getRange(row, CFG.WD.TRANG_THAI).getDisplayValue())) {
      const visitId = clean_(sh.getRange(row, CFG.WD.VISIT_ID).getValue());
      if (!visitId) {
        throw new Error(
          'Xe chờ giao thiếu visitId tại dòng ' + row + '.'
        );
      }
      clearAndCompactWaitingDeliveryRow_(
        sh,
        row,
        visitId
      );
      clearWaitingVisitCvdvUnsafe_(
        visitId
      );
    }
  }
}
/*********************************************************
 * ĐỒNG BỘ XE ĐÃ RA CỔNG TỪ DATA_CVDV
 *
 * V13.12.8 - RECONCILIATION AN TOÀN
 *
 * Nguồn:
 * - Data_CVDV!A:A = Ngày tiếp nhận
 * - Data_CVDV!C:C = Biển số
 * - Data_CVDV!E:E = Tình trạng
 *
 * Nguyên tắc P0 SAFE V13.15:
 * 1) Chỉ dùng EXACT Ngày tiếp nhận + BSX của chính lượt Xe chờ giao.
 * 2) EXACT phải có >=1 record và TẤT CẢ đều = "Đã ra cổng".
 * 3) Có "Trong xưởng" / trạng thái trống-mâu thuẫn / không có EXACT => GIỮ.
 * 4) Record ngày sau KHÔNG BAO GIỜ được dùng để auto-clear lượt hiện tại.
 * 5) visitId phải có, duy nhất, trạng thái hiện tại phải là "Chờ giao".
 * 6) WAITING_CVDV_GUARD mặc định BLOCK; chỉ được recovery nếu cùng thời điểm thỏa EXACT Ngày + BSX = toàn bộ Đã ra cổng, visitId duy nhất và trạng thái vẫn Chờ giao.
 * 7) Guard được re-check + evidence được re-check NGAY TRƯỚC bulk remove; guard mới phát sinh sau audit luôn BLOCK.
 * 8) Compact theo snapshot một lần để giảm lock/WebApp contention.
 *********************************************************/
function dongBoTrangThaiCVDV() {
  const p052Started = Date.now();
  const ss = laySpreadsheetWeb_();
  assertVisitSetup_(ss);
  const source = ss.getSheetByName(CFG.SHEET_CVDV_DATA);
  if (!source) {
    throw new Error('Không tìm thấy sheet "' + CFG.SHEET_CVDV_DATA + '".');
  }
  // Đọc Data_CVDV ngoài lock vì đây là phần có thể chậm.
  const evidence = loadLatestCVDVStatusByVisit_(source, ss);
  const p052EvidenceMs = Date.now() - p052Started;
  if (isWebAppConfirmPriorityActive_()) {
    scheduleCVDVSyncRetry_();
    return {
      success: true,
      busy: true,
      webAppPriority: true,
      queued: true,
      scanned: 0,
      matched: 0,
      completed: 0
    };
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(2500)) {
    scheduleCVDVSyncRetry_();
    try {
      ss.toast(
        'Hệ thống đang xử lý lượt xe khác. Sẽ thử lại ở lượt đồng bộ kế tiếp.',
        '⏳ ĐỒNG BỘ CVDV ĐANG CHỜ',
        8
      );
    } catch (err) {}
    return {
      success: true,
      busy: true,
      queued: true,
      scanned: 0,
      matched: 0,
      completed: 0
    };
  }
  try {
    if (isWebAppConfirmPriorityActive_()) {
      scheduleCVDVSyncRetry_();
      return {
        success: true,
        busy: true,
        webAppPriority: true,
        queued: true,
        scanned: 0,
        matched: 0,
        completed: 0,
        yieldedAfterLock: true
      };
    }
    // V13.18: trigger 5 phút chỉ làm nhiệm vụ P0-safe Xe chờ giao.
    // Advisor reconciliation đã chuyển sang:
    // 1) event LSC/WebApp ngay lúc xác nhận;
    // 2) watchdog 3 chiều bên CVDV mỗi 15 phút.
    const result = dongBoTrangThaiCVDV_(ss, evidence, {
      mode: 'auto'
    });
    result.crossCheck = {
      skipped: true,
      code: 'EVENT_DRIVEN_PLUS_CVDV_WATCHDOG'
    };
    // Đo thực tế, không tự báo tốc độ nếu chưa có log Production.
    result.p052Timing = {evidenceMs: p052EvidenceMs,
      processingMs: Date.now() - p052Started - p052EvidenceMs,
      partialTotalMs: Date.now() - p052Started};
    try {
      const msg =
        result.completed > 0
          ? 'Đã dọn ' + result.completed + ' xe: EXACT CVDV=' +
            Number(result.exactCompleted || 0) + ', Honda auto-disabled=' +
            Number(result.hondaCompleted || 0) +
            (Number(result.dateRepaired || 0) > 0
              ? ', sửa ngày tiếp nhận=' + Number(result.dateRepaired || 0)
              : '') + '.'
          : (Number(result.dateRepaired || 0) > 0
              ? 'Đã sửa ngày tiếp nhận cho ' + Number(result.dateRepaired || 0) + ' xe; chưa có xe đủ điều kiện dọn.'
              : 'Chưa có xe đủ điều kiện tự động dọn.');
      ss.toast(
        msg,
        '🚪 ĐỒNG BỘ CVDV',
        7
      );
    } catch (err) {}
    return result;
  } finally {
    const p052FlushStarted = Date.now();
    try {
      SpreadsheetApp.flush();
    } finally {
      // Ghi cả thời gian flush/giải phóng sau khi xử lý, để log phản ánh
      // gần đúng thời lượng toàn execution thay vì chỉ phần tính toán.
      console.log('[P052_PERF_CVDV] ' + JSON.stringify({
        evidenceMs: p052EvidenceMs,
        afterEvidenceMs: p052FlushStarted - p052Started - p052EvidenceMs,
        finalFlushMs: Date.now() - p052FlushStarted,
        totalMs: Date.now() - p052Started
      }));
      lock.releaseLock();
    }
  }
}
function scheduleCVDVSyncRetry_() {
  const handler = 'dongBoTrangThaiCVDV';
  const exists =
    ScriptApp
      .getProjectTriggers()
      .some(function(trigger) {
        return (
          trigger.getHandlerFunction() === handler &&
          trigger.getEventType() === ScriptApp.EventType.CLOCK &&
          trigger.getTriggerSourceId() === null
        );
      });
  if (exists) return false;
  ScriptApp
    .newTrigger(handler)
    .timeBased()
    .after(60 * 1000)
    .create();
  return true;
}
function dongBoTrangThaiCVDV_(ss, evidence, options) {
  if (!ss) throw new Error('Không xác định được Spreadsheet.');
  options = options || {};
  if (
    !evidence ||
    !evidence.exact ||
    !evidence.byPlateDates
  ) {
    const source = ss.getSheetByName(CFG.SHEET_CVDV_DATA);
    if (!source) {
      throw new Error('Không tìm thấy sheet "' + CFG.SHEET_CVDV_DATA + '".');
    }
    evidence =
      loadLatestCVDVStatusByVisit_(
        source,
        ss
      );
  }
  const waiting = waitingDeliverySheet_(ss);
  // Honda: không tự dọn theo ngày/giờ; giữ tương thích thống kê legacy.
  const honda = autoClearHondaWaiting_(waiting, ss);
  // Toyota: sửa NGÀY CHỜ về ngày tiếp nhận CVDV chỉ khi cùng BSX + cùng CVDV
  // và nguồn vẫn "Trong xưởng". Không sửa theo nguồn "Đã ra cổng" ngày lệch.
  const dateRepair = repairWaitingDatesFromCvdv_(waiting, evidence, ss);
  const audit =
    collectWaitingCvdvAudit_(
      waiting,
      evidence,
      ss
    );
  if (!audit.rows.length) {
    return {
      success: true,
      scanned: honda.checked,
      matched: 0,
      completed: honda.completed,
      exactCompleted: 0,
      hondaCompleted: honda.completed,
      hondaEligible: honda.eligible,
      dateRepaired: dateRepair.repaired,
      kept: 0
    };
  }
  /*********************************************************
   * V13.12.11
   *
   * AUTO và MANUAL cùng dùng 1 nguyên tắc vật lý:
   * - Chỉ xét record Data_CVDV có ngày >= Ngày sửa chữa.
   * - Chỉ xét tối đa 60 ngày kể từ Ngày sửa chữa.
   * - Trong cửa sổ đó, NGÀY MỚI NHẤT quyết định:
   *   + toàn bộ Đã ra cổng -> dọn;
   *   + có Trong xưởng -> giữ;
   *   + trống/mâu thuẫn -> giữ.
   *
   * Record trước Ngày sửa chữa TUYỆT ĐỐI không được dùng để clear.
   *********************************************************/
  const candidates =
    audit.rows.filter(function(item) {
      return (
        item.safeToComplete === true &&
        item.decision &&
        item.decision.matchType === 'exact' &&
        !!clean_(item.visitId) &&
        (item.guard !== true || item.guardRecoverable === true) &&
        item.duplicateVisitId !== true &&
        normalizeKey_(item.status) === 'cho giao'
      );
    });
  const applied =
    applyWaitingCvdvCandidatesBulk_(
      waiting,
      candidates,
      ss,
      evidence
    );
  console.log(
    'CVDV AUTO CLEAR V13.20.3' +
    ' | scanned=' + audit.rows.length +
    ' | exactSafe=' + audit.exactSafe +
    ' | exactCompleted=' + applied.completed +
    ' | hondaCompleted=' + honda.completed +
    ' | hondaEligible=' + honda.eligible +
    ' | dateRepaired=' + dateRepair.repaired +
    ' | guarded=' + audit.guarded +
    ' | guardRecoverable=' + audit.guardRecoverable +
    ' | guardRecovered=' + applied.guardRecovered +
    ' | missingVisitId=' + audit.missingVisitId +
    ' | duplicateVisitId=' + audit.duplicateVisitId +
    ' | notWaiting=' + audit.notWaiting +
    ' | skippedGuarded=' + applied.skippedGuarded +
    ' | skippedUnsafe=' + applied.skippedUnsafeEvidence
  );
  return {
    success: true,
    scanned: audit.rows.length,
    matched: audit.matched,
    completed: applied.completed + honda.completed,
    exactCompleted: applied.completed,
    hondaCompleted: honda.completed,
    hondaEligible: honda.eligible,
    hondaChecked: honda.checked,
    dateRepaired: dateRepair.repaired,
    dateRepairChecked: dateRepair.checked,
    kept: audit.rows.length - applied.completed,
    exactSafe: audit.exactSafe,
    afterRepairSafe: audit.afterRepairSafe,
    sourceMissing: audit.sourceMissing,
    sourceUnclear: audit.sourceUnclear,
    inShop: audit.inShop,
    sourceOlder: audit.sourceOlder,
    guarded: audit.guarded,
    guardRecoverable: audit.guardRecoverable,
    guardRecovered: applied.guardRecovered,
    missingVisitId: audit.missingVisitId,
    duplicateVisitId: audit.duplicateVisitId,
    notWaiting: audit.notWaiting,
    skippedVisitConflict: applied.skippedVisitConflict,
    skippedChangedRow: applied.skippedChangedRow,
    skippedGuarded: applied.skippedGuarded,
    skippedMissingVisitId: applied.skippedMissingVisitId,
    skippedUnsafeEvidence: applied.skippedUnsafeEvidence,
    mode: options.mode || 'auto',
    reconcileMaxDays: 0
  };
}
/*********************************************************
 * BUILD DATA_CVDV EVIDENCE INDEX
 *
 * Return:
 * {
 *   exact: {"yyyy-MM-dd|PLATE": summarizedGroup},
 *   byPlateDates: {
 *     "PLATE": {
 *       dateKeys: ["yyyy-MM-dd", ...],
 *       groups: {"yyyy-MM-dd": summarizedGroup}
 *     }
 *   },
 *   latestByPlate: {"PLATE": summarizedGroup} // compatibility
 * }
 *
 * Duplicate cùng ngày + cùng biển số:
 * - TẤT CẢ Đã ra cổng => safe.
 * - Có Trong xưởng => giữ.
 * - Trống/mâu thuẫn => giữ.
 *********************************************************/
function loadLatestCVDVStatusByVisit_(source, ss) {
  const first = CFG.CVDV_DATA_FIRST_ROW;
  const last = source.getLastRow();
  const exactGroups = {};
  const groupsByPlateRaw = {};
  if (last < first) {
    return {
      exact: {},
      byPlateDates: {},
      latestByPlate: {}
    };
  }
  const width = Math.max(
    CFG.CVDV_DATA.NGAY,
    CFG.CVDV_DATA.BIEN_SO,
    CFG.CVDV_DATA.TINH_TRANG,
    CFG.CVDV_DATA.CVDV
  );
  const rowCount = last - first + 1;
  const range = source.getRange(first, 1, rowCount, width);
  const values = range.getValues();
  const display = range.getDisplayValues();
  // P0.5.2: Đọc timezone một lần, không lặp RPC theo từng dòng.
  // Cache theo timestamp; cùng ngày/giờ thực tế luôn tạo chính xác key cũ.
  const timezone = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  const dateKeyCache = Object.create(null);
  for (let i = 0; i < values.length; i++) {
    const sourceDate = values[i][CFG.CVDV_DATA.NGAY - 1];
    let dateKey = '';
    if (sourceDate instanceof Date && !isNaN(sourceDate.getTime())) {
      const stamp = String(sourceDate.getTime());
      if (Object.prototype.hasOwnProperty.call(dateKeyCache, stamp)) {
        dateKey = dateKeyCache[stamp];
      } else {
        dateKey = Utilities.formatDate(sourceDate, timezone, 'yyyy-MM-dd');
        dateKeyCache[stamp] = dateKey;
      }
    }
    const plate =
      normalizePlate_(
        display[i][CFG.CVDV_DATA.BIEN_SO - 1]
      );
    const status =
      clean_(
        display[i][CFG.CVDV_DATA.TINH_TRANG - 1]
      );
    const advisor =
      clean_(
        display[i][CFG.CVDV_DATA.CVDV - 1]
      );
    if (!dateKey || !plate) continue;
    const key = dateKey + '|' + plate;
    if (!exactGroups[key]) {
      exactGroups[key] = {
        dateKey: dateKey,
        plate: plate,
        records: []
      };
    }
    exactGroups[key].records.push({
      status: status,
      advisor: advisor,
      sourceRow: first + i
    });
    if (!groupsByPlateRaw[plate]) {
      groupsByPlateRaw[plate] = {};
    }
    groupsByPlateRaw[plate][dateKey] =
      exactGroups[key];
  }
  const exact = {};
  Object.keys(exactGroups).forEach(function(key) {
    exact[key] =
      summarizeCvdvEvidenceGroup_(
        exactGroups[key]
      );
  });
  const byPlateDates = {};
  const latestByPlate = {};
  Object.keys(groupsByPlateRaw).forEach(function(plate) {
    const dateKeys =
      Object.keys(
        groupsByPlateRaw[plate]
      ).sort();
    const groups = {};
    dateKeys.forEach(function(dateKey) {
      groups[dateKey] =
        exact[dateKey + '|' + plate];
    });
    byPlateDates[plate] = {
      dateKeys: dateKeys,
      groups: groups
    };
    if (dateKeys.length) {
      latestByPlate[plate] =
        groups[
          dateKeys[dateKeys.length - 1]
        ];
    }
  });
  return {
    exact: exact,
    byPlateDates: byPlateDates,
    latestByPlate: latestByPlate
  };
}
function summarizeCvdvEvidenceGroup_(group) {
  const records =
    (group && group.records) || [];
  const normalized =
    records.map(function(rec) {
      return normalizeKey_(rec.status);
    });
  const allExited =
    records.length > 0 &&
    normalized.every(function(key) {
      return key === 'da ra cong';
    });
  const hasInShop =
    normalized.some(function(key) {
      return key === 'trong xuong';
    });
  const statuses =
    uniqueArray_(
      records.map(function(rec) {
        return clean_(rec.status) || '(trống)';
      })
    );
  return {
    dateKey: group.dateKey,
    plate: group.plate,
    records: records.slice(),
    sourceRows:
      records.map(function(rec) {
        return rec.sourceRow;
      }),
    statusText: statuses.join(' | '),
    advisorKeys: uniqueArray_(records.map(function(rec) {
      return waitingAdvisorKey_(rec.advisor);
    }).filter(function(key) { return !!key; })),
    allExited: allExited,
    hasInShop: hasInShop,
    unclear: !allExited && !hasInShop
  };
}
/*********************************************************
 * ĐÁNH GIÁ 1 DÒNG XE CHỜ GIAO
 *********************************************************/
function evaluateWaitingCvdvEvidence_(evidence, waitingDateKey, plate, waitingAdvisor) {
  const normalizedPlate = normalizePlate_(plate);
  if (!normalizedPlate || !waitingDateKey) {
    return waitingCvdvDecision_(
      false,
      'GIỮ - dữ liệu đầu vào chưa hợp lệ',
      'Thiếu biển số hoặc ngày sửa chữa hợp lệ.',
      '',
      '',
      [],
      'invalid'
    );
  }
  // ===================================================
  // V13.15 P0 SAFE
  // Chỉ EXACT Ngày tiếp nhận + BSX mới có quyền auto-clear.
  // Record ngày trước / ngày sau chỉ dùng để giải thích, tuyệt đối
  // không biến thành safeToComplete.
  // ===================================================
  const exactKey =
    waitingDateKey + '|' + normalizedPlate;
  const exact =
    evidence && evidence.exact
      ? evidence.exact[exactKey]
      : null;
  if (!exact) {
    const plateIndex =
      evidence && evidence.byPlateDates
        ? evidence.byPlateDates[normalizedPlate]
        : null;
    if (
      plateIndex &&
      plateIndex.dateKeys &&
      plateIndex.dateKeys.length
    ) {
      const later = plateIndex.dateKeys.filter(function(dateKey) {
        return dateKey > waitingDateKey;
      });
      const older = plateIndex.dateKeys.filter(function(dateKey) {
        return dateKey < waitingDateKey;
      });
      if (later.length) {
        const key = later[0];
        const group = plateIndex.groups[key];
        return waitingCvdvDecision_(
          false,
          'GIỮ - chỉ có nguồn ngày sau',
          'Có record CVDV ngày sau nhưng V13.15 không dùng ngày sau để auto-clear lượt hiện tại.',
          group.dateKey,
          group.statusText,
          group.sourceRows,
          'later-observed'
        );
      }
      if (older.length) {
        const key = older[older.length - 1];
        const group = plateIndex.groups[key];
        return waitingCvdvDecision_(
          false,
          'GIỮ - nguồn cũ hơn ngày sửa chữa',
          'Chỉ có record trước ngày sửa chữa; không dùng để clear lượt hiện tại.',
          group.dateKey,
          group.statusText,
          group.sourceRows,
          'older'
        );
      }
    }
    return waitingCvdvDecision_(
      false,
      'GIỮ - không có EXACT Ngày + BSX',
      'Không tìm thấy Data_CVDV đúng Ngày tiếp nhận + BSX của lượt Xe chờ giao.',
      '',
      '',
      [],
      'none'
    );
  }
  if (exact.hasInShop) {
    return waitingCvdvDecision_(
      false,
      'GIỮ - EXACT còn Trong xưởng',
      'Cùng Ngày + BSX vẫn có trạng thái Trong xưởng.',
      exact.dateKey,
      exact.statusText,
      exact.sourceRows,
      'exact'
    );
  }
  if (exact.allExited) {
    // P0.5.4.2 FAIL-CLOSED: CVDV xác nhận đúng Ngày+Biển vẫn chưa đủ
    // nếu cố vấn của chính lượt Xe chờ giao bị lệch hoặc trùng nguồn.
    // Áp dụng cho Toyota; Honda vẫn thuộc cơ chế tách riêng hiện hữu.
    const wanted = waitingAdvisorKey_(waitingAdvisor);
    if (!wanted || !exact.advisorKeys || exact.advisorKeys.length !== 1 ||
        exact.advisorKeys[0] !== wanted) {
      return waitingCvdvDecision_(
        false,
        'GIỮ - CVDV lệch cố vấn cùng lượt',
        'Dù CVDV ghi Đã ra cổng nhưng cố vấn không khớp Xe chờ giao; không tự dọn.',
        exact.dateKey, exact.statusText, exact.sourceRows, 'exact'
      );
    }
    return waitingCvdvDecision_(
      true,
      'DỌN - EXACT cùng ngày Đã ra cổng',
      'Tất cả record Data_CVDV đúng Ngày + BSX đều xác nhận Đã ra cổng.',
      exact.dateKey,
      exact.statusText,
      exact.sourceRows,
      'exact'
    );
  }
  return waitingCvdvDecision_(
    false,
    'GIỮ - EXACT trạng thái chưa rõ',
    'Cùng Ngày + BSX có trạng thái trống hoặc mâu thuẫn; không tự clear.',
    exact.dateKey,
    exact.statusText,
    exact.sourceRows,
    'exact'
  );
}

function waitingCvdvDecision_(
  safeToComplete,
  label,
  reason,
  sourceDateKey,
  sourceStatus,
  sourceRows,
  matchType
) {
  return {
    safeToComplete: !!safeToComplete,
    label: label || '',
    reason: reason || '',
    sourceDateKey: sourceDateKey || '',
    sourceStatus: sourceStatus || '',
    sourceRows: sourceRows || [],
    matchType: matchType || ''
  };
}
function collectWaitingCvdvAudit_(waiting, evidence, ss) {
  const last = waitingDeliveryDataLastRow_(waiting);
  const result = {
    rows: [],
    matched: 0,
    exactSafe: 0,
    afterRepairSafe: 0, // giữ field compatibility; guard recovery được đếm riêng ở guardRecoverable
    latestSafe: 0,
    sourceMissing: 0,
    sourceUnclear: 0,
    inShop: 0,
    sourceOlder: 0,
    sourceTooLate: 0,
    duplicateVisitId: 0,
    missingVisitId: 0,
    guarded: 0,
    guardRecoverable: 0,
    notWaiting: 0
  };
  if (last < CFG.WAITING_DELIVERY_FIRST_ROW) {
    return result;
  }
  const start = CFG.WAITING_DELIVERY_FIRST_ROW;
  const rowCount = last - start + 1;
  const width = CFG.WD.VISIT_ID;
  const range =
    waiting.getRange(
      start,
      1,
      rowCount,
      width
    );
  const values = range.getValues();
  const display = range.getDisplayValues();
  const visitCount = {};
  for (let i = 0; i < rowCount; i++) {
    const visitId =
      clean_(
        values[i][CFG.WD.VISIT_ID - 1]
      );
    if (visitId) {
      visitCount[visitId] =
        (visitCount[visitId] || 0) + 1;
    }
  }
  for (let i = 0; i < rowCount; i++) {
    const row = start + i;
    const plate =
      normalizePlate_(
        display[i][CFG.WD.BIEN_SO - 1]
      );
    const waitingDateKey =
      cvdvDateKey_(
        values[i][CFG.WD.NGAY - 1],
        ss
      );
    const visitId =
      clean_(
        values[i][CFG.WD.VISIT_ID - 1]
      );
    const status =
      clean_(
        display[i][CFG.WD.TRANG_THAI - 1]
      );
    const statusKey =
      normalizeKey_(status);
    let decision =
      evaluateWaitingCvdvEvidence_(
        evidence,
        waitingDateKey,
        plate,
        display[i][CFG.WD.CVDV - 1]
      );
    let guard =
      visitId
        ? isWaitingVisitCvdvUnsafe_(visitId)
        : false;
    let guardRecoverable = false;
    const duplicateVisitId =
      !!(
        visitId &&
        visitCount[visitId] > 1
      );
    // =================================================
    // P0 HARD BLOCKS - thứ tự ưu tiên an toàn
    // =================================================
    if (!visitId) {
      result.missingVisitId++;
      decision =
        waitingCvdvDecision_(
          false,
          'CẢNH BÁO - thiếu visitId',
          'Xe chờ giao thiếu visitId; không được auto-clear.',
          decision.sourceDateKey,
          decision.sourceStatus,
          decision.sourceRows,
          decision.matchType
        );
    }
    else if (duplicateVisitId) {
      result.duplicateVisitId++;
      decision =
        waitingCvdvDecision_(
          false,
          'CẢNH BÁO - visitId trùng',
          'visitId xuất hiện nhiều hơn 1 dòng trong Xe chờ giao; không tự dọn.',
          decision.sourceDateKey,
          decision.sourceStatus,
          decision.sourceRows,
          decision.matchType
        );
    }
    else if (statusKey !== 'cho giao') {
      result.notWaiting++;
      decision =
        waitingCvdvDecision_(
          false,
          'GIỮ - trạng thái không phải Chờ giao',
          'Chỉ dòng đang ở trạng thái Chờ giao mới được xét auto-clear.',
          decision.sourceDateKey,
          decision.sourceStatus,
          decision.sourceRows,
          decision.matchType
        );
    }
    else if (guard) {
      guardRecoverable =
        isWaitingCvdvGuardRecoverable_(
          visitId,
          duplicateVisitId,
          statusKey,
          decision
        );
      if (guardRecoverable) {
        // V13.20.2: chưa xóa guard ở audit. Chỉ đánh dấu candidate recoverable.
        // Guard + evidence sẽ được re-check lại ngay trước remove trong bulk apply.
        result.guardRecoverable++;
      }
      else {
        result.guarded++;
        decision =
          waitingCvdvDecision_(
            false,
            'CẢNH BÁO - CVDV unsafe guard',
            'Lượt xe đang mang WAITING_CVDV_GUARD và chưa đủ điều kiện recovery an toàn.',
            decision.sourceDateKey,
            decision.sourceStatus,
            decision.sourceRows,
            decision.matchType
          );
      }
    }
    if (decision.sourceDateKey) {
      result.matched++;
    }
    if (decision.safeToComplete) {
      // V13.15: safe chỉ có thể là exact.
      if (decision.matchType === 'exact') {
        result.exactSafe++;
      }
    }
    else if (decision.label.indexOf('Trong xưởng') >= 0) {
      result.inShop++;
    }
    else if (decision.label.indexOf('nguồn cũ hơn') >= 0) {
      result.sourceOlder++;
    }
    else if (
      decision.label.indexOf('không có') >= 0 ||
      decision.label.indexOf('ngày sau') >= 0
    ) {
      result.sourceMissing++;
    }
    else {
      result.sourceUnclear++;
    }
    result.rows.push({
      row: row,
      offset: i,
      dateValue: values[i][CFG.WD.NGAY - 1],
      dateKey: waitingDateKey,
      plate: plate,
      advisor: clean_(display[i][CFG.WD.CVDV - 1]),
      model: clean_(display[i][CFG.WD.MODEL - 1]),
      deliveryDate: values[i][CFG.WD.NGAY_GIAO - 1],
      deliveryTime: values[i][CFG.WD.GIO_GIAO - 1],
      status: status,
      visitId: visitId,
      guard: guard,
      guardRecoverable: guardRecoverable,
      duplicateVisitId: duplicateVisitId,
      safeToComplete: decision.safeToComplete,
      decision: decision
    });
  }
  return result;
}
/*********************************************************
 * BULK COMPACT - 1 SNAPSHOT / 1 WRITE
 *********************************************************/
function applyWaitingCvdvCandidatesBulk_(waiting, candidates, ss, evidence) {
  const out = {
    completed: 0,
    skippedVisitConflict: 0,
    skippedChangedRow: 0,
    skippedGuarded: 0,
    guardRecovered: 0,
    skippedMissingVisitId: 0,
    skippedUnsafeEvidence: 0
  };
  if (!candidates || !candidates.length) {
    return out;
  }
  const start = CFG.WAITING_DELIVERY_FIRST_ROW;
  const last = waitingDeliveryDataLastRow_(waiting);
  if (last < start) return out;
  const width = CFG.WD.VISIT_ID;
  const rowCount = last - start + 1;
  const range = waiting.getRange(start, 1, rowCount, width);
  const snapshot = waitingDeliveryCompactSnapshot_(range);
  const display = range.getDisplayValues();
  const visitCount = {};
  for (let i = 0; i < snapshot.values.length; i++) {
    const id =
      clean_(
        snapshot.values[i][CFG.WD.VISIT_ID - 1]
      );
    if (id) visitCount[id] = (visitCount[id] || 0) + 1;
  }
  const removeOffsets = [];
  const guardsToClear = [];
  candidates.forEach(function(item) {
    const offset = item.row - start;
    if (
      offset < 0 ||
      offset >= snapshot.values.length
    ) {
      out.skippedChangedRow++;
      return;
    }
    // V13.15: candidate phải vẫn là exact-safe từ audit.
    if (
      !item ||
      item.safeToComplete !== true ||
      !item.decision ||
      item.decision.matchType !== 'exact'
    ) {
      out.skippedUnsafeEvidence++;
      return;
    }
    if (!clean_(item.visitId)) {
      out.skippedMissingVisitId++;
      return;
    }
    const currentPlate =
      normalizePlate_(
        display[offset][CFG.WD.BIEN_SO - 1]
      );
    const currentDateKey =
      cvdvDateKey_(
        snapshot.values[offset][CFG.WD.NGAY - 1],
        ss
      );
    const currentStatus =
      normalizeKey_(
        display[offset][CFG.WD.TRANG_THAI - 1]
      );
    const currentVisitId =
      clean_(
        snapshot.values[offset][CFG.WD.VISIT_ID - 1]
      );
    if (
      currentPlate !== item.plate ||
      currentDateKey !== item.dateKey ||
      currentStatus !== 'cho giao' ||
      waitingAdvisorKey_(display[offset][CFG.WD.CVDV - 1]) !==
        waitingAdvisorKey_(item.advisor)
    ) {
      out.skippedChangedRow++;
      return;
    }
    if (!currentVisitId) {
      out.skippedMissingVisitId++;
      return;
    }
    if (visitCount[currentVisitId] !== 1) {
      out.skippedVisitConflict++;
      return;
    }
    // visitId phải đúng ownership đã audit.
    if (currentVisitId !== item.visitId) {
      out.skippedChangedRow++;
      return;
    }
    // V13.20.2: re-check evidence NGAY TRƯỚC khi remove.
    // Chỉ EXACT Ngày + BSX và toàn bộ record Đã ra cổng mới được tiếp tục.
    const currentDecision =
      evaluateWaitingCvdvEvidence_(
        evidence,
        currentDateKey,
        currentPlate,
        display[offset][CFG.WD.CVDV - 1]
      );
    if (
      !currentDecision ||
      currentDecision.safeToComplete !== true ||
      currentDecision.matchType !== 'exact'
    ) {
      out.skippedUnsafeEvidence++;
      return;
    }
    // Guard mới phát sinh sau audit luôn BLOCK. Guard có từ audit chỉ được
    // recovery khi chính candidate đã được đánh dấu recoverable và mọi
    // điều kiện ownership/status/evidence vẫn còn nguyên ở snapshot hiện tại.
    if (isWaitingVisitCvdvUnsafe_(currentVisitId)) {
      const mayRecover =
        item.guard === true &&
        item.guardRecoverable === true &&
        isWaitingCvdvGuardRecoverable_(
          currentVisitId,
          visitCount[currentVisitId] !== 1,
          currentStatus,
          currentDecision
        );
      if (!mayRecover) {
        out.skippedGuarded++;
        return;
      }
      clearWaitingVisitCvdvUnsafe_(currentVisitId);
      if (isWaitingVisitCvdvUnsafe_(currentVisitId)) {
        out.skippedGuarded++;
        return;
      }
      out.guardRecovered++;
    }
    removeOffsets.push(offset);
    guardsToClear.push(currentVisitId);
  });
  if (!removeOffsets.length) {
    return out;
  }
  const uniqueOffsets =
    Array.from(
      new Set(removeOffsets)
    ).sort(function(a, b) {
      return a - b;
    });
  const compacted =
    compactWaitingDeliverySnapshotMulti_(
      snapshot,
      uniqueOffsets,
      width
    );
  try {
    applyWaitingDeliveryCompact_(
      range,
      compacted
    );
    SpreadsheetApp.flush();
    const count =
      waiting.getMaxRows() - start + 1;
    if (count > 0) {
      waiting
        .getRange(
          start,
          CFG.WD.BIEN_SO,
          count,
          1
        )
        .setNumberFormat('@');
    }
    guardsToClear.forEach(function(visitId) {
      clearWaitingVisitCvdvUnsafe_(visitId);
    });
    out.completed = uniqueOffsets.length;
    return out;
  } catch (err) {
    try {
      restoreWaitingDeliveryCompactSnapshot_(
        range,
        snapshot
      );
      SpreadsheetApp.flush();
    } catch (restoreErr) {
      throw new Error(
        'Bulk compact Xe chờ giao lỗi: ' +
        err.message +
        ' | Khôi phục snapshot lỗi: ' +
        restoreErr.message
      );
    }
    throw new Error(
      'Bulk compact Xe chờ giao lỗi; đã khôi phục snapshot dữ liệu + validation: ' +
      err.message
    );
  }
}
function compactWaitingDeliverySnapshotMulti_(snapshot, offsets, width) {
  const remove = {};
  (offsets || []).forEach(function(offset) {
    remove[offset] = true;
  });
  const values = [];
  const notes = [];
  const validations = [];
  for (let i = 0; i < snapshot.values.length; i++) {
    if (remove[i]) continue;
    values.push(snapshot.values[i].slice());
    notes.push(snapshot.notes[i].slice());
    validations.push(snapshot.validations[i].slice());
  }
  while (values.length < snapshot.values.length) {
    values.push(Array(width).fill(''));
    notes.push(Array(width).fill(''));
    validations.push(Array(width).fill(null));
  }
  return {
    values: values,
    notes: notes,
    validations: validations
  };
}
function cvdvDateKey_(value, ss) {
  if (!(value instanceof Date) || isNaN(value.getTime())) return '';
  const tz =
    (ss && ss.getSpreadsheetTimeZone && ss.getSpreadsheetTimeZone()) ||
    Session.getScriptTimeZone();
  return Utilities.formatDate(
    value,
    tz,
    'yyyy-MM-dd'
  );
}
function caiTriggerDongBoCVDV() {
  const handler = 'dongBoTrangThaiCVDV';
  ScriptApp
    .getProjectTriggers()
    .filter(function(trigger) {
      return trigger.getHandlerFunction() === handler;
    })
    .forEach(function(trigger) {
      ScriptApp.deleteTrigger(trigger);
    });
  ScriptApp
    .newTrigger(handler)
    .timeBased()
    .everyMinutes(5)
    .create();
  const ss = SpreadsheetApp.getActive();
  if (ss) {
    ss.toast(
      'Đã cài đồng bộ Data_CVDV → Xe chờ giao mỗi 5 phút.',
      '✅ TRIGGER CVDV',
      6
    );
  }
  return {
    success: true,
    everyMinutes: 5
  };
}
function goTriggerDongBoCVDV() {
  const handler = 'dongBoTrangThaiCVDV';
  let removed = 0;
  ScriptApp
    .getProjectTriggers()
    .filter(function(trigger) {
      return trigger.getHandlerFunction() === handler;
    })
    .forEach(function(trigger) {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    });
  const ss = SpreadsheetApp.getActive();
  if (ss) {
    ss.toast(
      removed > 0
        ? 'Đã gỡ ' + removed + ' trigger đồng bộ CVDV.'
        : 'Không có trigger đồng bộ CVDV để gỡ.',
      '🛑 TRIGGER CVDV',
      6
    );
  }
  return {
    success: true,
    removed: removed
  };
}
/*********************************************************
 * MIGRATION V4 - THÊM TIẾN ĐỘ RỬA XE TẠI CỘT R
 *********************************************************/

function repairVisitSchema_(ss) {
  const configured = clean_(PropertiesService.getScriptProperties().getProperty('VISIT_SCHEMA_' + ss.getId()));
  const schema = visitSchema_().map(spec => ({
    name: spec.name, start: spec.start, header: spec.header, col: spec.col, fields: spec.fields.slice()
  }));
  if (configured === '3') schema[0].col = 21;
  return schema;
}
function auditDuplicateVisitIds_(ss) {
  const byId = {};
  allVisitIdSpecs_(ss).forEach(function(spec) {
    const sh = ss.getSheetByName(spec.name);
    if (!sh) return;
    const last = sh.getLastRow();
    if (last < spec.start) return;
    const count = last - spec.start + 1;
    const maxCol = Math.max(spec.plateCol, spec.idCol);
    const values = sh.getRange(spec.start, 1, count, maxCol).getValues();
    const display = sh.getRange(spec.start, 1, count, maxCol).getDisplayValues();
    for (let i = 0; i < count; i++) {
      const id = clean_(values[i][spec.idCol - 1]);
      const plate = normalizePlate_(display[i][spec.plateCol - 1]);
      if (!id || !plate) continue;
      (byId[id] || (byId[id] = [])).push({
        sheet: spec.name,
        row: spec.start + i,
        plate: plate,
        idCol: spec.idCol
      });
    }
  });
  const conflicts = [];
  Object.keys(byId).forEach(function(id) {
    const rows = byId[id];
    const plates = {};
    rows.forEach(function(item) {
      plates[item.plate] = true;
    });
    const uniquePlates = Object.keys(plates);
    if (uniquePlates.length > 1) {
      conflicts.push({
        visitId: id,
        plates: uniquePlates,
        rows: rows
      });
    }
  });
  return {
    conflictCount: conflicts.length,
    conflicts: conflicts
  };
}
function kiemTraDuplicateVisitIdToanHeThong() {
  const ss = laySpreadsheetWeb_();
  const result = auditDuplicateVisitIds_(ss);
  console.log(JSON.stringify(result, null, 2));
  try {
    if (!result.conflictCount) {
      ss.toast(
        'Không phát hiện visitId nào đang được dùng cho 2 biển số khác nhau.',
        'Kiểm tra visitId',
        8
      );
    } else {
      ss.toast(
        'Phát hiện ' + result.conflictCount +
        ' visitId bị trùng xe. Xem Execution log để xem chi tiết.',
        '⚠️ Kiểm tra visitId',
        10
      );
    }
  } catch (err) {
    console.log('Không hiển thị toast được: ' + err.message);
  }
  return result;
}
/*********************************************************
 * CASE DUPLICATE PHÁT HIỆN 15:39
 * 51L26555 và 64A07649 cùng dùng:
 * 653b66eb-71d9-4cd7-92b4-1b52eec85565
 *
 * Theo cấu trúc hiện tại:
 * - 51L26555 nằm ở Data EM + Theo dõi KT
 * - 64A07649 nằm ở Xe chờ giao
 *
 * Hàm dưới đây chỉ PREVIEW/SỬA đúng 51L26555.
 *********************************************************/

/*********************************************************
 * V13.20.1 - RESTORE THÔNG TIN TẢI XƯỞNG CONFIG
 *
 * Hai constant này tồn tại ở V13.14.1 nhưng bị rơi khỏi
 * file MAIN V13.20.0 trong quá trình merge/offload.
 * Thiếu chúng sẽ làm capNhatTaiXuong/runTaiXuongAuto/
 * caiTriggerTaiXuong lỗi ReferenceError và dashboard tải
 * xưởng không thể refresh.
 *********************************************************/
const TAI_XUONG_CFG = Object.freeze({
  SHEET_NAME: 'Thông tin tải xưởng',
  APPOINTMENT_SHEET: 'Data_LichHen',
  APPOINTMENT_SOURCE_ID: '10dxPgBNvoYZ4oDydtf1-dNsGLxfHxJiRJ5l7uoJIqso',
  APPOINTMENT_SOURCE_SHEET: 'Danh sách hẹn',

  FIRST_TIME_COL: 2, // B
  TIMES: Object.freeze([
    '07:30', '08:00', '08:30', '09:00', '09:30',
    '10:00', '10:30', '11:00', '11:30', '12:00',
    '13:00', '13:30', '14:00', '14:30', '15:00',
    '15:30', '16:00', '16:30', '17:00'
  ]),

  EM_GROUPS: Object.freeze(['EM1', 'EM2', 'EM3', 'EM4', 'EM5']),
  SCC_GROUPS: Object.freeze([
    'Tú - Hải',
    'Ân - Dũng - Việt',
    'Nhân - Hạnh - Quí',
    'Lợi - Duy',
    'Minh - Khang'
  ]),
  INSPECTORS: Object.freeze([
    'Mr. Thanh',
    'Mr. Đua',
    'Mr. Lợi',
    'Mr. Giang',
    'Mr. Phong'
  ]),

  MIN_LANE_HEIGHT: 84,
  TIME_COL_WIDTH: 78,
  LABEL_COL_WIDTH: 165,
  LAST_SLOT_MINUTES: 30,

  COLORS: Object.freeze({
    TITLE_BG: '#B7D7A8',
    TITLE_TEXT: '#102A13',
    TIME_BG: '#F3F5F7',
    GRID: '#D9DEE5',
    APPT_BG: '#FFF2CC',
    APPT_HIT_BG: '#FFD966',
    EM_SECTION: '#D9EAF0',
    EM_BLOCK: '#B7DEE8',
    EM_BORDER: '#4F81BD',
    SCC_SECTION: '#FCE4D6',
    SCC_BLOCK: '#F8CBAD',
    SCC_BORDER: '#C55A11',
    KT_SECTION: '#E4DFEC',
    KT_BLOCK: '#D9D2E9',
    KT_BORDER: '#8064A2',
    OVERDUE: '#F4CCCC',
    OVERDUE_BORDER: '#CC0000',
    SOON: '#FFF2CC',
    SOON_BORDER: '#BF9000',
    EMPTY: '#FFFFFF',
    TEXT: '#202124',
    MUTED: '#5F6368',
    NOW: '#D93025'
  })
});

const TAI_XUONG_AUTO_CFG = Object.freeze({
  LAST_RUN_PROP: 'TPMH_TAI_XUONG_LAST_AUTO_MS',
  EFFECTIVE_INTERVAL_MS: 3 * 60 * 1000,
  // Apps Script không hỗ trợ everyMinutes(3). Trigger đánh thức mỗi 1 phút,
  // nhưng phần render nặng chỉ chạy tối đa 1 lần / 3 phút.
  TRIGGER_WAKE_MINUTES: 1
});

function capNhatTaiXuong() {
  const result = renderTaiXuong_({showToast: true});
  // Cập nhật tay cũng được xem là một lần render thành công,
  // tránh trigger tự động chạy lại ngay sau đó.
  if (result && result.success && !result.busy) {
    PropertiesService
      .getScriptProperties()
      .setProperty(TAI_XUONG_AUTO_CFG.LAST_RUN_PROP, String(Date.now()));
  }
  return result;
}
function runTaiXuongAuto() {
  try {
    // Ưu tiên tuyệt đối thao tác XÁC NHẬN từ WebApp.
    // Check trước mọi openById / render / lock để lượt auto gần như không tạo tải khi WebApp đang bận.
    if (isWebAppConfirmPriorityActive_()) {
      return {
        success: true,
        skipped: true,
        webAppPriority: true,
        message: 'Bỏ qua auto tải xưởng vì WebApp đang ưu tiên.'
      };
    }
    const props = PropertiesService.getScriptProperties();
    const now = Date.now();
    const last = Number(props.getProperty(TAI_XUONG_AUTO_CFG.LAST_RUN_PROP) || 0);
    if (last > 0 && now - last < TAI_XUONG_AUTO_CFG.EFFECTIVE_INTERVAL_MS) {
      return {
        success: true,
        skipped: true,
        throttled: true,
        nextInMs: TAI_XUONG_AUTO_CFG.EFFECTIVE_INTERVAL_MS - (now - last)
      };
    }
    const result = renderTaiXuong_({showToast: false, auto: true});
    // Chỉ ghi mốc khi render thực sự thành công và không bị busy.
    if (result && result.success && !result.busy) {
      props.setProperty(TAI_XUONG_AUTO_CFG.LAST_RUN_PROP, String(Date.now()));
    }
    return result;
  }
  catch (err) {
    console.log('Không cập nhật được Thông tin tải xưởng: ' + err.message);
    return {
      success: false,
      error: err.message
    };
  }
}
function caiTriggerTaiXuong() {
  goTriggerTaiXuong();
  ScriptApp
    .newTrigger('runTaiXuongAuto')
    .timeBased()
    .everyMinutes(TAI_XUONG_AUTO_CFG.TRIGGER_WAKE_MINUTES)
    .create();
  // Reset mốc để lượt auto đầu tiên sau khi cài có thể chạy ngay.
  PropertiesService
    .getScriptProperties()
    .deleteProperty(TAI_XUONG_AUTO_CFG.LAST_RUN_PROP);
  const ss = laySpreadsheetWeb_();
  try {
    ss.toast(
      'Đã cài tải xưởng tự động: render nặng tối đa 1 lần / 3 phút.',
      '📊 TẢI XƯỞNG',
      6
    );
  }
  catch (err) {}
  return {
    success: true,
    effectiveMinutes: 3,
    message: 'Đã cài tự động tải xưởng theo nhịp hiệu dụng 3 phút.'
  };
}
function goTriggerTaiXuong() {
  let removed = 0;
  ScriptApp
    .getProjectTriggers()
    .forEach(function(trigger) {
      if (trigger.getHandlerFunction() === 'runTaiXuongAuto') {
        ScriptApp.deleteTrigger(trigger);
        removed++;
      }
    });
  PropertiesService
    .getScriptProperties()
    .deleteProperty(TAI_XUONG_AUTO_CFG.LAST_RUN_PROP);
  return {
    success: true,
    removed: removed
  };
}
function renderTaiXuong_(options) {
  options = options || {};
  const p052Started = Date.now(), p052Phases = {};
  let p052Last = p052Started;
  const p052Mark = function(label) {
    const now = Date.now(); p052Phases[label] = now - p052Last; p052Last = now;
  };
  if (options.auto && isWebAppConfirmPriorityActive_()) {
    return {
      success: true,
      skipped: true,
      webAppPriority: true
    };
  }
  // UserLock chỉ chống 2 lượt render dashboard đè nhau.
  // Không dùng ScriptLock / DocumentLock để tránh tranh lock với WebApp / pending queue.
  const renderLock = LockService.getUserLock();
  if (!renderLock.tryLock(1200)) {
    return {
      success: true,
      busy: true,
      message: 'Dashboard đang được cập nhật bởi một lượt khác.'
    };
  }
  try {
    const ss = laySpreadsheetWeb_();
    const sh = ss.getSheetByName(TAI_XUONG_CFG.SHEET_NAME);
    const dp = ss.getSheetByName(CFG.SHEET_DISPATCH);
    if (!sh) {
      throw new Error('Không tìm thấy sheet "' + TAI_XUONG_CFG.SHEET_NAME + '".');
    }
    if (!dp) {
      throw new Error('Không tìm thấy sheet "' + CFG.SHEET_DISPATCH + '".');
    }
    const selectedDateInfo = taiXuongSelectedDate_(sh);
    const selectedDate = selectedDateInfo.date;
    // C2 là ngày làm việc chuẩn của dashboard.
    taiXuongPrimeAppointmentDate_(
      sh,
      selectedDate,
      selectedDateInfo.formula
    );
    // V13.12.5: không dùng IMPORTRANGE cho helper Data_LichHen nữa.
    // Đọc trực tiếp file nguồn bằng SpreadsheetApp.openById() rồi ghi snapshot A:G.
    // Cách này tránh IFERROR che lỗi IMPORTRANGE / spill chưa refresh khiến dashboard báo 0 giả.
    const appointmentSync = taiXuongSyncAppointmentData_(ss, selectedDate);
    p052Mark('sourceAppointmentMs');
    const records = taiXuongReadDispatch_(dp, selectedDate);
    // Người KT dùng giờ bắt đầu/kết thúc thực tế từ tracking thay vì kéo từ giờ nhận tới giờ giao.
    records.kt = taiXuongReadTracking_(ss, selectedDate, records);
    const appointments = taiXuongReadAppointments_(ss, selectedDate, records.allPlates);
    p052Mark('readDispatchTrackingMs');
    const emGroups = taiXuongBuildGroups_(
      TAI_XUONG_CFG.EM_GROUPS,
      records.em,
      'emGroup'
    );
    const sccGroups = taiXuongBuildGroups_(
      TAI_XUONG_CFG.SCC_GROUPS,
      records.scc,
      'sccGroup'
    );
    const ktGroups = taiXuongBuildGroups_(
      TAI_XUONG_CFG.INSPECTORS,
      records.kt,
      'inspector'
    );
    const layout = taiXuongBuildLayout_(emGroups, sccGroups, ktGroups);
    taiXuongPrepareSheet_(sh, layout.lastRow);
    p052Mark('prepareGridMs');
    taiXuongRenderHeader_(
      sh,
      selectedDate,
      selectedDateInfo.formula,
      appointments,
      records,
      layout.lastRow
    );
    p052Mark('renderHeaderMs');
    taiXuongRenderSection_(sh, layout.em, 'NHÓM EM', emGroups, 'EM');
    taiXuongRenderSection_(sh, layout.scc, 'NHÓM SCC', sccGroups, 'SCC');
    taiXuongRenderSection_(sh, layout.kt, 'NGƯỜI KT', ktGroups, 'KT');
    p052Mark('renderCardsMs');
    taiXuongMarkCurrentTime_(sh, selectedDate, layout.lastRow);
    // Khóa lại toàn bộ dashboard về Arial SAU khi render RichText.
    taiXuongLockArial_(sh, layout.lastRow);
    SpreadsheetApp.flush();
    p052Mark('finalizeAndFlushMs');
    console.log('[P052_PERF_TAI_XUONG] ' + JSON.stringify({
      phasesMs: p052Phases, totalMs: Date.now() - p052Started,
      appointmentSnapshotUnchanged: appointmentSync.unchanged === true,
      rows: layout.lastRow
    }));
    if (options.showToast) {
      try {
        ss.toast(
          'Đã cập nhật timeline tải xưởng.',
          '📊 THÔNG TIN TẢI XƯỞNG',
          4
        );
      }
      catch (err) {}
    }
    return {
      success: true,
      busy: false,
      date: selectedDate,
      activeVehicles: records.activeVehicleCount,
      appointments: appointments.total,
      appointmentSourceRows: appointmentSync.rows,
      rows: layout.lastRow
    };
  }
  finally {
    renderLock.releaseLock();
  }
}
function taiXuongSelectedDate_(sh) {
  // C2 là ô ngày chuẩn của dashboard.
  // Để nâng cấp không làm mất ngày ở bản cũ, lần đầu vẫn fallback sang B2.
  const candidates = ['C2', 'B2'];
  for (let i = 0; i < candidates.length; i++) {
    const cell = sh.getRange(candidates[i]);
    const value = cell.getValue();
    const formula = clean_(cell.getFormula());
    if (
      value instanceof Date &&
      !isNaN(value.getTime())
    ) {
      return {
        date: stripTime_(value),
        formula: formula
      };
    }
  }
  return {
    date: todayOnly_(),
    formula: '=TODAY()'
  };
}
function taiXuongPrimeAppointmentDate_(sh, selectedDate, selectedDateFormula) {
  const cell = sh.getRange('C2');
  const current = cell.getValue();
  const hadSameDate =
    current instanceof Date &&
    !isNaN(current.getTime()) &&
    sameDate_(current, selectedDate);
  // Tránh ghi C2 lặp lại trước khi dashboard được render; header render sẽ
  // xây lại C2 sau khi clear. Khi ngày/công thức thay đổi mới prime + flush.
  const sameFormula = clean_(cell.getFormula()) === clean_(selectedDateFormula);
  if (hadSameDate && sameFormula) return false;
  if (selectedDateFormula) {
    cell.setFormula(selectedDateFormula);
  }
  else {
    cell.setValue(selectedDate);
  }
  cell
    .setNumberFormat('dd/MM/yyyy')
    .setFontFamily('Arial')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  if (!hadSameDate) {
    SpreadsheetApp.flush();
  }
  return !hadSameDate;
}
function taiXuongSyncAppointmentData_(ss, selectedDate) {
  const helper = ss.getSheetByName(TAI_XUONG_CFG.APPOINTMENT_SHEET);
  if (!helper) {
    throw new Error('Không tìm thấy sheet "' + TAI_XUONG_CFG.APPOINTMENT_SHEET + '".');
  }
  if (!(selectedDate instanceof Date) || isNaN(selectedDate.getTime())) {
    throw new Error('Ngày tải xưởng không hợp lệ khi đồng bộ lịch hẹn.');
  }
  // Header helper cố định A:G; chỉ cập nhật nếu thực sự bị thay đổi.
  const headerText = ['Biển số', 'CVDV', 'Nguồn hẹn', 'Kiểu xe',
    'Nội dung công việc', 'Thời gian hẹn', 'Trạng thái'];
  const header = helper.getRange(1, 1, 1, 7);
  const oldHeader = header.getDisplayValues()[0];
  const headerChanged = headerText.some(function(x, i) { return x !== oldHeader[i]; });
  if (headerChanged) {
    header.setValues([headerText]).setFontFamily('Arial').setFontWeight('bold');
  }
  let sourceSs;
  try {
    sourceSs = SpreadsheetApp.openById(TAI_XUONG_CFG.APPOINTMENT_SOURCE_ID);
  }
  catch (err) {
    throw new Error(
      'Không mở được file nguồn lịch hẹn. Kiểm tra quyền truy cập file GIỜ RA VÀO - LỊCH HẸN. Chi tiết: ' +
      err.message
    );
  }
  const sourceSh = sourceSs.getSheetByName(TAI_XUONG_CFG.APPOINTMENT_SOURCE_SHEET);
  if (!sourceSh) {
    throw new Error(
      'Không tìm thấy sheet "' +
      TAI_XUONG_CFG.APPOINTMENT_SOURCE_SHEET +
      '" trong file nguồn lịch hẹn.'
    );
  }
  const last = sourceSh.getLastRow();
  const output = [];
  if (last >= 2) {
    // Chỉ quét cột M (Thời gian hẹn) để xác định vùng của ngày đang xem.
    // File nguồn có thể >30.000 dòng nên cách này nhẹ hơn đọc toàn bộ B:Q mỗi phút.
    const dateCount = last - 1;
    const dateRange = sourceSh.getRange(2, 13, dateCount, 1); // M
    const dateValues = dateRange.getValues();
    const dateDisplay = dateRange.getDisplayValues(); // Giữ timezone hiển thị của nguồn.
    const matchedRows = [];
    for (let i = 0; i < dateValues.length; i++) {
      const appointmentAt = taiXuongAppointmentDateTime_(
        dateValues[i][0],
        dateDisplay[i][0],
        selectedDate
      );
      if (appointmentAt && sameDate_(appointmentAt, selectedDate)) {
        matchedRows.push(i + 2);
      }
    }
    if (matchedRows.length) {
      // Lịch hẹn trong file nguồn được sắp theo thời gian; đọc một vùng B:Q bao quanh
      // các dòng cùng ngày rồi lọc lại lần nữa để vẫn an toàn nếu có dòng xen kẽ.
      const firstRow = matchedRows[0];
      const lastRow = matchedRows[matchedRows.length - 1];
      const rowCount = lastRow - firstRow + 1;
      const sourceRange = sourceSh.getRange(firstRow, 2, rowCount, 16); // B:Q
      const values = sourceRange.getValues();
      const display = sourceRange.getDisplayValues();
      for (let i = 0; i < values.length; i++) {
        const v = values[i];
        const d = display[i];
        // B:Q index: B0 C1 F4 G5 L10 M11 Q15.
        const appointmentAt = taiXuongAppointmentDateTime_(
          v[11],
          d[11],
          selectedDate
        );
        if (!appointmentAt || !sameDate_(appointmentAt, selectedDate)) {
          continue;
        }
        const plate = clean_(d[0] || v[0]);
        if (!plate) continue;
        output.push([
          plate,
          clean_(d[1] || v[1]),
          clean_(d[4] || v[4]),
          clean_(d[5] || v[5]),
          clean_(d[10] || v[10]),
          appointmentAt,
          clean_(d[15] || v[15])
        ]);
      }
    }
  }
  // Sắp theo thời gian hẹn để Data_LichHen luôn dễ kiểm tra bằng mắt.
  output.sort(function(a, b) {
    const aa = a[5] instanceof Date ? a[5].getTime() : 0;
    const bb = b[5] instanceof Date ? b[5].getTime() : 0;
    return aa - bb;
  });
  // P0.5.2: Data_LichHen là snapshot. Không clear/rewrite và định dạng cả
  // cột mỗi 3 phút nếu dữ liệu nguồn không đổi (tốn nhiều lệnh Sheets).
  const oldCount = Math.max(helper.getLastRow() - 1, 0);
  const oldRows = oldCount ? helper.getRange(2, 1, oldCount, 7).getValues() : [];
  const sameCell = function(a, b) {
    if (a instanceof Date || b instanceof Date) {
      return a instanceof Date && b instanceof Date && a.getTime() === b.getTime();
    }
    return a === b || ((a === '' || a == null) && (b === '' || b == null));
  };
  let unchanged = output.length <= oldRows.length;
  if (unchanged) {
    for (let i = 0; i < oldRows.length && unchanged; i++) {
      for (let j = 0; j < 7; j++) {
        const expected = i < output.length ? output[i][j] : '';
        if (!sameCell(oldRows[i][j], expected)) { unchanged = false; break; }
      }
    }
  }
  if (!unchanged) {
    const clearRows = Math.max(oldCount, 1);
    helper.getRange(2, 1, clearRows, 7).clearContent().clearNote();
    if (output.length) {
      helper.getRange(2, 1, output.length, 7)
        .setValues(output).setFontFamily('Arial');
      helper.getRange(2, 6, output.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    }
  }
  if (helper.getFrozenRows() !== 1) helper.setFrozenRows(1);
  if (!unchanged || headerChanged) SpreadsheetApp.flush();
  return {
    success: true,
    rows: output.length,
    unchanged: unchanged && !headerChanged,
    selectedDate: selectedDate
  };
}
function taiXuongReadDispatch_(sh, selectedDate) {
  const last = sh.getLastRow();
  const result = {
    em: [],
    scc: [],
    kt: [],
    allPlates: new Set(),
    activeVehicleCount: 0,
    byVisitId: {},
    byPlate: {}
  };
  if (last < CFG.FIRST_DATA_ROW) {
    return result;
  }
  const count = last - CFG.FIRST_DATA_ROW + 1;
  const width = CFG.DP.LAST_SCC_STATUS;
  const range = sh.getRange(CFG.FIRST_DATA_ROW, 1, count, width);
  const values = range.getValues();
  const display = range.getDisplayValues();
  const activePlateSet = new Set();
  const timezone = sh.getParent().getSpreadsheetTimeZone();
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    const d = display[i];
    const plate = normalizePlate_(d[CFG.DP.BIEN_SO - 1]);
    if (!plate) continue;
    result.allPlates.add(plate);
    const emRunning = normalizeKey_(d[CFG.DP.TD_EM - 1]) === 'dang lam';
    const sccRunning = normalizeKey_(d[CFG.DP.TD_SCC - 1]) === 'dang lam';
    if (!emRunning && !sccRunning) {
      continue;
    }
    const receiveDate = taiXuongDateOnly_(
      v[CFG.DP.NGAY - 1],
      d[CFG.DP.NGAY - 1]
    );
    const deliveryDate = taiXuongDateOnly_(
      v[CFG.DP.NGAY_GIAO - 1],
      d[CFG.DP.NGAY_GIAO - 1]
    ) || receiveDate;
    const receiveAt = taiXuongCombineDateTime_(
      receiveDate,
      v[CFG.DP.GIO_NHAN - 1],
      d[CFG.DP.GIO_NHAN - 1]
    );
    const deliveryAt = taiXuongCombineDateTime_(
      deliveryDate,
      v[CFG.DP.GIO_GIAO - 1],
      d[CFG.DP.GIO_GIAO - 1]
    );
    const emStamp = taiXuongTimestamp_(v[CFG.DP.MOC_TD_EM - 1]);
    const sccStamp = taiXuongTimestamp_(v[CFG.DP.MOC_TD_SCC - 1]);
    const advisor = clean_(d[CFG.DP.CVDV - 1]);
    const model = clean_(d[CFG.DP.MODEL - 1]);
    const note = clean_(d[CFG.DP.GHI_CHU - 1]);
    const workText = clean_(d[CFG.DP.CV - 1]);
    const bdLevel = clean_(d[CFG.DP.CAP_BD - 1]);
    const gtgt = clean_(d[CFG.DP.GTGT - 1]);
    const emGroup = clean_(d[CFG.DP.NHOM_EM - 1]);
    const inspector = clean_(d[CFG.DP.NGUOI_KT - 1]);
    const sccGroup = clean_(d[CFG.DP.KTV1 - 1]);
    const visitId = clean_(d[CFG.DP.VISIT_ID - 1]);
    const base = {
      row: CFG.FIRST_DATA_ROW + i,
      visitId: visitId,
      plate: plate,
      advisor: advisor,
      model: model,
      note: note,
      workText: workText,
      bdLevel: bdLevel,
      gtgt: gtgt,
      deliveryAt: deliveryAt,
      deliveryDisplay: taiXuongDeliveryDisplay_(deliveryAt, selectedDate, timezone),
      emStatus: clean_(d[CFG.DP.TD_EM - 1]),
      sccStatus: clean_(d[CFG.DP.TD_SCC - 1]),
      emGroup: emGroup,
      sccGroup: sccGroup,
      inspector: inspector,
      priority: normalizeKey_(note).indexOf('xe uu tien') !== -1
    };
    if (visitId) result.byVisitId[visitId] = base;
    result.byPlate[plate] = base;
    let countedActive = false;
    if (emRunning && emGroup) {
      const start = emStamp || receiveAt;
      const span = taiXuongVisibleSpan_(start, deliveryAt, selectedDate);
      if (span) {
        result.em.push(Object.assign({}, base, {
          startAt: span.start,
          drawEndAt: span.end
        }));
        countedActive = true;
      }
    }
    if (sccRunning && sccGroup) {
      const start = sccStamp || receiveAt;
      const span = taiXuongVisibleSpan_(start, deliveryAt, selectedDate);
      if (span) {
        result.scc.push(Object.assign({}, base, {
          startAt: span.start,
          drawEndAt: span.end
        }));
        countedActive = true;
      }
    }
    if (countedActive) {
      activePlateSet.add(plate);
    }
  }
  result.activeVehicleCount = activePlateSet.size;
  return result;
}
function taiXuongReadTracking_(ss, selectedDate, dispatchRecords) {
  const out = [];
  const now = new Date();
  const sources = [CFG.SHEET_KT, CFG.SHEET_KT_LAI];
  sources.forEach(function(sheetName) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    const visitCol = visitColumn_(sheetName);
    const sideCol = visitCol + 1;
    const lastCol = Math.max(8, sideCol);
    const rowCount = sh.getLastRow() - 1;
    const range = sh.getRange(2, 1, rowCount, lastCol);
    const values = range.getValues();
    const display = range.getDisplayValues();
    for (let i = 0; i < values.length; i++) {
      const v = values[i];
      const d = display[i];
      const plate = normalizePlate_(d[1]);
      if (!plate) continue;
      const trackingDate = taiXuongDateOnly_(v[0], d[0]) || selectedDate;
      const startAt = taiXuongTrackingDateTime_(trackingDate, v[6], d[6]);
      if (!startAt) continue;
      let endAt = taiXuongTrackingDateTime_(trackingDate, v[7], d[7]);
      const visitId = clean_(d[visitCol - 1] || v[visitCol - 1]);
      const linked =
        (visitId && dispatchRecords.byVisitId[visitId]) ||
        dispatchRecords.byPlate[plate] ||
        null;
      // Phiên KT còn mở: chỉ kéo tới 30 phút sau hiện tại để biểu thị đang chiếm người,
      // không kéo giả tới toàn bộ giờ giao dự kiến.
      if (!endAt) {
        if (sameDate_(selectedDate, now)) {
          endAt = new Date(Math.max(
            now.getTime() + 30 * 60 * 1000,
            startAt.getTime() + 30 * 60 * 1000
          ));
        }
        else {
          endAt = new Date(startAt.getTime() + 30 * 60 * 1000);
        }
      }
      const span = taiXuongClipSpan_(startAt, endAt, selectedDate);
      if (!span) continue;
      const performer = clean_(d[5]);
      const inspector = linked && linked.inspector
        ? linked.inspector
        : performer;
      if (!inspector) continue;
      out.push({
        row: i + 2,
        visitId: visitId,
        plate: plate,
        advisor: linked && linked.advisor ? linked.advisor : clean_(d[2]),
        model: linked && linked.model ? linked.model : clean_(d[3]),
        note: linked ? linked.note : '',
        workText: clean_(d[4]) || (linked ? linked.workText : ''),
        bdLevel: linked ? linked.bdLevel : '',
        gtgt: linked ? linked.gtgt : '',
        deliveryAt: linked ? linked.deliveryAt : null,
        deliveryDisplay: linked ? linked.deliveryDisplay : '',
        emStatus: linked ? linked.emStatus : '',
        sccStatus: linked ? linked.sccStatus : '',
        emGroup: linked ? linked.emGroup : '',
        sccGroup: linked ? linked.sccGroup : '',
        inspector: inspector,
        priority: linked ? linked.priority : false,
        trackingSheet: sheetName,
        trackingSide: clean_(d[sideCol - 1]),
        trackingClosed: !!taiXuongTrackingDateTime_(trackingDate, v[7], d[7]),
        startAt: span.start,
        drawEndAt: span.end
      });
    }
  });
  return out;
}
function taiXuongTrackingDateTime_(dateValue, timeValue, timeDisplay) {
  // Timestamp đầy đủ do Apps Script ghi.
  if (timeValue instanceof Date && !isNaN(timeValue.getTime())) {
    if (timeValue.getFullYear() >= 2000) {
      return new Date(timeValue.getTime());
    }
  }
  // Trường hợp ô chỉ lưu giờ (Date 1899 / number fraction / chuỗi HH:mm),
  // ghép với ngày của phiên tracking.
  return taiXuongCombineDateTime_(dateValue, timeValue, timeDisplay);
}
function taiXuongClipSpan_(startAt, endAt, selectedDate) {
  if (!(startAt instanceof Date) || !(endAt instanceof Date) || !(selectedDate instanceof Date)) {
    return null;
  }
  const bounds = taiXuongSlotBounds_();
  const firstMinutes = bounds[0].start;
  const lastMinutes = bounds[bounds.length - 1].end;
  const windowStart = new Date(selectedDate.getTime());
  const windowEnd = new Date(selectedDate.getTime());
  windowStart.setHours(Math.floor(firstMinutes / 60), firstMinutes % 60, 0, 0);
  windowEnd.setHours(Math.floor(lastMinutes / 60), lastMinutes % 60, 0, 0);
  let start = new Date(startAt.getTime());
  let end = new Date(endAt.getTime());
  if (end.getTime() <= start.getTime()) {
    end = new Date(start.getTime() + 30 * 60 * 1000);
  }
  if (end.getTime() <= windowStart.getTime() || start.getTime() >= windowEnd.getTime()) {
    return null;
  }
  if (start.getTime() < windowStart.getTime()) start = windowStart;
  if (end.getTime() > windowEnd.getTime()) end = windowEnd;
  return {start: start, end: end};
}
function taiXuongReadAppointments_(ss, selectedDate, dispatchPlateSet) {
  const sh = ss.getSheetByName(TAI_XUONG_CFG.APPOINTMENT_SHEET);
  const result = {
    total: 0,
    slots: TAI_XUONG_CFG.TIMES.map(function() { return []; })
  };
  if (!sh || sh.getLastRow() < 2) {
    return result;
  }
  const last = sh.getLastRow();
  const values = sh.getRange(2, 1, last - 1, 7).getValues();
  const display = sh.getRange(2, 1, last - 1, 7).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    const d = display[i];
    const status = normalizeKey_(d[6] || v[6]);
    if (status !== 'giu hen') continue;
    const plate = normalizePlate_(d[0] || v[0]);
    if (!plate) continue;
    if (dispatchPlateSet && dispatchPlateSet.has(plate)) continue;
    const appointmentAt = taiXuongAppointmentDateTime_(v[5], d[5], selectedDate);
    if (!appointmentAt || !sameDate_(appointmentAt, selectedDate)) continue;
    const slotIndex = taiXuongAppointmentSlotIndex_(appointmentAt);
    if (slotIndex < 0) continue;
    result.slots[slotIndex].push({
      plate: plate,
      advisor: clean_(d[1]),
      work: clean_(d[4]),
      time: Utilities.formatDate(
        appointmentAt,
        ss.getSpreadsheetTimeZone(),
        'HH:mm'
      )
    });
    result.total++;
  }
  return result;
}
function taiXuongBuildGroups_(groupNames, records, key) {
  const out = [];
  const byKey = {};
  groupNames.forEach(function(name) {
    byKey[normalizeKey_(name)] = [];
  });
  records.forEach(function(rec) {
    const groupKey = normalizeKey_(rec[key]);
    if (Object.prototype.hasOwnProperty.call(byKey, groupKey)) {
      byKey[groupKey].push(rec);
    }
  });
  groupNames.forEach(function(name) {
    const items = byKey[normalizeKey_(name)] || [];
    out.push({
      name: name,
      lanes: taiXuongAllocateLanes_(items)
    });
  });
  return out;
}
function taiXuongAllocateLanes_(records) {
  const sorted = records.slice().sort(function(a, b) {
    const diff = a.startAt.getTime() - b.startAt.getTime();
    if (diff !== 0) return diff;
    return a.drawEndAt.getTime() - b.drawEndAt.getTime();
  });
  const lanes = [];
  const laneEnds = [];
  sorted.forEach(function(rec) {
    let laneIndex = -1;
    for (let i = 0; i < laneEnds.length; i++) {
      if (laneEnds[i] <= rec.startAt.getTime()) {
        laneIndex = i;
        break;
      }
    }
    if (laneIndex < 0) {
      laneIndex = lanes.length;
      lanes.push([]);
      laneEnds.push(0);
    }
    lanes[laneIndex].push(rec);
    laneEnds[laneIndex] = rec.drawEndAt.getTime();
  });
  if (lanes.length === 0) {
    lanes.push([]);
  }
  return lanes;
}
function taiXuongBuildLayout_(emGroups, sccGroups, ktGroups) {
  let row = 7;
  const em = {
    titleRow: row,
    firstGroupRow: row + 1
  };
  row++;
  emGroups.forEach(function(group) {
    group.startRow = row;
    group.rowCount = Math.max(1, group.lanes.length);
    row += group.rowCount;
  });
  em.lastRow = row - 1;
  row += 1; // dòng trống
  const scc = {
    titleRow: row,
    firstGroupRow: row + 1
  };
  row++;
  sccGroups.forEach(function(group) {
    group.startRow = row;
    group.rowCount = Math.max(1, group.lanes.length);
    row += group.rowCount;
  });
  scc.lastRow = row - 1;
  row += 1;
  const kt = {
    titleRow: row,
    firstGroupRow: row + 1
  };
  row++;
  ktGroups.forEach(function(group) {
    group.startRow = row;
    group.rowCount = Math.max(1, group.lanes.length);
    row += group.rowCount;
  });
  kt.lastRow = row - 1;
  return {
    em: em,
    scc: scc,
    kt: kt,
    lastRow: Math.max(row - 1, 26)
  };
}
function taiXuongPrepareSheet_(sh, lastRow) {
  const lastCol = TAI_XUONG_CFG.FIRST_TIME_COL + TAI_XUONG_CFG.TIMES.length - 1;
  const clearRows = Math.max(sh.getLastRow(), lastRow, 40);
  if (sh.getMaxRows() < clearRows) {
    sh.insertRowsAfter(
      sh.getMaxRows(),
      clearRows - sh.getMaxRows()
    );
  }
  if (sh.getMaxColumns() < lastCol) {
    sh.insertColumnsAfter(
      sh.getMaxColumns(),
      lastCol - sh.getMaxColumns()
    );
  }
  const dashboardRange = sh.getRange(1, 1, clearRows, lastCol);
  dashboardRange.breakApart();
  dashboardRange.clear();
  // clear() không phải lúc nào cũng dọn Note cũ theo cách người dùng mong đợi.
  // Xóa Note rõ ràng để ô từng là "khung giờ hiện tại" không giữ tooltip ở lần render sau.
  dashboardRange.clearNote();
  sh.setConditionalFormatRules([]);
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(4);
  sh.setFrozenColumns(1);
  sh.setColumnWidth(1, TAI_XUONG_CFG.LABEL_COL_WIDTH);
  sh.setColumnWidths(TAI_XUONG_CFG.FIRST_TIME_COL,
    lastCol - TAI_XUONG_CFG.FIRST_TIME_COL + 1,
    TAI_XUONG_CFG.TIME_COL_WIDTH);
  sh.setRowHeight(1, 30);
  sh.setRowHeight(2, 28);
  sh.setRowHeight(4, 28);
  sh.setRowHeight(5, 28);
  sh
    .getRange(1, 1, clearRows, lastCol)
    .setFontFamily('Arial')
    .setFontColor(TAI_XUONG_CFG.COLORS.TEXT)
    .setVerticalAlignment('middle');
}
function taiXuongRenderHeader_(sh, selectedDate, selectedDateFormula, appointments, records, lastRow) {
  const lastCol = TAI_XUONG_CFG.FIRST_TIME_COL + TAI_XUONG_CFG.TIMES.length - 1;
  const ss = sh.getParent();
  const tz = ss.getSpreadsheetTimeZone();
  // Cột A được cố định. Google Sheets không cho merge qua ranh giới
  // giữa cột cố định và cột không cố định, nên chỉ merge phần timeline B:lastCol.
  sh
    .getRange(1, 1, 1, lastCol)
    .setBackground(TAI_XUONG_CFG.COLORS.TITLE_BG)
    .setFontColor(TAI_XUONG_CFG.COLORS.TITLE_TEXT);
  if (lastCol >= 2) {
    sh.getRange(1, 2, 1, lastCol - 1).merge();
  }
  sh
    .getRange(1, lastCol >= 2 ? 2 : 1)
    .setValue('THÔNG TIN TẢI XƯỞNG')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center');
  sh.getRange('A2').setValue('Ngày làm việc').setFontWeight('bold');
  // C2 cố ý giữ làm ô ngày để tương thích công thức Data_LichHen hiện hữu.
  sh.getRange('B2').clearContent();
  const dateCell = sh.getRange('C2');
  if (selectedDateFormula) {
    dateCell.setFormula(selectedDateFormula);
  }
  else {
    dateCell.setValue(selectedDate);
  }
  dateCell
    .setNumberFormat('dd/MM/yyyy')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sh.getRange('D2').setValue('Tổng xe hẹn').setFontWeight('bold');
  sh.getRange('E2').setValue(appointments.total).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('F2').setValue('Xe đang hoạt động').setFontWeight('bold');
  sh.getRange('G2').setValue(records.activeVehicleCount).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(4, 1).setValue('Khu vực / Nhóm').setFontWeight('bold');
  const headerValues = [TAI_XUONG_CFG.TIMES];
  sh
    .getRange(4, TAI_XUONG_CFG.FIRST_TIME_COL, 1, TAI_XUONG_CFG.TIMES.length)
    .setValues(headerValues)
    .setBackground(TAI_XUONG_CFG.COLORS.TIME_BG)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sh
    .getRange(4, 1, 1, lastCol)
    .setBorder(
      true, true, true, true, true, true,
      TAI_XUONG_CFG.COLORS.GRID,
      SpreadsheetApp.BorderStyle.SOLID
    );
  sh
    .getRange(5, 1)
    .setValue('XE HẸN')
    .setFontWeight('bold')
    .setBackground(TAI_XUONG_CFG.COLORS.APPT_BG);
  let maxAppointmentInSlot = 0;
  for (let i = 0; i < appointments.slots.length; i++) {
    const list = appointments.slots[i];
    const cell = sh.getRange(5, TAI_XUONG_CFG.FIRST_TIME_COL + i);
    maxAppointmentInSlot = Math.max(maxAppointmentInSlot, list.length);
    let displayText = '';
    if (list.length === 1) {
      displayText = list[0].plate + '\n' + list[0].time;
    }
    else if (list.length > 1) {
      displayText = list.length + ' XE\n' +
        list.slice(0, 3).map(function(item) { return item.plate; }).join('\n');
      if (list.length > 3) displayText += '\n+' + (list.length - 3);
    }
    cell
      .setBackground(list.length ? TAI_XUONG_CFG.COLORS.APPT_HIT_BG : TAI_XUONG_CFG.COLORS.APPT_BG)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true)
      .setFontFamily('Arial')
      .setFontSize(8)
      .setFontWeight(list.length ? 'bold' : 'normal')
      .setValue(displayText);
    if (list.length) {
      cell.setNote(
        list.map(function(item) {
          return (
            item.time + ' • ' + item.plate +
            (item.advisor ? ' • ' + item.advisor : '') +
            (item.work ? '\nViệc: ' + item.work : '')
          );
        }).join('\n\n')
      );
    }
  }
  sh.setRowHeight(5, maxAppointmentInSlot ? Math.min(82, 38 + Math.min(maxAppointmentInSlot, 3) * 12) : 28);
  sh
    .getRange(5, 1, 1, lastCol)
    .setBorder(
      true, true, true, true, true, true,
      TAI_XUONG_CFG.COLORS.GRID,
      SpreadsheetApp.BorderStyle.SOLID
    );
  // H2:T2 = bảng chú giải màu + ký hiệu.
  // Màu nền block thể hiện khu vực/trạng thái; ★ chỉ xe có ghi chú "Xe ưu tiên".
  const legend = [
    {col: 8,  text: 'EM',        bg: TAI_XUONG_CFG.COLORS.EM_BLOCK,      border: TAI_XUONG_CFG.COLORS.EM_BORDER},
    {col: 9,  text: 'SCC',       bg: TAI_XUONG_CFG.COLORS.SCC_BLOCK,     border: TAI_XUONG_CFG.COLORS.SCC_BORDER},
    {col: 10, text: 'Người KT',  bg: TAI_XUONG_CFG.COLORS.KT_BLOCK,      border: TAI_XUONG_CFG.COLORS.KT_BORDER},
    {col: 11, text: 'Sắp giao',  bg: TAI_XUONG_CFG.COLORS.SOON,          border: TAI_XUONG_CFG.COLORS.SOON_BORDER},
    {col: 12, text: 'Trễ giao',  bg: TAI_XUONG_CFG.COLORS.OVERDUE,       border: TAI_XUONG_CFG.COLORS.OVERDUE_BORDER},
    {col: 13, text: 'Xe hẹn',    bg: TAI_XUONG_CFG.COLORS.APPT_HIT_BG,   border: TAI_XUONG_CFG.COLORS.SOON_BORDER},
    {col: 14, text: '★ Ưu tiên', bg: '#FFFFFF',                           border: '#777777'}
  ];
  legend.forEach(function(item) {
    if (item.col > lastCol) return;
    sh
      .getRange(2, item.col)
      .setValue(item.text)
      .setBackground(item.bg)
      .setFontFamily('Arial')
      .setFontSize(8)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(
        true, true, true, true, false, false,
        item.border,
        SpreadsheetApp.BorderStyle.SOLID
      );
  });
  // O2:T2 dành cho mốc cập nhật.
  const stampCol = 15;
  if (stampCol <= lastCol) {
    if (lastCol > stampCol) {
      sh.getRange(2, stampCol, 1, lastCol - stampCol + 1).merge();
    }
    sh
      .getRange(2, stampCol)
      .setValue(
        'Cập nhật: ' +
        Utilities.formatDate(new Date(), tz, 'HH:mm dd/MM/yyyy')
      )
      .setFontFamily('Arial')
      .setFontSize(8)
      .setFontColor(TAI_XUONG_CFG.COLORS.MUTED)
      .setHorizontalAlignment('right');
  }
  // Khung chung vùng timeline.
  if (lastRow >= 4) {
    sh
      .getRange(4, 1, lastRow - 3, lastCol)
      .setVerticalAlignment('middle');
  }
}
function taiXuongRenderSection_(sh, section, title, groups, type) {
  const lastCol = TAI_XUONG_CFG.FIRST_TIME_COL + TAI_XUONG_CFG.TIMES.length - 1;
  const palette = taiXuongPalette_(type);
  sh
    .getRange(section.titleRow, 1, 1, lastCol)
    .setBackground(palette.section)
    .setFontWeight('bold')
    .setBorder(
      true, true, true, true, false, false,
      palette.border,
      SpreadsheetApp.BorderStyle.SOLID
    );
  sh.getRange(section.titleRow, 1).setValue(title);
  sh.setRowHeight(section.titleRow, 26);
  groups.forEach(function(group) {
    const startRow = group.startRow;
    const rowCount = group.rowCount;
    // P0.5.2: một lần chỉnh chiều cao và đường kẻ cho toàn bộ lane của khu vực.
    sh.setRowHeights(startRow, rowCount, TAI_XUONG_CFG.MIN_LANE_HEIGHT);
    sh.getRange(startRow, TAI_XUONG_CFG.FIRST_TIME_COL,
      rowCount, TAI_XUONG_CFG.TIMES.length)
      .setBackground(TAI_XUONG_CFG.COLORS.EMPTY)
      .setBorder(true, true, true, true, true, true,
        TAI_XUONG_CFG.COLORS.GRID, SpreadsheetApp.BorderStyle.SOLID);
    const labelRange = sh.getRange(startRow, 1, rowCount, 1);
    if (rowCount > 1) {
      labelRange.merge();
    }
    sh
      .getRange(startRow, 1)
      .setValue(group.name)
      .setFontWeight('bold')
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setBackground(palette.section)
      .setBorder(
        true, true, true, true, false, false,
        palette.border,
        SpreadsheetApp.BorderStyle.SOLID
      );
    group.lanes.forEach(function(lane, laneIndex) {
      const row = startRow + laneIndex;
      lane.forEach(function(rec) {
        taiXuongRenderBlock_(sh, row, rec, type);
      });
    });
  });
}
function taiXuongRenderBlock_(sh, row, rec, type) {
  const cols = taiXuongBlockColumns_(rec.startAt, rec.drawEndAt);
  if (!cols) return;
  const startCol = TAI_XUONG_CFG.FIRST_TIME_COL + cols.startIndex;
  const width = cols.endIndex - cols.startIndex + 1;
  const blockRange = sh.getRange(row, startCol, 1, width);
  if (width > 1) {
    blockRange.merge();
  }
  const now = new Date();
  const deliveryAt = rec.deliveryAt;
  let state = 'normal';
  if (deliveryAt instanceof Date && !isNaN(deliveryAt.getTime())) {
    const diff = deliveryAt.getTime() - now.getTime();
    if (diff < 0) {
      state = 'overdue';
    }
    else if (diff <= 60 * 60 * 1000) {
      state = 'soon';
    }
  }
  const palette = taiXuongPalette_(type);
  const bg = state === 'overdue'
    ? TAI_XUONG_CFG.COLORS.OVERDUE
    : state === 'soon'
      ? TAI_XUONG_CFG.COLORS.SOON
      : palette.block;
  const border = state === 'overdue'
    ? TAI_XUONG_CFG.COLORS.OVERDUE_BORDER
    : state === 'soon'
      ? TAI_XUONG_CFG.COLORS.SOON_BORDER
      : palette.border;
  const plateLine = (rec.priority ? '★ ' : '') + rec.plate;
  const advisorLine = 'CVDV: ' + (rec.advisor || '—');
  const workLine = 'Việc: ' + taiXuongCompactWork_(rec.workText || '—', 44);
  const serviceParts = [];
  if (clean_(rec.bdLevel)) serviceParts.push('BD: ' + clean_(rec.bdLevel));
  if (clean_(rec.gtgt)) serviceParts.push('GTGT: ' + clean_(rec.gtgt));
  const serviceLine = serviceParts.length
    ? taiXuongCompactWork_(serviceParts.join(' | '), 52)
    : '';
  const deliveryLine = 'Giao: ' + (rec.deliveryDisplay || '—');
  const displayLines = [plateLine, advisorLine, workLine];
  if (serviceLine) displayLines.push(serviceLine);
  displayLines.push(deliveryLine);
  const text = displayLines.join('\n');
  const textBuilder = SpreadsheetApp.newRichTextValue().setText(text);
  textBuilder.setTextStyle(
    0,
    text.length,
    SpreadsheetApp
      .newTextStyle()
      .setFontFamily('Arial')
      .setFontSize(8)
      .build()
  );
  textBuilder.setTextStyle(
    0,
    plateLine.length,
    SpreadsheetApp
      .newTextStyle()
      .setFontFamily('Arial')
      .setBold(true)
      .setFontSize(9)
      .build()
  );
  // Các dòng chi tiết giữ Arial 8; dòng Giao in đậm.
  const deliveryStart = text.lastIndexOf(deliveryLine);
  if (deliveryStart >= 0) {
    textBuilder.setTextStyle(
      deliveryStart,
      text.length,
      SpreadsheetApp
        .newTextStyle()
        .setFontFamily('Arial')
        .setFontSize(8)
        .setBold(true)
        .build()
    );
  }
  sh.getRange(row, startCol).setRichTextValue(textBuilder.build());
  blockRange
    .setBackground(bg)
    .setFontFamily('Arial')
    .setFontColor(TAI_XUONG_CFG.COLORS.TEXT)
    .setWrap(true)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setBorder(
      true, true, true, true, false, false,
      border,
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
  sh.getRange(row, startCol).setNote(
    'Biển số: ' + rec.plate +
    '\nCVDV: ' + (rec.advisor || '—') +
    '\nModel: ' + (rec.model || '—') +
    '\nCông việc: ' + (rec.workText || '—') +
    '\nCấp bảo dưỡng: ' + (rec.bdLevel || '—') +
    '\nGTGT: ' + (rec.gtgt || '—') +
    '\nGiờ giao dự kiến: ' + (rec.deliveryDisplay || '—') +
    '\nTiến độ EM: ' + (rec.emStatus || '—') +
    '\nTiến độ SCC: ' + (rec.sccStatus || '—')
  );
}
function taiXuongLockArial_(sh, lastRow) {
  const lastCol =
    TAI_XUONG_CFG.FIRST_TIME_COL +
    TAI_XUONG_CFG.TIMES.length -
    1;
  const rows = Math.max(5, Number(lastRow || 5));
  sh
    .getRange(1, 1, rows, lastCol)
    .setFontFamily('Arial');
}
function taiXuongCompactWork_(value, maxLength) {
  const text = clean_(value).replace(/\s+/g, ' ');
  const limit = Math.max(12, Number(maxLength || 36));
  if (!text) return '';
  if (text.length <= limit) return text;
  return text.substring(0, limit - 1).trim() + '…';
}
function taiXuongPalette_(type) {
  if (type === 'EM') {
    return {
      section: TAI_XUONG_CFG.COLORS.EM_SECTION,
      block: TAI_XUONG_CFG.COLORS.EM_BLOCK,
      border: TAI_XUONG_CFG.COLORS.EM_BORDER
    };
  }
  if (type === 'SCC') {
    return {
      section: TAI_XUONG_CFG.COLORS.SCC_SECTION,
      block: TAI_XUONG_CFG.COLORS.SCC_BLOCK,
      border: TAI_XUONG_CFG.COLORS.SCC_BORDER
    };
  }
  return {
    section: TAI_XUONG_CFG.COLORS.KT_SECTION,
    block: TAI_XUONG_CFG.COLORS.KT_BLOCK,
    border: TAI_XUONG_CFG.COLORS.KT_BORDER
  };
}
function taiXuongBlockColumns_(startAt, endAt) {
  if (!(startAt instanceof Date) || !(endAt instanceof Date)) return null;
  const startMinutes = startAt.getHours() * 60 + startAt.getMinutes();
  const endMinutes = endAt.getHours() * 60 + endAt.getMinutes();
  const bounds = taiXuongSlotBounds_();
  let startIndex = -1;
  let endIndex = -1;
  for (let i = 0; i < bounds.length; i++) {
    const intersects = endMinutes > bounds[i].start && startMinutes < bounds[i].end;
    if (!intersects) continue;
    if (startIndex < 0) startIndex = i;
    endIndex = i;
  }
  if (startIndex < 0 || endIndex < 0) return null;
  return {
    startIndex: startIndex,
    endIndex: endIndex
  };
}
function taiXuongVisibleSpan_(startAt, deliveryAt, selectedDate) {
  if (!(selectedDate instanceof Date)) return null;
  const dayStart = new Date(selectedDate.getTime());
  const bounds = taiXuongSlotBounds_();
  const firstMinutes = bounds[0].start;
  const lastMinutes = bounds[bounds.length - 1].end;
  const windowStart = new Date(dayStart.getTime());
  windowStart.setHours(Math.floor(firstMinutes / 60), firstMinutes % 60, 0, 0);
  const windowEnd = new Date(dayStart.getTime());
  windowEnd.setHours(Math.floor(lastMinutes / 60), lastMinutes % 60, 0, 0);
  let start = startAt instanceof Date && !isNaN(startAt.getTime())
    ? new Date(startAt.getTime())
    : new Date(windowStart.getTime());
  let end = deliveryAt instanceof Date && !isNaN(deliveryAt.getTime())
    ? new Date(deliveryAt.getTime())
    : new Date(start.getTime() + 30 * 60 * 1000);
  // Xe vẫn Đang làm nhưng đã trễ giờ giao: kéo block tới ít nhất 30 phút sau hiện tại
  // để dashboard vẫn thể hiện tài nguyên đang bị chiếm.
  const now = new Date();
  if (sameDate_(selectedDate, now) && end.getTime() < now.getTime()) {
    end = new Date(now.getTime() + 30 * 60 * 1000);
  }
  if (end.getTime() <= start.getTime()) {
    end = new Date(start.getTime() + 30 * 60 * 1000);
  }
  if (end.getTime() <= windowStart.getTime() || start.getTime() >= windowEnd.getTime()) {
    return null;
  }
  if (start.getTime() < windowStart.getTime()) start = windowStart;
  if (end.getTime() > windowEnd.getTime()) end = windowEnd;
  return {
    start: start,
    end: end
  };
}
function taiXuongSlotBounds_() {
  const minutes = TAI_XUONG_CFG.TIMES.map(function(text) {
    const parts = text.split(':');
    return Number(parts[0]) * 60 + Number(parts[1]);
  });
  return minutes.map(function(start, index) {
    const end = index + 1 < minutes.length
      ? minutes[index + 1]
      : start + TAI_XUONG_CFG.LAST_SLOT_MINUTES;
    return {
      start: start,
      end: end
    };
  });
}
function taiXuongAppointmentSlotIndex_(appointmentAt) {
  if (!(appointmentAt instanceof Date) || isNaN(appointmentAt.getTime())) {
    return -1;
  }
  const minutes = appointmentAt.getHours() * 60 + appointmentAt.getMinutes();
  const bounds = taiXuongSlotBounds_();
  /*********************************************************
   * Slot dùng quy tắc [start, end):
   * - 13:00 / 13:10 / 13:15 / 13:29 -> slot 13:00
   * - 13:30 / 13:45 / 13:59         -> slot 13:30
   * - 14:00                          -> slot 14:00
   *
   * Riêng khoảng nghỉ trưa, vì timeline không có 12:30 nên slot 12:00
   * vẫn kéo tới trước 13:00 theo đúng cấu hình TIMES hiện tại.
   *********************************************************/
  for (let i = 0; i < bounds.length; i++) {
    if (minutes >= bounds[i].start && minutes < bounds[i].end) {
      return i;
    }
  }
  return -1;
}
function taiXuongDateOnly_(value, displayValue) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return stripTime_(value);
  }
  const text = clean_(displayValue || value);
  const m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  d.setHours(0, 0, 0, 0);
  return d;
}
function taiXuongTimeParts_(value, displayValue) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return {
      hour: value.getHours(),
      minute: value.getMinutes()
    };
  }
  if (typeof value === 'number' && isFinite(value)) {
    let total = Math.round(value * 24 * 60) % (24 * 60);
    if (total < 0) total += 24 * 60;
    return {
      hour: Math.floor(total / 60),
      minute: total % 60
    };
  }
  const text = clean_(displayValue || value);
  const m = text.match(/^(\d{1,2}):(\d{2})/);
  if (!m) return null;
  const hour = Number(m[1]);
  const minute = Number(m[2]);
  if (hour < 0 || hour > 23 || minute < 0 || minute > 59) return null;
  return {
    hour: hour,
    minute: minute
  };
}
function taiXuongCombineDateTime_(dateValue, timeValue, timeDisplay) {
  if (!(dateValue instanceof Date) || isNaN(dateValue.getTime())) {
    return null;
  }
  const parts = taiXuongTimeParts_(timeValue, timeDisplay);
  if (!parts) return null;
  const out = stripTime_(dateValue);
  out.setHours(parts.hour, parts.minute, 0, 0);
  return out;
}
function taiXuongTimestamp_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getTime());
  }
  // IMPORTRANGE đôi lúc trả serial number thay vì Date object.
  if (typeof value === 'number' && isFinite(value)) {
    const wholeDays = Math.floor(value);
    const fraction = value - wholeDays;
    const out = new Date(1899, 11, 30);
    out.setDate(out.getDate() + wholeDays);
    let totalMinutes = Math.round(fraction * 24 * 60);
    if (totalMinutes >= 24 * 60) {
      out.setDate(out.getDate() + 1);
      totalMinutes -= 24 * 60;
    }
    out.setHours(
      Math.floor(totalMinutes / 60),
      totalMinutes % 60,
      0,
      0
    );
    return isNaN(out.getTime()) ? null : out;
  }
  return null;
}
function taiXuongAppointmentDateTime_(value, displayValue, selectedDate) {
  /*********************************************************
   * V13.12.6 - ƯU TIÊN CHUỖI HIỂN THỊ CỦA GOOGLE SHEETS
   *
   * Với dữ liệu lịch hẹn từ một Spreadsheet khác, Date object có thể
   * chịu ảnh hưởng timezone của file nguồn / project khi đọc bằng Apps Script.
   * Chuỗi hiển thị cột M (dd/MM/yyyy HH:mm) là thứ người dùng đang thấy
   * và là nguồn chính xác nhất để xếp xe vào slot 30 phút.
   *********************************************************/
  let out =
    taiXuongParseDateTimeDisplay_(displayValue) ||
    taiXuongParseDateTimeDisplay_(value);
  if (out) return out;
  out = taiXuongTimestamp_(value);
  // Serial/time-only Date 1899: ghép giờ vào ngày đang xem.
  if (out && out.getFullYear() < 2000 && selectedDate instanceof Date) {
    const rebuilt = new Date(selectedDate.getTime());
    rebuilt.setHours(out.getHours(), out.getMinutes(), 0, 0);
    out = rebuilt;
  }
  if (out) return out;
  // Trường hợp chỉ còn HH:mm: ghép với ngày đang xem.
  const parts = taiXuongTimeParts_(value, displayValue);
  if (parts && selectedDate instanceof Date) {
    const rebuilt = new Date(selectedDate.getTime());
    rebuilt.setHours(parts.hour, parts.minute, 0, 0);
    return rebuilt;
  }
  return null;
}
function taiXuongParseDateTimeDisplay_(value) {
  const text = clean_(value);
  if (!text) return null;
  // Hỗ trợ dd/MM/yyyy HH:mm hoặc dd/MM/yyyy H:mm.
  const m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+)(\d{1,2}):(\d{2})/);
  if (!m) return null;
  const d = new Date(
    Number(m[3]),
    Number(m[2]) - 1,
    Number(m[1]),
    Number(m[4]),
    Number(m[5]),
    0,
    0
  );
  return isNaN(d.getTime()) ? null : d;
}
function taiXuongDeliveryDisplay_(deliveryAt, selectedDate, timezone) {
  if (!(deliveryAt instanceof Date) || isNaN(deliveryAt.getTime())) {
    return '';
  }
  // Chuyển timezone từ nguồn đã mở: không openById trên từng xe.
  const tz = timezone || laySpreadsheetWeb_().getSpreadsheetTimeZone();
  return sameDate_(deliveryAt, selectedDate)
    ? Utilities.formatDate(deliveryAt, tz, 'HH:mm')
    : Utilities.formatDate(deliveryAt, tz, 'dd/MM HH:mm');
}
function taiXuongMarkCurrentTime_(sh, selectedDate, lastRow) {
  // Dọn mọi Note cũ trên header thời gian trước khi đánh dấu slot hiện tại mới.
  sh
    .getRange(4, TAI_XUONG_CFG.FIRST_TIME_COL, 1, TAI_XUONG_CFG.TIMES.length)
    .clearNote();
  const now = new Date();
  if (!sameDate_(selectedDate, now)) return;
  const currentMinutes = now.getHours() * 60 + now.getMinutes();
  const bounds = taiXuongSlotBounds_();
  let currentIndex = -1;
  for (let i = 0; i < bounds.length; i++) {
    if (currentMinutes >= bounds[i].start && currentMinutes < bounds[i].end) {
      currentIndex = i;
      break;
    }
  }
  if (currentIndex < 0) return;
  const col = TAI_XUONG_CFG.FIRST_TIME_COL + currentIndex;
  const tz = sh.getParent().getSpreadsheetTimeZone();
  sh
    .getRange(4, col)
    .setFontColor(TAI_XUONG_CFG.COLORS.NOW)
    .setFontWeight('bold')
    .setNote('Hiện tại: ' + Utilities.formatDate(now, tz, 'HH:mm'));
  // Vẽ vạch đỏ dọc tại đầu khung giờ hiện tại.
  // Block xe đang merge có thể che vạch; chủ động bỏ qua cell thuộc merge để tránh lỗi
  // "phải chọn toàn bộ vùng hợp nhất". Vạch sẽ chạy qua các khoảng trống và header,
  // tạo cảm giác line hiện tại nằm phía sau card xe.
  const colRange = sh.getRange(4, col, lastRow - 3, 1);
  const mergedRows = Object.create(null);
  colRange.getMergedRanges().forEach(function(merged) {
    const a = Math.max(4, merged.getRow());
    const b = Math.min(lastRow, merged.getLastRow());
    for (let row = a; row <= b; row++) mergedRows[row] = true;
  });
  let segmentStart = 0;
  for (let row = 4; row <= lastRow + 1; row++) {
    const blocked = row > lastRow || mergedRows[row] === true;
    if (blocked) {
      if (segmentStart) {
        sh.getRange(segmentStart, col, row - segmentStart, 1)
          .setBorder(null, true, null, null, null, null,
            TAI_XUONG_CFG.COLORS.NOW, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        segmentStart = 0;
      }
    } else if (!segmentStart) segmentStart = row;
  }
}


/*********************************************************
 * V13.18 - STABLE EVENT CONTRACT FOR LSC/WEBAPP
 * Gọi từ WebApp sau khi ghi/xử lý xong một dòng Điều phối.
 * Fail-soft ở caller: correction lỗi không được rollback lượt xe LSC.
 *********************************************************/


/*********************************************************
 * P0.3 - AUDIT PHYSICAL VISIT (READ ONLY)
 * Không xóa dòng, không sửa giá trị hay visitId.
 * Mục đích: lập chứng cứ trước 1 lần sửa dữ liệu lịch sử.
 *********************************************************/


/*********************************************************
 * P0.3 - ONE-TIME 52LD4526 PHYSICAL-VISIT REPAIR
 * CHỈ chạy có chủ đích, KHÔNG do trigger gọi.
 * Trước khi chạy: người quản trị phải tạo bản sao file Spreadsheet.
 * Giữ Data SCC cả 2 visitId; bản snapshot dòng Chờ giao cũ nằm ở
 * sheet P03_AUDIT_LUOT_XE trước khi compact bằng hàm hiện hữu.
 * Mọi điều kiện chưa chắc chắn -> không thay đổi trạng thái nghiệp vụ.
 *********************************************************/




/*********************************************************
 * P0.4 - PHYSICAL VISIT PRE-COMMIT GUARD
 * Chặn trước mọi đồng bộ tiến độ khi xe có 2 lượt mở.
 * Không tự xóa/gộp visitId và không kết luận từ biển số đơn độc.
 * Xe quay lại chỉ mở lượt sau khi lượt cũ được đóng hợp lệ.
 *********************************************************/
function p04AssertPhysicalVisitOwnership_(sh, row, plate) {
  if (!sh || sh.getName() !== CFG.SHEET_DISPATCH) return true;
  const ss = sh.getParent();
  const target = normalizePlate_(plate);
  if (!target) return true;
  const id = clean_(sh.getRange(row, CFG.DP.VISIT_ID).getDisplayValue());
  const first = CFG.FIRST_DATA_ROW;
  const last = sh.getLastRow();
  const matches = [];
  if (last >= first) {
    // Một range B:V để giảm số lần gọi Sheets trong worker.
    const values = sh.getRange(first, CFG.DP.BIEN_SO, last-first+1,
      CFG.DP.VISIT_ID-CFG.DP.BIEN_SO+1).getDisplayValues();
    values.forEach(function(r, i) {
      const at = first+i;
      if (at === row || normalizePlate_(r[0]) !== target) return;
      matches.push({row:at, visitId:clean_(r[CFG.DP.VISIT_ID-CFG.DP.BIEN_SO])});
    });
  }
  if (matches.length) {
    throw new Error('P04_PHYSICAL_DUPLICATE_DISPATCH: ' + target +
      ' đang có thêm dòng Điều phối ' +
      matches.map(function(x){ return x.row+':'+x.visitId; }).join(',') +
      '. Giữ nguyên dữ liệu, không tự gộp visitId.');
  }
  const wd = ss.getSheetByName(CFG.SHEET_WAITING_DELIVERY);
  if (!wd) throw new Error('P04_WAITING_SHEET_MISSING');
  const wdFirst = CFG.WAITING_DELIVERY_FIRST_ROW;
  const wdLast = wd.getLastRow();
  const conflicts = [];
  if (wdLast >= wdFirst) {
    // Một range B:H, chỉ các lượt đang Chờ giao mới chặn.
    wd.getRange(wdFirst, CFG.WD.BIEN_SO, wdLast-wdFirst+1,
      CFG.WD.VISIT_ID-CFG.WD.BIEN_SO+1).getDisplayValues().forEach(function(r,i) {
      if (normalizePlate_(r[0])!==target ||
          normalizeKey_(r[CFG.WD.TRANG_THAI-CFG.WD.BIEN_SO])!=='cho giao') return;
      const owner=clean_(r[CFG.WD.VISIT_ID-CFG.WD.BIEN_SO]);
      if (!id || owner!==id) conflicts.push(wdFirst+i);
    });
  }
  if (conflicts.length) {
    throw new Error('P04_PHYSICAL_VISIT_UNCLOSED: ' + target +
      ' Điều phối dòng ' + row + ' visitId=' + (id||'<chưa cấp>') +
      ' nhưng còn Chờ giao tại dòng ' + conflicts.join(',') +
      '. Chưa có bằng chứng lượt cũ đã đóng. Ngừng đồng bộ lượt này.');
  }
  return true;
}

/**
 * P0.4 - Chỉ đọc/không ghi. Thẩm định một trường hợp xe ở cả 2 danh sách.
 * Không suy diễn hoàn thành công việc từ trạng thái sheet Điều phối.
 */


/**
 * P0.4 - Sửa một lần 51H41700 khi chứng cứ vẫn khớp.
 * Chỉ dọn bản ghi Chờ giao cũ bị công bố sớm; giữ nguyên Data SCC.
 * Không được gọi từ trigger; audit snapshot bắt buộc trước ghi.
 */



/**
 * P0.4.3: READ-ONLY Production QA. Run once from Apps Script UI after
 * installing this candidate. Returns evidence, never mutates a workbook.
 */
function p043ProductionPreflightReadOnly() {
  const result = {
    version: 'TPMH P0.4.5 Controlled Reopen + Data Capacity CANDIDATE',
    observedAt: new Date().toISOString(),
    dispatchRows: 0,
    waitingRows: 0,
    pendingQueueCount: 0,
    workerTriggerCount: 0,
    monitor: [],
    differences: [],
    warnings: []
  };
  const ss = laySpreadsheetWeb_();
  if (!ss) throw new Error('P043_PREFLIGHT_SPREADSHEET_UNAVAILABLE');
  const dp = ss.getSheetByName(CFG.SHEET_DISPATCH);
  const wd = ss.getSheetByName(CFG.SHEET_WAITING_DELIVERY);
  const em = ss.getSheetByName(CFG.SHEET_DATA_EM);
  const scc = ss.getSheetByName(CFG.SHEET_DATA_SCC);
  const mon = ss.getSheetByName('P04_GIAM_SAT_AN_TOAN');
  if (!dp || !wd || !em || !scc || !mon) {
    throw new Error('P043_PREFLIGHT_REQUIRED_SHEET_MISSING');
  }
  const records = {};
  [{sheet: em, type: 'EM'}, {sheet: scc, type: 'SCC'}].forEach(function(item) {
    const last = item.sheet.getLastRow();
    if (last < 2) return;
    const data = item.sheet.getRange(2, 2, last - 1, 14).getDisplayValues();
    data.forEach(function(r) {
      const visit = clean_(r[13]);
      if (!visit) return;
      const key = item.type + '/' + visit;
      if (records[key]) {
        result.differences.push('DUPLICATE_' + key);
      } else records[key] = {plate: normalizePlate_(r[0]), status: clean_(r[11])};
    });
  });
  const dpLast = dp.getLastRow();
  if (dpLast >= CFG.FIRST_DATA_ROW) {
    const live = dp.getRange(CFG.FIRST_DATA_ROW, 1,
      dpLast - CFG.FIRST_DATA_ROW + 1, CFG.DP.VISIT_ID).getDisplayValues();
    live.forEach(function(r, i) {
      const plate = normalizePlate_(r[CFG.DP.BIEN_SO - 1]);
      if (!plate) return;
      result.dispatchRows++;
      const id = clean_(r[CFG.DP.VISIT_ID - 1]);
      if (!id) {
        result.differences.push('DP_MISSING_VISIT row=' + (i + CFG.FIRST_DATA_ROW));
        return;
      }
      [{side:'EM',active:!!clean_(r[CFG.DP.NHOM_EM-1]),status:r[CFG.DP.TD_EM-1]},
       {side:'SCC',active:!!(clean_(r[CFG.DP.NGUOI_KT-1]) ||
                               clean_(r[CFG.DP.KTV1-1]) ||
                               clean_(r[CFG.DP.KTV2-1])),status:r[CFG.DP.TD_SCC-1]}]
      .forEach(function(t) {
        if (!t.active) return;
        const stored = records[t.side + '/' + id];
        if (!stored) result.differences.push('DATA_MISSING ' + t.side + ' ' + plate);
        else if (stored.plate !== plate) result.differences.push('PLATE_MISMATCH ' + t.side + ' ' + plate);
        else if (clean_(t.status) &&
            normalizeKey_(stored.status) !== normalizeKey_(t.status)) {
          result.differences.push('STATUS_MISMATCH ' + t.side + ' ' + plate +
            ' ' + clean_(t.status) + ' <> ' + stored.status);
        }
      });
    });
  }
  if (wd.getLastRow() >= CFG.WAITING_DELIVERY_FIRST_ROW) {
    result.waitingRows = wd.getRange(CFG.WAITING_DELIVERY_FIRST_ROW, 2,
      wd.getLastRow() - CFG.WAITING_DELIVERY_FIRST_ROW + 1, 1)
        .getDisplayValues().filter(function(r) { return !!clean_(r[0]); }).length;
  }
  result.monitor = mon.getRange(5, 1, 13, 3).getDisplayValues().map(function(r,i) {
    return {row:i+5,check:r[0],count:r[1],state:r[2]};
  });
  try {
    const props = PropertiesService.getScriptProperties().getProperties();
    result.pendingQueueCount = Object.keys(props).filter(function(k) {
      return k.indexOf(CONCURRENCY_CFG.DIRTY_PREFIX) === 0 ||
        k.indexOf(CONCURRENCY_CFG.DIRTY_FALLBACK_PREFIX) === 0;
    }).length;
  } catch (err) { result.warnings.push('QUEUE_INSPECTION_FAILED ' + err.message); }
  try {
    const triggers = ScriptApp.getProjectTriggers().map(function(t) {
      return t.getHandlerFunction();
    });
    result.workerTriggerCount = triggers.filter(function(t) {
      return t === 'runPendingVisitWorker';
    }).length;
    if (!result.workerTriggerCount) result.warnings.push('WORKER_TRIGGER_NOT_FOUND');
  } catch (err) { result.warnings.push('TRIGGER_INSPECTION_FAILED ' + err.message); }
  console.log('[P043_PRODUCTION_PREFLIGHT] ' + JSON.stringify(result));
  return result;
}


/***************************************************************
 * TPMH P0.4.6 - STANDARD UNDELIVERED ADDITIONAL-WORK FLOW
 * Safe universal workflow (not bound to any license plate).
 * A manual affirmative confirmation is required for physical
 * "vehicle has not been handed over" evidence. Never auto-merge
 * visitIds based on plate; never create completion timestamps.
 ***************************************************************/
const P046_EXTRA_WORK = Object.freeze({
  AUDIT: 'P046_AUDIT_PHAT_SINH_CONG_VIEC',
  FIRST: CFG.FIRST_DATA_ROW
});

function p046Audit_(ss, create) {
  let sh = ss.getSheetByName(P046_EXTRA_WORK.AUDIT);
  if (!sh && create) {
    sh = ss.insertSheet(P046_EXTRA_WORK.AUDIT);
    const h = ['RecordedAt','Phase','Plate','Old waiting visitId','New dispatch visitId',
      'Waiting row','Dispatch row','Old waiting A:H (raw)','Old Data EM A:O (raw)',
      'Old Data SCC A:P (raw)','New dispatch A:X (raw)','New Data EM row',
      'New Data SCC row','Reason / exception'];
    sh.getRange(1,1,1,h.length).setValues([h]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function p046Journal_(audit, phase, ctx, reason) {
  if (!audit) throw new Error('P046_AUDIT_MISSING');
  const row = [new Date(), phase, ctx.plate, ctx.oldId, ctx.newId,
    ctx.waitRow, ctx.dispatchRow, JSON.stringify(ctx.oldWaiting || []),
    JSON.stringify(ctx.oldEm || []), JSON.stringify(ctx.oldScc || []),
    JSON.stringify(ctx.dispatchRaw || []), ctx.newEmRow || '', ctx.newSccRow || '',
    String(reason || '')];
  const at = Math.max(2, audit.getLastRow() + 1);
  audit.getRange(at,1,1,row.length).setValues([row]);
  audit.getRange(at,1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  SpreadsheetApp.flush();
}
function p046AlreadyCommitted_(audit, newId) {
  if (!audit || audit.getLastRow() < 2) return false;
  return audit.getRange(2,2,audit.getLastRow()-1,4).getDisplayValues()
    .some(function(row){ return row[0] === 'COMMITTED' && row[3] === newId; });
}
/** Read-only preflight: no mutations of business sheets, timestamps or status. */
function p046Inspect_(ss, newId, expectedPlate) {
  const dp = ss.getSheetByName(CFG.SHEET_DISPATCH);
  const wd = ss.getSheetByName(CFG.SHEET_WAITING_DELIVERY);
  const em = ss.getSheetByName(CFG.SHEET_DATA_EM);
  const scc = ss.getSheetByName(CFG.SHEET_DATA_SCC);
  if (!dp || !wd || !em || !scc) throw new Error('P046_REQUIRED_SHEET_MISSING');
  newId = clean_(newId);
  if (!newId) throw new Error('P046_DISPATCH_VISIT_ID_REQUIRED');
  const owners = resolveDispatchRowsByVisitId_(dp, newId);
  if (owners.length !== 1) throw new Error('P046_NEW_VISIT_NOT_UNIQUE ' + owners.length);
  const dispatchRow = owners[0];
  const dispatchRaw = dp.getRange(dispatchRow,1,1,CFG.DP.LAST_SCC_STATUS).getValues()[0];
  const plate = normalizePlate_(dp.getRange(dispatchRow,CFG.DP.BIEN_SO).getDisplayValue());
  if (!plate || (expectedPlate && plate !== normalizePlate_(expectedPlate))) {
    throw new Error('P046_SELECTED_PLATE_CHANGED');
  }
  if (clean_(dispatchRaw[CFG.DP.VISIT_ID - 1]) !== newId) throw new Error('P046_NEW_ID_CHANGED');
  const dpLast = dp.getLastRow();
  const physicalDispatch = dp.getRange(CFG.FIRST_DATA_ROW,1,
    dpLast-CFG.FIRST_DATA_ROW+1,CFG.DP.VISIT_ID).getDisplayValues()
    .filter(function(r){return normalizePlate_(r[CFG.DP.BIEN_SO-1])===plate;});
  if (physicalDispatch.length !== 1) throw new Error('P046_MULTIPLE_DISPATCH_ROWS');
  const wdLast = waitingDeliveryDataLastRow_(wd);
  if (wdLast < CFG.WAITING_DELIVERY_FIRST_ROW) throw new Error('P046_NO_OLD_WAITING');
  const waiting = wd.getRange(CFG.WAITING_DELIVERY_FIRST_ROW,1,
    wdLast-CFG.WAITING_DELIVERY_FIRST_ROW+1,CFG.WD.VISIT_ID).getDisplayValues();
  const physicalWait = [];
  waiting.forEach(function(r,i){
    if (normalizePlate_(r[CFG.WD.BIEN_SO-1])===plate) {
      physicalWait.push({row:CFG.WAITING_DELIVERY_FIRST_ROW+i,display:r});
    }
  });
  if (physicalWait.length !== 1) throw new Error('P046_WAITING_OWNER_COUNT ' + physicalWait.length);
  const waitRow = physicalWait[0].row;
  const display = physicalWait[0].display;
  if (normalizeKey_(display[CFG.WD.TRANG_THAI-1]) !== 'cho giao') {
    throw new Error('P046_OLD_NOT_WAITING');
  }
  const oldId = clean_(display[CFG.WD.VISIT_ID-1]);
  if (!oldId || oldId === newId || waitingDeliveryRowsByVisitId_(wd,oldId).length !== 1) {
    throw new Error('P046_OLD_OWNER_NOT_UNIQUE');
  }
  if (resolveDispatchRowsByVisitId_(dp,oldId).length) {
    throw new Error('P046_OLD_VISIT_STILL_IN_DISPATCH');
  }
  const rec = readDispatchRow_(dp,dispatchRow); // visitId already verified, not created
  if (clean_(rec.visitId) !== newId || normalizePlate_(rec.plate) !== plate) {
    throw new Error('P046_DISPATCH_CHANGED');
  }
  if (normalizeKey_(display[CFG.WD.CVDV-1]) !== normalizeKey_(rec.advisor) ||
      normalizeKey_(display[CFG.WD.MODEL-1]) !== normalizeKey_(rec.model)) {
    throw new Error('P046_ADVISOR_OR_MODEL_DIFFERENT');
  }
  const oldRec = Object.assign({},rec,{visitId:oldId});
  const oldEmRow = findDataRow_(em,oldRec);
  const oldSccRow = findDataRow_(scc,oldRec);
  if (!oldEmRow && !oldSccRow) throw new Error('P046_NO_PREVIOUS_WORK_HISTORY');
  const oldEm = oldEmRow ? em.getRange(oldEmRow,1,1,15).getValues()[0] : [];
  const oldScc = oldSccRow ? scc.getRange(oldSccRow,1,1,16).getValues()[0] : [];
  [[em,oldEmRow,'EM'],[scc,oldSccRow,'SCC']].forEach(function(t) {
    if (!t[1]) return;
    if (normalizePlate_(t[0].getRange(t[1],2).getDisplayValue()) !== plate ||
        !isComplete_(t[0].getRange(t[1],13).getDisplayValue())) {
      throw new Error('P046_PREVIOUS_WORK_NOT_COMPLETE_' + t[2]);
    }
  });
  const route = classifyWork_(rec);
  const active = [
    {side:'EM',sheet:em,on:!!route.em.active,status:rec.emStatus,work:route.em.workText},
    {side:'SCC',sheet:scc,on:!!route.scc.active,status:rec.sccStatus,work:route.scc.workText}
  ].filter(function(t){return t.on;});
  // P0.4.8: The vehicle can still be physically here while the new job is
  // waiting for technician allocation. Never create fake Data / assignment.
  const pendingOnly = !active.length &&
    !!clean_(rec.workRaw) &&
    normalizeKey_(rec.sccStatus) === 'cho giao viec' &&
    !clean_(rec.emGroup) && !clean_(rec.inspector) &&
    !clean_(rec.ktv1) && !clean_(rec.ktv2) &&
    !clean_(rec.emStatus) && !valueExists_(rec.itemCount);
  if (!active.length && !pendingOnly) throw new Error('P046_NO_NEW_WORK_ASSIGNED');
  if (pendingOnly) {
    // No Data for an unassigned job; require absence on both sides.
    if (findDataRow_(em,rec) || findDataRow_(scc,rec))
      throw new Error('P048_PENDING_JOB_HAS_DATA_CONFLICT');
  }
  const targets=[];
  active.forEach(function(t){
    if (!clean_(t.status) || isComplete_(t.status) || !clean_(t.work)) {
      throw new Error('P046_NEW_WORK_MUST_BE_OPEN_' + t.side);
    }
    const dataRow = findDataRow_(t.sheet,rec);
    if (dataRow) {
      const stored = t.sheet.getRange(dataRow,1,1,15).getDisplayValues()[0];
      if (normalizePlate_(stored[1])!==plate || clean_(stored[14])!==newId ||
          normalizeKey_(stored[12])!==normalizeKey_(t.status) || clean_(stored[11])) {
        throw new Error('P046_NEW_DATA_CONFLICT_' + t.side);
      }
    }
    targets.push(Object.assign({},t,{row:dataRow}));
  });
  return {
    plate:plate, oldId:oldId, newId:newId, dispatchRow:dispatchRow,
    pendingOnly:pendingOnly,
    waitRow:waitRow, oldWaiting:wd.getRange(waitRow,1,1,8).getValues()[0],
    oldEm:oldEm, oldScc:oldScc, oldEmRow:oldEmRow,oldSccRow:oldSccRow,
    dispatchRaw:dispatchRaw, rec:rec, targets:targets, dp:dp,wd:wd,em:em,scc:scc,
    newEmRow:targets.some(function(t){return t.side==='EM';}) ?
      targets.filter(function(t){return t.side==='EM';})[0].row : 0,
    newSccRow:targets.some(function(t){return t.side==='SCC';}) ?
      targets.filter(function(t){return t.side==='SCC';})[0].row : 0
  };
}

/** Safe read-only function invoked from custom menu. */
function p046KiemTraCongViecPhatSinh() {
  const sel=p048SelectedRow_();
  try {
    const c=sel.visitId ? p046Inspect_(sel.ss,sel.visitId,sel.plate) :
      p048InspectBlankVisitId_(sel.ss,sel.row,sel.plate);
    const result={status:sel.visitId?'READY_FOR_CONFIRMATION':'READY_NEW_ID_AFTER_CONFIRMATION',
      plate:c.plate,oldVisitId:c.oldId,newVisitId:c.newId||'',
      newSides:(c.targets||[]).map(function(t){return t.side;}),
      pendingOnly:!!c.pendingOnly,oldWaitingRow:c.waitRow,
      dispatchRow:c.dispatchRow,
      note:c.pendingOnly?'NO Data until assigned, do not invent a technician':
        (!c.newId?'New unique visitId issued only during confirmed commit':'')};
    console.log('[P048_PREFLIGHT_READ_ONLY] '+JSON.stringify(result));
    return result;
  } catch(e) {
    const result={status:'BLOCKED',plate:sel.plate,
      visitId:sel.visitId||'',reason:String(e.message||e)};
    console.log('[P048_PREFLIGHT_READ_ONLY] '+JSON.stringify(result));
    return result;
  }
}
function p046Rollback_(ctx, snap, unsafeValue, created) {
  const errors=[];
  created.forEach(function(item){
    try {
      const rec=Object.assign({},ctx.rec,{visitId:ctx.newId});
      const row=findDataRow_(item.sheet,rec);
      if (row) {
        const d=item.sheet.getRange(row,1,1,15).getDisplayValues()[0];
        if (clean_(d[14])!==ctx.newId || normalizePlate_(d[1])!==ctx.plate ||
            isComplete_(d[12]) || clean_(d[11])) throw new Error('DATA_CHANGED');
        item.sheet.getRange(row,1,1,15).clearContent().clearNote();
      }
    } catch(e){errors.push(item.side+': '+e.message);}
  });
  try {restoreWaitingDeliveryCompactSnapshot_(snap.range,snap.data);} 
  catch(e){errors.push('WAITING: '+e.message);}
  try {
    const key=waitingCvdvUnsafeKey_(ctx.oldId);
    if (key && unsafeValue) PropertiesService.getScriptProperties().setProperty(key,unsafeValue);
  } catch(e){errors.push('WAITING_GUARD: '+e.message);}
  SpreadsheetApp.flush();
  return errors;
}
/** Mutating core must be called only after confirmation, under ScriptLock. */
function p046Commit_(ss,newId,plate,p049SourceProof) {
  const ctx=p046Inspect_(ss,newId,plate);
  // P0.4.9: a fully source-proven automatic operation may enter the same
  // locked, journalled transaction. Never treat inference as operator consent.
  if (p049SourceProof) p049AssertFreshSourceProof_(ss,ctx,p049SourceProof);
  if (isWebAppConfirmPriorityActive_()) throw new Error('P046_WEBAPP_PRIORITY_ACTIVE');
  const audit=p046Audit_(ss,true);
  if (p046AlreadyCommitted_(audit,newId)) throw new Error('P046_VISIT_ALREADY_LINKED');
  p046Journal_(audit,'PREPARED',ctx,p049SourceProof ?
    'P049_SOURCE_PROVEN_UNDELIVERED; reception+CVDV+history independently verified; not an operator confirmation' :
    'Operator explicitly confirmed: vehicle has NOT been delivered');
  // Snapshot including validation and notes; only use it under ScriptLock.
  const start=CFG.WAITING_DELIVERY_FIRST_ROW;
  const last=waitingDeliveryDataLastRow_(ctx.wd);
  const range=ctx.wd.getRange(start,1,last-start+1,CFG.WD.VISIT_ID);
  const snap={range:range,data:waitingDeliveryCompactSnapshot_(range)};
  const unsafeKey=waitingCvdvUnsafeKey_(ctx.oldId);
  const unsafeValue=unsafeKey ? PropertiesService.getScriptProperties().getProperty(unsafeKey) : null;
  const created=[];
  let changed=false;
  try {
    const again=p046Inspect_(ss,newId,plate);
    if (again.oldId!==ctx.oldId || again.waitRow!==ctx.waitRow ||
        again.dispatchRow!==ctx.dispatchRow) throw new Error('P046_OWNER_CHANGED');
    if (p049SourceProof) p049AssertFreshSourceProof_(ss,again,p049SourceProof);
    if (isWebAppConfirmPriorityActive_()) throw new Error('P046_WEBAPP_BECAME_ACTIVE');
    changed=true;
    if (!clearAndCompactWaitingDeliveryRow_(ctx.wd,ctx.waitRow,ctx.oldId)) {
      throw new Error('P046_WAITING_RELEASE_FAILED');
    }
    if (waitingDeliveryRowsByVisitId_(ctx.wd,ctx.oldId).length) throw new Error('P046_OLD_WAITING_STILL_EXISTS');
    if (resolveDispatchRowsByVisitId_(ctx.dp,ctx.newId).length!==1) throw new Error('P046_NEW_DISPATCH_OWNER_CHANGED');
    p04AssertPhysicalVisitOwnership_(ctx.dp,ctx.dispatchRow,ctx.plate);
    // Work statuses must remain open. Existing OPEN records are never overwritten.
    ctx.targets.forEach(function(t) {
      if (isWebAppConfirmPriorityActive_()) throw new Error('P046_WEBAPP_BECAME_ACTIVE');
      if (t.row) return;
      created.push({side:t.side,sheet:t.sheet}); // register BEFORE write: rollback even if write partially fails
      if (t.side==='EM') syncDataEM_(ctx.rec,t.work);
      else syncDataSCC_(ctx.rec,t.work);
    });
    SpreadsheetApp.flush();
    ctx.targets.forEach(function(t){
      const row=findDataRow_(t.sheet,ctx.rec);
      if (!row) throw new Error('P046_NEW_DATA_NOT_CREATED_'+t.side);
      const d=t.sheet.getRange(row,1,1,15).getDisplayValues()[0];
      if (normalizePlate_(d[1])!==ctx.plate || clean_(d[14])!==ctx.newId ||
          normalizeKey_(d[12])!==normalizeKey_(t.status) || clean_(d[11])) {
        throw new Error('P046_NEW_DATA_VERIFY_FAILED_'+t.side);
      }
      if (t.side==='EM') ctx.newEmRow=row;
      else ctx.newSccRow=row;
    });
    [[ctx.em,ctx.oldEmRow,ctx.oldEm,'EM'],[ctx.scc,ctx.oldSccRow,ctx.oldScc,'SCC']]
      .forEach(function(a){
        if (!a[1]) return;
        const unchanged=a[0].getRange(a[1],1,1,a[2].length).getValues()[0];
        if (JSON.stringify(unchanged)!==JSON.stringify(a[2])) {
          throw new Error('P046_OLD_HISTORY_CHANGED_'+a[3]);
        }
      });
    if (p049SourceProof) p049AssertFreshSourceProof_(ss,ctx,p049SourceProof);
    p046Journal_(audit,'COMMITTED',ctx,(p049SourceProof ?
      'P049_AUTO_SOURCE_PROVEN '+p049SourceProof.signature+' | ' : '')+
      (ctx.pendingOnly ?
      'Old waiting withdrawn (NOT delivered); new job WAITING FOR ASSIGNMENT; NO Data created' :
      'Old waiting withdrawn (NOT delivered); new OPEN work synchronized'));
    changed=false;
    const mon=ss.getSheetByName('P04_GIAM_SAT_AN_TOAN');
    let guard=[];
    try {guard=mon?mon.getRange('A9:C17').getDisplayValues():[];}
    catch(monErr){console.log('[P046_MONITOR_READ_WARNING] '+monErr.message);}
    const result={status:'COMMITTED',plate:ctx.plate,oldVisitId:ctx.oldId,
      newVisitId:ctx.newId,dataEMRow:ctx.newEmRow||0,dataSCCRow:ctx.newSccRow||0,
      watch:guard};
    console.log('[P046_EXTRA_WORK_RESULT] '+JSON.stringify(result));
    return result;
  } catch(e){
    if (changed) {
      const errors=p046Rollback_(ctx,snap,unsafeValue,created);
      try{p046Journal_(audit,errors.length?'ROLLBACK_INCOMPLETE':'ROLLED_BACK',ctx,
        e.message+(errors.length?' | '+errors.join(' | '):''));}
      catch(logErr){console.error('[P046_AUDIT_FAILED] '+logErr.message);}
      if(errors.length)throw new Error('P046_MANUAL_REVIEW_REQUIRED '+e.message+' / '+errors.join(' | '));
    }
    throw e;
  }
}
/** Universal UI operation: select the new work row in Điều phối. */
function p046XacNhanXeChuaGiaoPhatSinhViec() {
  const choice=p048SelectedRow_();
  const inspected=choice.visitId ?
    p046Inspect_(choice.ss,choice.visitId,choice.plate) :
    p048InspectBlankVisitId_(choice.ss,choice.row,choice.plate);
  const ui=SpreadsheetApp.getUi();
  const mode=inspected.pendingOnly ?
    'CHỜ PHÂN CÔNG: chưa có KTV. Sẽ KHÔNG tạo Data SCC trước khi phân công.\n' :
    (!choice.visitId ? 'THIẾU visitId: sẽ cấp ID duy nhất khi xác nhận, có kiểm tra và rollback.\n' : '');
  const prompt='Biển số: '+inspected.plate+'\n'+
    'Lượt Chờ giao cũ: '+inspected.oldId+'\n'+
    'Lượt công việc mới: '+(inspected.newId||'[sẽ cấp ID mới sau xác nhận]')+'\n'+
    'Trạng thái: '+mode+'\n'+
    'CHỈ chọn CÓ khi đã xác minh xe CHƯA RA CỔNG.\n'+
    'Thu hồi Chờ giao cũ (KHÔNG ghi Đã giao); giữ Data và giờ hoàn thành cũ.\n'+
    'Nếu không chắc chắn, chọn KHÔNG.';
  if (ui.alert('TPMH P0.4.8 – Xe chưa giao, phát sinh việc',prompt,ui.ButtonSet.YES_NO)!==ui.Button.YES) {
    console.log('[P048_CANCELLED] '+choice.plate);
    return {status:'CANCELLED'};
  }
  const lock=LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('P048_SYSTEM_BUSY_RETRY_LATER');
  try {
    if (isWebAppConfirmPriorityActive_()) throw new Error('P048_WEBAPP_PRIORITY_ACTIVE');
    const result=choice.visitId ?
      p046Commit_(choice.ss,choice.visitId,choice.plate) :
      p048CommitBlankVisitIdUnderLock_(choice.ss,choice.row,choice.plate,inspected);
    ui.alert('Đã liên kết công việc',result.plate+' đã liên kết đúng lượt.'+
      (result.pendingOnly?' Chờ phân công KTV; chưa tạo Data SCC.':'')+
      ' Xem '+P046_EXTRA_WORK.AUDIT,ui.ButtonSet.OK);
    return result;
  } finally {lock.releaseLock();}
}

/** Universal read-only Production smoke check; doesn't require a real conflict. */
function p046ProductionPreflightReadOnly() {
  const old=p043ProductionPreflightReadOnly();
  const mon=old.monitor || [];
  const required=['Điều phối thiếu visitId','Chờ giao thiếu visitId',
    'Trùng visitId Điều phối','Trùng visitId Chờ giao',
    'Một biển số ở cả hai danh sách','Trùng visitId giữa hai danh sách',
    'Phân công EM chưa có Data EM','Phân công SCC chưa có Data SCC',
    'Chờ giao còn việc chưa hoàn thành','Trạng thái EM Điều phối ≠ Data EM',
    'Trạng thái SCC Điều phối ≠ Data SCC'];
  const found=required.filter(function(k){return mon.some(function(item){return item.check===k;});});
  const issues=(old.differences||[]).concat(old.warnings||[]);
  const r={
    version:'TPMH P0.4.6 GENERAL EXTRA-WORK candidate',
    capturedAt:new Date().toISOString(),
    smoke:'READ_ONLY',
    checkedGuards:found.length,
    workerTriggerCount:old.workerTriggerCount,
    pendingQueueCount:old.pendingQueueCount,
    existingDifferences:issues,
    totalAnomalies:(mon.filter(function(m){return m.check==='TỔNG ĐIỂM BẤT THƯỜNG';})[0]||{}).count,
    generalWorkflowAvailable:typeof p046XacNhanXeChuaGiaoPhatSinhViec==='function',
    historyAuditSheetName:P046_EXTRA_WORK.AUDIT,
    note:'Chỉ kiểm tra; chức năng sửa yêu cầu chọn dòng Điều phối và xác nhận xe CHƯA GIAO'
  };
  console.log('[P046_PRODUCTION_PREFLIGHT_READ_ONLY] '+JSON.stringify(r));
  return r;
}



/***************************************************************
 * TPMH P0.4.8 – UNDELIVERED REOPEN HARDENING
 * Supports (a) new job awaiting assignment, (b) new dispatch row
 * missing visitId. Never merges old/new visits or invents completion.
 * Uses existing P046 audit, rollback and commit under ScriptLock.
 ***************************************************************/
function p048SelectedRow_() {
  const active=SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error('P048_OPEN_SPREADSHEET');
  const ss=laySpreadsheetWeb_();
  if (active.getId()!==ss.getId()) throw new Error('P048_WRONG_SPREADSHEET');
  const range=active.getActiveRange();
  if (!range || range.getNumRows()!==1 ||
      range.getSheet().getName()!==CFG.SHEET_DISPATCH ||
      range.getRow()<CFG.FIRST_DATA_ROW) throw new Error('P048_SELECT_ONE_DISPATCH_ROW');
  const row=range.getRow(), sh=range.getSheet();
  const plate=normalizePlate_(sh.getRange(row,CFG.DP.BIEN_SO).getDisplayValue());
  const visitId=clean_(sh.getRange(row,CFG.DP.VISIT_ID).getValue());
  if (!plate) throw new Error('P048_SELECTED_PLATE_EMPTY');
  return {ss:ss,row:row,plate:plate,visitId:visitId};
}

/** Read-only preflight for an existing dispatch visit without visitId.
 * Do not call readDispatchRow_ here: it can allocate visitId as a side effect.
 */
function p048InspectBlankVisitId_(ss, row, expectedPlate) {
  const dp=ss.getSheetByName(CFG.SHEET_DISPATCH);
  const wd=ss.getSheetByName(CFG.SHEET_WAITING_DELIVERY);
  const em=ss.getSheetByName(CFG.SHEET_DATA_EM);
  const scc=ss.getSheetByName(CFG.SHEET_DATA_SCC);
  if (!dp||!wd||!em||!scc) throw new Error('P048_REQUIRED_SHEET_MISSING');
  if (row<CFG.FIRST_DATA_ROW || row>dp.getLastRow())
    throw new Error('P048_DISPATCH_ROW_OUT_OF_BOUNDS');
  const raw=dp.getRange(row,1,1,CFG.DP.LAST_SCC_STATUS).getValues()[0];
  const display=dp.getRange(row,1,1,CFG.DP.LAST_SCC_STATUS).getDisplayValues()[0];
  const plate=normalizePlate_(display[CFG.DP.BIEN_SO-1]);
  if (!plate||plate!==normalizePlate_(expectedPlate))
    throw new Error('P048_SELECTED_PLATE_CHANGED');
  const idCell=dp.getRange(row,CFG.DP.VISIT_ID);
  if (clean_(raw[CFG.DP.VISIT_ID-1]) || idCell.getFormula())
    throw new Error('P048_DISPATCH_VISIT_ID_NOT_BLANK');
  const start=CFG.FIRST_DATA_ROW;
  const allDp=dp.getRange(start,1,dp.getLastRow()-start+1,CFG.DP.VISIT_ID).getDisplayValues();
  const owners=allDp.map(function(r,i){return {row:start+i,plate:normalizePlate_(r[CFG.DP.BIEN_SO-1])};})
    .filter(function(o){return o.plate===plate;});
  if (owners.length!==1 || owners[0].row!==row) throw new Error('P048_DISPATCH_PLATE_NOT_UNIQUE');
  const wdStart=CFG.WAITING_DELIVERY_FIRST_ROW;
  const wdEnd=waitingDeliveryDataLastRow_(wd);
  if (wdEnd<wdStart) throw new Error('P048_NO_OLD_WAITING');
  const wdRows=wd.getRange(wdStart,1,wdEnd-wdStart+1,CFG.WD.VISIT_ID).getDisplayValues();
  const oldMatches=wdRows.map(function(r,i){return {row:wdStart+i,data:r};})
    .filter(function(o){return normalizePlate_(o.data[CFG.WD.BIEN_SO-1])===plate;});
  if (oldMatches.length!==1) throw new Error('P048_WAITING_OWNER_COUNT '+oldMatches.length);
  const old=oldMatches[0];
  if (normalizeKey_(old.data[CFG.WD.TRANG_THAI-1])!=='cho giao')
    throw new Error('P048_OLD_NOT_WAITING');
  const oldId=clean_(old.data[CFG.WD.VISIT_ID-1]);
  if (!oldId || waitingDeliveryRowsByVisitId_(wd,oldId).length!==1 ||
      resolveDispatchRowsByVisitId_(dp,oldId).length)
    throw new Error('P048_OLD_OWNER_NOT_UNIQUE');
  if (normalizeKey_(old.data[CFG.WD.CVDV-1])!==normalizeKey_(display[CFG.DP.CVDV-1]) ||
      normalizeKey_(old.data[CFG.WD.MODEL-1])!==normalizeKey_(display[CFG.DP.MODEL-1]))
    throw new Error('P048_ADVISOR_OR_MODEL_DIFFERENT');
  const oldRec={visitId:oldId,plate:plate};
  const oldEmRow=findDataRow_(em,oldRec),oldSccRow=findDataRow_(scc,oldRec);
  if (!oldEmRow&&!oldSccRow) throw new Error('P048_OLD_WORK_NOT_FOUND');
  [[em,oldEmRow,'EM'],[scc,oldSccRow,'SCC']].forEach(function(t){
    if (!t[1]) return;
    const original=t[0].getRange(t[1],2,1,12).getDisplayValues()[0];
    if (normalizePlate_(original[0])!==plate || !isComplete_(original[11]))
      throw new Error('P048_OLD_WORK_NOT_COMPLETE_'+t[2]);
  });
  const hasEM=!!clean_(display[CFG.DP.NHOM_EM-1]);
  const hasSCC=!!(clean_(display[CFG.DP.NGUOI_KT-1]) ||
    clean_(display[CFG.DP.KTV1-1]) || clean_(display[CFG.DP.KTV2-1]));
  const emState=clean_(display[CFG.DP.TD_EM-1]);
  const sccState=clean_(display[CFG.DP.TD_SCC-1]);
  if (!clean_(display[CFG.DP.CV-1]) || !(hasEM||hasSCC) ||
      (hasEM && (!emState||isComplete_(emState))) ||
      (hasSCC && (!sccState||isComplete_(sccState))))
    throw new Error('P048_NEW_WORK_NOT_OPEN_OR_NOT_ASSIGNED');
  return {plate:plate,oldId:oldId,newId:'',dispatchRow:row,waitRow:old.row,
    pendingOnly:false,targets:[].concat(hasEM?[{side:'EM'}]:[],hasSCC?[{side:'SCC'}]:[]),dispatchRaw:raw,
    oldWaiting:wd.getRange(old.row,1,1,CFG.WD.VISIT_ID).getValues()[0],
    oldEmRow:oldEmRow,oldSccRow:oldSccRow,
    oldEm:oldEmRow?em.getRange(oldEmRow,1,1,15).getValues()[0]:[],
    oldScc:oldSccRow?scc.getRange(oldSccRow,1,1,16).getValues()[0]:[]};
}

/** Execute ONLY with ScriptLock and confirmed physical non-delivery.
 * No artificial timestamp; ID and waiting changes share one rollback path.
 */
function p048CommitBlankVisitIdUnderLock_(ss,row,plate,initial) {
  if (isWebAppConfirmPriorityActive_()) throw new Error('P048_WEBAPP_PRIORITY_ACTIVE');
  const ctx=p048InspectBlankVisitId_(ss,row,plate);
  if (ctx.oldId!==initial.oldId || ctx.waitRow!==initial.waitRow ||
      JSON.stringify(ctx.dispatchRaw)!==JSON.stringify(initial.dispatchRaw) ||
      JSON.stringify(ctx.oldWaiting)!==JSON.stringify(initial.oldWaiting))
    throw new Error('P048_SOURCE_CHANGED_AFTER_CONFIRMATION');
  const dp=ss.getSheetByName(CFG.SHEET_DISPATCH);
  const cell=dp.getRange(row,CFG.DP.VISIT_ID);
  const snap={value:cell.getValue(),note:cell.getNote(),
    validation:cell.getDataValidation(),format:cell.getNumberFormat()};
  if (clean_(snap.value)||cell.getFormula()) throw new Error('P048_VISIT_ID_CHANGED');
  const newId=generateUniqueVisitId_(ss);
  const audit=p046Audit_(ss,true);
  let allocated=false;
  try {
    allocated=true; // set before write: a partially successful API write must be rolled back
    cell.setNumberFormat('@').setValue(newId);
    SpreadsheetApp.flush();
    if (clean_(cell.getDisplayValue())!==newId ||
        resolveDispatchRowsByVisitId_(dp,newId).length!==1)
      throw new Error('P048_VISIT_ID_ALLOCATION_VERIFY_FAILED');
    const logCtx=Object.assign({},ctx,{newId:newId});
    p046Journal_(audit,'ID_ALLOCATED',logCtx,
      'Unique ID issued under ScriptLock after explicit undelivered confirmation');
    const result=p046Commit_(ss,newId,plate); // existing verified, journaled transaction
    if (!result || result.status!=='COMMITTED')
      throw new Error('P048_UNEXPECTED_COMMIT_RESULT');
    result.issuedVisitId=true;
    console.log('[P048_RECOVERY_RESULT] '+JSON.stringify(result));
    return result;
  } catch(e) {
    if (!allocated) throw e;
    // If commit wrote anything lasting, do NOT orphan it by deleting its ID.
    const waitingSheet=ss.getSheetByName(CFG.SHEET_WAITING_DELIVERY);
    const ownerRows=waitingDeliveryRowsByVisitId_(waitingSheet,ctx.oldId);
    const oldStillWaiting=ownerRows.length===1 && ownerRows[0]===ctx.waitRow &&
      JSON.stringify(waitingSheet.getRange(ctx.waitRow,1,1,CFG.WD.VISIT_ID).getValues()[0])===
      JSON.stringify(ctx.oldWaiting);
    let createdData=false;
    try {
      const test={visitId:newId,plate:plate};
      createdData=!!(findDataRow_(ss.getSheetByName(CFG.SHEET_DATA_EM),test)||
        findDataRow_(ss.getSheetByName(CFG.SHEET_DATA_SCC),test));
    } catch(dataErr) {createdData=true;}
    if (!oldStillWaiting || createdData || p046AlreadyCommitted_(audit,newId)) {
      console.error('[P048_MANUAL_REVIEW_REQUIRED] '+plate+'/'+newId+': '+String(e.message||e));
      throw new Error('P048_MANUAL_REVIEW_REQUIRED '+newId+' '+String(e.message||e));
    }
    try {
      cell.clearDataValidations();
      cell.setValue(snap.value);
      cell.setNote(snap.note);
      cell.setNumberFormat(snap.format);
      cell.setDataValidation(snap.validation);
      SpreadsheetApp.flush();
      if (clean_(cell.getValue())) throw new Error('VISIT_ID_NOT_RESTORED');
      p046Journal_(audit,'ID_ROLLED_BACK',Object.assign({},ctx,{newId:newId}),
        'Waiting preserved; no new Data; allocation rolled back: '+String(e.message||e));
    } catch(restoreError) {
      throw new Error('P048_ID_ROLLBACK_INCOMPLETE '+newId+': '+String(restoreError.message||restoreError));
    }
    throw e;
  }
}




/***************************************************************
 * TPMH P0.4.7 - SOURCE-BOUND OPEN DATA RECONCILIATION
 * This is deliberately restricted to already-open work.
 * Main Worker holds ScriptLock before calling this function.
 * No new time / completion status; never deletes old data.
 ***************************************************************/
function p047RestoreMissingOpenDataUnderLock_(ss, workerDeadline) {
  const out = {checked:0, repaired:0, skipped:0, errors:[], busy:false};
  if (isWebAppConfirmPriorityActive_()) { out.busy=true; return out; }
  // P0.5.7.4: allow one extra bounded scan/commit window for OPEN work.
  // This is cooperative, NOT a hard wall-clock deadline for Sheets calls.
  // The worker's shared ScriptLock and WebApp priority remain in force.
  const deadline = Math.min(Number(workerDeadline || Number.MAX_SAFE_INTEGER), Date.now()+8000);
  if (Date.now()>=deadline) return out;
  const dp=ss.getSheetByName(CFG.SHEET_DISPATCH);
  const em=ss.getSheetByName(CFG.SHEET_DATA_EM);
  const scc=ss.getSheetByName(CFG.SHEET_DATA_SCC);
  if (!dp || !em || !scc) {out.errors.push('P047_REQUIRED_SHEET_MISSING');return out;}
  const first=CFG.FIRST_DATA_ROW;
  const last=dp.getLastRow();
  if (last<first) return out;
  try {
    // 3 bulk reads instead of per-row Data scans, even with an empty queue.
    const dispatch=dp.getRange(first,1,last-first+1,CFG.DP.VISIT_ID).getDisplayValues();
    function idSet_(sh) {
      const n=sh.getLastRow()-1;
      return new Set(n>0?sh.getRange(2,15,n,1).getDisplayValues()
        .map(function(r){return clean_(r[0]);}).filter(Boolean):[]);
    }
    const knownEM=idSet_(em),knownSCC=idSet_(scc);
    // Fail closed if a visitId appears more than once in active Dispatch.
    // A same-visit duplicate must NEVER cause a blind Data append.
    const countById={};
    dispatch.forEach(function(v){
      const key=clean_(v[CFG.DP.VISIT_ID-1]);
      if (key) countById[key]=(countById[key]||0)+1;
    });
    // Best-effort cache: retries are always idempotent, and a cache eviction
    // only causes an earlier SAFE retry, not a duplicate write.
    const retryCache=CacheService.getScriptCache();
    for (let i=0;i<dispatch.length && out.repaired<2 && Date.now()<deadline;i++) {
      if (isWebAppConfirmPriorityActive_()) {out.busy=true;break;}
      const v=dispatch[i];
      const plate=normalizePlate_(v[CFG.DP.BIEN_SO-1]);
      const visitId=clean_(v[CFG.DP.VISIT_ID-1]);
      if (!plate || !visitId) continue;
      out.checked++;
      // Only active assignments with explicitly OPEN status qualify.
      // A blank state or Hoàn thành always needs normal source validation.
      const inspector=clean_(v[CFG.DP.NGUOI_KT-1]);
      const ktv1=clean_(v[CFG.DP.KTV1-1]);
      const ktv2=clean_(v[CFG.DP.KTV2-1]);
      // P0.5.3.1 invariant: Mr. Khởi without separate SCC KTV is EM only.
      const mrKhoiOnly=/^mr\.?\s*khoi$/.test(normalizeKey_(inspector)) &&
        !ktv1 && !ktv2;
      const emAssigned=!!clean_(v[CFG.DP.NHOM_EM-1]) || mrKhoiOnly;
      const sccAssigned=!!(inspector||ktv1||ktv2) && !mrKhoiOnly;
      const emStatus=clean_(v[CFG.DP.TD_EM-1]);
      const sccStatus=clean_(v[CFG.DP.TD_SCC-1]);
      const needEM=emAssigned && emStatus && !isComplete_(emStatus) && !knownEM.has(visitId);
      const needSCC=sccAssigned && sccStatus && !isComplete_(sccStatus) && !knownSCC.has(visitId);
      if (!needEM && !needSCC) continue;
      const row=first+i;
      if (countById[visitId]!==1) {
        out.errors.push('P0532_DUPLICATE_DISPATCH_VISIT_ID row='+row+' / '+visitId);
        continue;
      }
      const retryKey='TPMH_P0532_OPEN_'+visitId;
      let retry={attempts:0,nextAt:0};
      try {retry=JSON.parse(retryCache.get(retryKey)||'null')||retry;} catch(ignored){}
      if (Number(retry.nextAt||0)>Date.now()) {out.skipped++;continue;}
      try {
        // Re-read before any write. Do not trust stale sheet positions.
        const rec=readDispatchRow_(dp,row);
        if (clean_(rec.visitId)!==visitId || normalizePlate_(rec.plate)!==plate)
          throw new Error('P047_IDENTITY_CHANGED');
        p04AssertPhysicalVisitOwnership_(dp,row,plate);
        const route=classifyWork_(rec);
        if (Date.now()>=deadline || isWebAppConfirmPriorityActive_()) {out.busy=true;break;}
        if (needEM && route.em.active && clean_(rec.emStatus) && !isComplete_(rec.emStatus)
          && !findDataRow_(em,rec)) {
          syncDataForVisit_(rec,route.em.workText,true);
          if (!findDataRow_(em,rec)) throw new Error('P047_EM_VERIFY_FAILED');
          knownEM.add(visitId);
          out.repaired++;
        }
        if (out.repaired>=2 || Date.now()>=deadline || isWebAppConfirmPriorityActive_()) break;
        if (needSCC && route.scc.active && clean_(rec.sccStatus) && !isComplete_(rec.sccStatus)
          && !findDataRow_(scc,rec)) {
          syncDataForVisit_(rec,route.scc.workText,false);
          if (!findDataRow_(scc,rec)) throw new Error('P047_SCC_VERIFY_FAILED');
          knownSCC.add(visitId);
          out.repaired++;
        }
        // Clean retry state only after a successful processing pass.
        try {retryCache.remove(retryKey);} catch(ignored){}
      } catch(e) {
        const attempts=Math.min(8,Number(retry.attempts||0)+1);
        const delayMs=Math.min(5*60000,Math.pow(2,Math.min(5,attempts-1))*30000);
        try {retryCache.put(retryKey,JSON.stringify({
          attempts:attempts,nextAt:Date.now()+delayMs
        }),600);} catch(ignored){}
        out.errors.push(plate+'/'+visitId+': '+String(e.message||e));
        console.warn('[P0532_OPEN_DATA_RETRY] '+JSON.stringify({
          plate:plate,visitId:visitId,attempt:attempts,
          retryAfterSeconds:Math.round(delayMs/1000),
          error:String(e.message||e).slice(0,320)
        }));
      }
    }
  } catch(e) {out.errors.push('P047_SCAN_FAILED: '+String(e.message||e));}
  if (out.repaired || out.errors.length) {
    console.log('[P047_OPEN_RECONCILE] '+JSON.stringify(out));
  }
  return out;
}





/*****************************************************************
 * TPMH P0.4.9 - SOURCE-BOUND AUTO-RECOVERY (Production candidate)
 * Upgrade on exact P0.4.8 Main. No additional trigger / new clone.
 *
 * Switch: OFF until explicitly enabled after read-only acceptance.
 * Auto-recovery only for SAME-DAY, single reception IN and no OUT,
 * matching single CVDV "Trong xưởng", completed old EM/SCC,
 * and a new open dispatch with a UNIQUE assigned visitId.
 *
 * Fails CLOSED for any missing/different/late source, multiple owner,
 * new missing visitId, completed new work, or conflicting history.
 * Never invents gate exit or job-completion timestamps.
 * Existing P0.4.7 handles missing OPEN Data after source-safe recovery.
 * Existing Monitor P1 closes stale LSC alarms during its own scan;
 * this module does not write to the separate Monitor project.
 *****************************************************************/
const P049_AUTO_CFG = Object.freeze({
  VERSION:'TPMH P0.4.9 Source-Bound Auto-Recovery CANDIDATE',
  ENABLE_KEY:'TPMH_P049_AUTO_RECOVERY_ENABLED',
  RECEPTION_ID:'10dxPgBNvoYZ4oDydtf1-dNsGLxfHxJiRJ5l7uoJIqso',
  RECEPTION_TAB:'Giờ ra vào 2026',
  RECEPTION_TAIL:1800,
  MIN_BUDGET_MS:6500,
  MAX_ONE_RECOVERY_PER_TICK:1,
  SKIP_CACHE_SECONDS:120,
  // P0.5: Chỉ cho phép liên kết công việc phát sinh trong cùng lượt lưu xưởng
  // tối đa 60 ngày. Nhiều ngày không đồng nghĩa nhiều lần vào cổng.
  MAX_CONTINUOUS_DAYS:60
});
function p049IsEnabled_(){
  return PropertiesService.getScriptProperties().getProperty(P049_AUTO_CFG.ENABLE_KEY)==='ON';
}
function p049HHmm_(value,ss){
  if(value instanceof Date&&!isNaN(value.getTime())){
    const tz=ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone();
    return Utilities.formatDate(value,tz,'HH:mm');
  }
  const m=String(value||'').match(/(?:^|\s)(\d{1,2}):(\d{2})(?::\d{2})?(?:$|\s)/);
  return m?('0'+Number(m[1])).slice(-2)+':'+m[2]:'';
}
function p049ClockMinutes_(h){
  const m=String(h||'').match(/^(\d\d):(\d\d)$/);
  if(!m)return -1;
  const a=Number(m[1]),b=Number(m[2]);
  return a<24&&b<60?a*60+b:-1;
}
function p049ReceptionStamp_(v){
  const t=String(v||'').trim();
  const m=t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if(!m)return null;
  const dd=Number(m[1]),mm=Number(m[2]),year=Number(m[3]);
  const hh=Number(m[4]),mi=Number(m[5]);
  if(dd<1||dd>31||mm<1||mm>12||hh>23||mi>59)return null;
  const yyyy=year<100?2000+year:year;
  const date=new Date(Date.UTC(yyyy,mm-1,dd));
  if(date.getUTCFullYear()!==yyyy||date.getUTCMonth()!==mm-1||date.getUTCDate()!==dd)return null;
  const day=yyyy+'-'+('0'+mm).slice(-2)+'-'+('0'+dd).slice(-2);
  return {day:day,clock:hh*60+mi,display:t};
}
function p049AdvisorKey_(v){
  // P0.5.4: only comparison against external CVDV/Reception source.
  // All other owner, visitId and unique-event proof guards stay untouched.
  return p054AdvisorMatchKey_(v);
}
/**
 * P0.5: So sánh mốc ngày/giờ theo phút lịch (không tạo timestamp nghiệp vụ).
 * Chỉ nhận ISO day + đồng hồ được đọc từ nguồn. Date.UTC dùng cho SO SÁNH,
 * tuyệt đối không ghi giá trị này về Sheet như một giờ hoàn thành.
 */
function p050CalendarMinute_(day, clockMinutes) {
  const match=/^(\d{4})-(\d{2})-(\d{2})$/.exec(String(day||''));
  if(!match || !Number.isInteger(clockMinutes) || clockMinutes<0 || clockMinutes>=1440)
    throw new Error('P050_DAY_CLOCK_INVALID');
  const year=Number(match[1]),month=Number(match[2]),date=Number(match[3]);
  const utc=Date.UTC(year,month-1,date);
  const verify=new Date(utc);
  if(verify.getUTCFullYear()!==year||verify.getUTCMonth()!==month-1||
     verify.getUTCDate()!==date)throw new Error('P050_CALENDAR_DAY_INVALID');
  return Math.round(utc/60000)+clockMinutes;
}
/**
 * Data!L có thể được format chỉ HH:mm nhưng lưu datetime thực tế.
 * Với công việc QUA NGÀY, tuyệt đối không suy ngày từ một ô chỉ chứa giờ
 * (Date 1899 hoặc text 12:07): thiếu ngày -> chặn tự phục hồi.
 */
function p050DatedCompletion_(raw, ss) {
  if(!(raw instanceof Date)||isNaN(raw.getTime()))return null;
  const tz=ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone();
  const day=Utilities.formatDate(raw,tz,'yyyy-MM-dd');
  if(day<'2020-01-01')return null;
  const clock=p049ClockMinutes_(Utilities.formatDate(raw,tz,'HH:mm'));
  if(clock<0)return null;
  return {day:day,clock:clock,minute:p050CalendarMinute_(day,clock)};
}
/**
 * P0.5 READ-ONLY proof: một xe có thể lưu xưởng qua ngày, nhưng chỉ một
 * lượt vào thực tế chưa có giờ ra, một hồ sơ CVDV còn Trong xưởng, lịch
 * sử công việc hoàn thành và công việc phát sinh vẫn OPEN mới được AUTO.
 * Không tự liên kết xe đã ghi Hoàn thành trước khi chứng cứ event được
 * xác minh; những trường hợp đó tiếp tục dùng đối soát có kiểm soát.
 */
function p049UndeliveredSourceProof_(ss,ctx){
  const oldDay=cvdvDateKey_(ctx.oldWaiting[0],ss);
  const newDay=cvdvDateKey_(ctx.rec.date,ss);
  if(!oldDay||!newDay)throw new Error('P050_WORK_DATES_REQUIRED');
  const oldStart=p050CalendarMinute_(oldDay,0);
  const newStart=p050CalendarMinute_(newDay,0);
  const gapDays=(newStart-oldStart)/1440;
  if(!Number.isInteger(gapDays)||gapDays<0||gapDays>P049_AUTO_CFG.MAX_CONTINUOUS_DAYS)
    throw new Error('P050_CONTINUOUS_STAY_DATE_OUT_OF_SCOPE');
  if(!ctx.oldId||!ctx.newId||ctx.oldId===ctx.newId)
    throw new Error('P049_TWO_DISTINCT_VISITS_REQUIRED');
  const receiving=p049ClockMinutes_(p049HHmm_(ctx.rec.receiveTime,ss));
  if(receiving<0)throw new Error('P049_NEW_WORK_RECEIVE_TIME_REQUIRED');
  const newReceiveMinute=p050CalendarMinute_(newDay,receiving);

  // p046Inspect_ đã kiểm tra tất cả old Data đều Hoàn thành, biển số và ID.
  const prev=[];
  [[ctx.em,ctx.oldEmRow,'EM'],[ctx.scc,ctx.oldSccRow,'SCC']].forEach(function(item){
    if(!item[1])return;
    const raw=item[0].getRange(item[1],1,1,15).getValues()[0];
    const disp=item[0].getRange(item[1],1,1,15).getDisplayValues()[0];
    if(cvdvDateKey_(raw[0],ss)!==oldDay)
      throw new Error('P050_OLD_WORK_DATE_MISMATCH_'+item[2]);
    const clock=p049ClockMinutes_(p049HHmm_(disp[8],ss));
    const shownDone=p049ClockMinutes_(p049HHmm_(disp[11],ss));
    const finished=p050DatedCompletion_(raw[11],ss);
    if(clock<0||shownDone<0)
      throw new Error('P050_OLD_WORK_TIME_MISSING_'+item[2]);
    if(gapDays>0&&!finished)
      throw new Error('P050_OLD_WORK_COMPLETION_DATE_REQUIRED_'+item[2]);
    const began=p050CalendarMinute_(oldDay,clock);
    const done=finished?finished.minute:p050CalendarMinute_(oldDay,shownDone);
    if(finished && finished.clock!==shownDone)
      throw new Error('P050_COMPLETION_DISPLAY_RAW_MISMATCH_'+item[2]);
    if(done<began||done>=newReceiveMinute)
      throw new Error('P050_OLD_WORK_TIMELINE_NOT_VERIFIED_'+item[2]);
    prev.push({side:item[2],row:item[1],started:began,finished:done});
  });
  if(!prev.length)throw new Error('P049_OLD_COMPLETED_WORK_REQUIRED');

  const cv=ss.getSheetByName(CFG.SHEET_CVDV_DATA);
  if(!cv)throw new Error('P049_CVDV_SOURCE_NOT_FOUND');
  const first=CFG.CVDV_DATA_FIRST_ROW,last=cv.getLastRow();
  if(last<first)throw new Error('P049_CVDV_EMPTY');
  const width=Math.max(CFG.CVDV_DATA.NGAY,CFG.CVDV_DATA.BIEN_SO,
                       CFG.CVDV_DATA.TINH_TRANG,CFG.CVDV_DATA.CVDV);
  const range=cv.getRange(first,1,last-first+1,width);
  const raws=range.getValues(),shown=range.getDisplayValues();
  const matched=[];
  for(let i=0;i<shown.length;i++){
    const row=shown[i];
    if(normalizePlate_(row[CFG.CVDV_DATA.BIEN_SO-1])!==ctx.plate)continue;
    const date=cvdvDateKey_(raws[i][CFG.CVDV_DATA.NGAY-1],ss);
    if(!date)throw new Error('P049_CVDV_DATE_MISSING');
    const status=normalizeKey_(row[CFG.CVDV_DATA.TINH_TRANG-1]);
    if(date>=oldDay)matched.push({row:first+i,date:date,status:status,
      advisor:p049AdvisorKey_(row[CFG.CVDV_DATA.CVDV-1])});
  }
  if(matched.length!==1||matched[0].date!==oldDay||
     matched[0].status!=='trong xuong'||
     matched[0].advisor!==p049AdvisorKey_(ctx.rec.advisor))
    throw new Error('P050_CVDV_CONTINUOUS_OPEN_VISIT_REQUIRED');

  // Không suy luận từ trạng thái Điều phối: Lễ tân chứng minh ra/vào,
  // CVDV là nguồn đối chiếu trạng thái cuối. Thiếu một bên -> BLOCK.
  let lt;
  try{
    lt=SpreadsheetApp.openById(P049_AUTO_CFG.RECEPTION_ID)
      .getSheetByName(P049_AUTO_CFG.RECEPTION_TAB);
  }catch(e){throw new Error('P049_RECEPTION_ACCESS_FAILED '+String(e.message||e));}
  if(!lt)throw new Error('P049_RECEPTION_TAB_MISSING');
  const end=lt.getLastRow(),begin=Math.max(2,end-P049_AUTO_CFG.RECEPTION_TAIL+1);
  if(end<begin)throw new Error('P049_RECEPTION_EMPTY');
  const rows=lt.getRange(begin,2,end-begin+1,6).getDisplayValues();
  const gate=[];
  rows.forEach(function(r,i){
    if(normalizePlate_(r[0])!==ctx.plate)return;
    const entry=p049ReceptionStamp_(r[3]);
    if(!entry)throw new Error('P049_RECEPTION_INGRESS_INVALID row='+String(begin+i));
    if(entry.day<oldDay)return;
    // P056: ignore no inference – this is never a service gate visit.
    if (p056IsNoServiceType_(r[1]))
      throw new Error('P056_NO_SERVICE_GATE_BLOCK row='+String(begin+i));
    gate.push({row:begin+i,day:entry.day,entryClock:entry.clock,
      out:String(r[4]||'').trim(),advisor:p049AdvisorKey_(r[5])});
  });
  const g=gate[0];
  if(gate.length!==1||!g||g.day!==oldDay||g.out||
     g.advisor!==p049AdvisorKey_(ctx.rec.advisor))
    throw new Error('P050_SINGLE_UNEXITED_RECEPTION_REQUIRED');
  const gateEntry=p050CalendarMinute_(g.day,g.entryClock);
  if(gateEntry>=newReceiveMinute||prev.some(function(x){return gateEntry>x.started;}))
    throw new Error('P050_GATE_ENTRY_AFTER_WORK');
  const completedSignature=prev.map(function(x){return x.side+':'+x.finished;}).join(',');
  return {plate:ctx.plate,oldId:ctx.oldId,newId:ctx.newId,
    oldDay:oldDay,newDay:newDay,cvdvRow:matched[0].row,receptionRow:g.row,
    gateEntry:gateEntry,oldCompleted:completedSignature,
    signature:[ctx.plate,ctx.oldId,ctx.newId,oldDay,newDay,newReceiveMinute,
      matched[0].row,g.row,gateEntry,completedSignature].join('|')};
}
function p049AssertFreshSourceProof_(ss,ctx,proof){
  if(!proof||!proof.signature)throw new Error('P049_AUTO_PROOF_REQUIRED');
  const fresh=p049UndeliveredSourceProof_(ss,ctx);
  if(fresh.signature!==proof.signature)
    throw new Error('P049_SOURCE_CHANGED_RETRY_LATER');
  return fresh;
}
/** Read-only discovery; no new visitId generated and no mutation. */
function p049ListConflicts_(ss){
  const dp=ss.getSheetByName(CFG.SHEET_DISPATCH);
  const wd=ss.getSheetByName(CFG.SHEET_WAITING_DELIVERY);
  if(!dp||!wd)throw new Error('P049_REQUIRED_SHEET_MISSING');
  const first=CFG.FIRST_DATA_ROW,last=dp.getLastRow();
  const wf=CFG.WAITING_DELIVERY_FIRST_ROW,wl=waitingDeliveryDataLastRow_(wd);
  if(last<first||wl<wf)return [];
  const waiting={},wrows=wd.getRange(wf,1,wl-wf+1,8).getDisplayValues();
  wrows.forEach(function(v,i){
    const p=normalizePlate_(v[1]);
    if(p&&normalizeKey_(v[6])==='cho giao'){
      if(!waiting[p])waiting[p]=[];
      waiting[p].push({row:wf+i,id:clean_(v[7])});
    }
  });
  const seen={},out=[];
  dp.getRange(first,1,last-first+1,CFG.DP.VISIT_ID).getDisplayValues()
  .forEach(function(v,i){
    const plate=normalizePlate_(v[CFG.DP.BIEN_SO-1]);
    if(!plate||!waiting[plate])return;
    seen[plate]=(seen[plate]||0)+1;
    out.push({plate:plate,dispatchRow:first+i,
      newId:clean_(v[CFG.DP.VISIT_ID-1]),waiting:waiting[plate]});
  });
  out.forEach(function(x){x.dispatchOwners=seen[x.plate];});
  return out;
}
function p049AutoTickUnderLock_(ss,deadline){
  const result={version:P049_AUTO_CFG.VERSION,status:'DISABLED',recovered:0};
  if(!p049IsEnabled_())return result;
  result.status='NOT_NEEDED';
  if(isWebAppConfirmPriorityActive_()){result.status='WEBAPP_PRIORITY';return result;}
  if(Date.now()+P049_AUTO_CFG.MIN_BUDGET_MS >= Number(deadline||0)){
    result.status='TIME_BUDGET_SKIP';return result;
  }
  const list=p049ListConflicts_(ss);
  if(!list.length)return result;
  // Bounded: at most one candidate per 1-min worker tick.
  const cache=CacheService.getScriptCache();
  for(let i=0;i<list.length;i++){
    const x=list[i];
    if(x.dispatchOwners!==1||x.waiting.length!==1||!x.newId)continue;
    const cooldown='P049_BLOCKED_'+x.newId;
    if(cache.get(cooldown))continue;
    result.plate=x.plate;
    try{
      if(Date.now()+P049_AUTO_CFG.MIN_BUDGET_MS>=Number(deadline||0)){
        result.status='TIME_BUDGET_SKIP';break;
      }
      const ctx=p046Inspect_(ss,x.newId,x.plate);
      const proof=p049UndeliveredSourceProof_(ss,ctx);
      const committed=p046Commit_(ss,x.newId,x.plate,proof);
      if(!committed||committed.status!=='COMMITTED')
        throw new Error('P049_COMMIT_NOT_CONFIRMED');
      result.status='COMMITTED';result.recovered=1;
      result.newId=x.newId;result.proof=proof.signature;
      console.log('[P049_AUTO_RECOVERY_COMMITTED] '+JSON.stringify(result));
      return result;
    }catch(e){
      result.status='BLOCKED_SOURCE_OR_COMMIT';
      result.reason=String(e.message||e);
      console.warn('[P049_AUTO_RECOVERY_BLOCKED] '+JSON.stringify(result));
      if(!/MANUAL_REVIEW_REQUIRED|ROLLBACK_INCOMPLETE/.test(result.reason))
        cache.put(cooldown,'1',P049_AUTO_CFG.SKIP_CACHE_SECONDS);
      // Never try a second potentially hazardous conflict in this tick.
      return result;
    }
  }
  return result;
}
function p049ProductionPreflightReadOnly(){
  const ss=laySpreadsheetWeb_();
  const target=ss.getSheetByName(CFG.SHEET_DISPATCH);
  const waiting=ss.getSheetByName(CFG.SHEET_WAITING_DELIVERY);
  const em=ss.getSheetByName(CFG.SHEET_DATA_EM);
  const sc=ss.getSheetByName(CFG.SHEET_DATA_SCC);
  const cv=ss.getSheetByName(CFG.SHEET_CVDV_DATA);
  if(!target||!waiting||!em||!sc||!cv)throw new Error('P049_REQUIRED_SHEET_MISSING');
  const triggers=ScriptApp.getProjectTriggers()
    .filter(function(t){return t.getHandlerFunction()==='runPendingVisitWorker';}).length;
  const output={version:P049_AUTO_CFG.VERSION,readOnly:true,
    enabled:p049IsEnabled_(),workerTriggerCount:triggers,
    receptionReachable:false,conflicts:[],monitor:[],safeToEnable:false};
  try {
    const rt=SpreadsheetApp.openById(P049_AUTO_CFG.RECEPTION_ID)
      .getSheetByName(P049_AUTO_CFG.RECEPTION_TAB);
    if(!rt)throw new Error('P049_RECEPTION_TAB_MISSING');
    // Probe a small, existing header range to verify actual read access.
    rt.getRange('B1:G1').getDisplayValues();
    output.receptionReachable=true;
  } catch(e) {output.receptionError=String(e.message||e);}
  output.safeToEnable=triggers===1&&output.receptionReachable;
  const candidates=p049ListConflicts_(ss).slice(0,5);
  candidates.forEach(function(x){
    const entry={plate:x.plate,dispatchRow:x.dispatchRow,
      newId:x.newId,waitingRows:x.waiting.map(function(w){return w.row;})};
    try{
      if(x.dispatchOwners!==1||x.waiting.length!==1||!x.newId)
        throw new Error('P049_OWNER_OR_NEW_ID_AMBIGUOUS');
      const ctx=p046Inspect_(ss,x.newId,x.plate);
      entry.proof=p049UndeliveredSourceProof_(ss,ctx).signature;
      entry.status='SOURCE_PROVEN_CANDIDATE';
    }catch(e){entry.status='NEEDS_OPERATOR_REVIEW';entry.reason=String(e.message||e);}
    output.conflicts.push(entry);
  });
  const mon=ss.getSheetByName('P04_GIAM_SAT_AN_TOAN');
  if(mon)output.monitor=mon.getRange('A9:C17').getDisplayValues();
  console.log('[P049_PREFLIGHT_READ_ONLY] '+JSON.stringify(output));
  return output;
}
function p049EnableAutoRecovery(){
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(10000))throw new Error('P049_BUSY');
  try{
    const check=p049ProductionPreflightReadOnly();
    if(!check.safeToEnable)throw new Error('P049_WORKER_TRIGGER_MUST_EQUAL_ONE');
    if(isWebAppConfirmPriorityActive_())throw new Error('P049_WEBAPP_PRIORITY_ACTIVE');
    PropertiesService.getScriptProperties().setProperty(P049_AUTO_CFG.ENABLE_KEY,'ON');
    console.log('[P049_AUTORECOVERY_ENABLED] '+JSON.stringify({version:check.version,workers:check.workerTriggerCount}));
    return {enabled:true,version:check.version,existingWorkerTriggers:check.workerTriggerCount};
  }finally{lock.releaseLock();}
}
function p049DisableAutoRecovery(){
  PropertiesService.getScriptProperties().setProperty(P049_AUTO_CFG.ENABLE_KEY,'OFF');
  console.log('[P049_AUTORECOVERY_DISABLED]');
  return {enabled:false};
}




/**************************************************************************
 * TPMH P0.5.3 — LSC phát sinh trong xe đang ở Điều phối
 * - LSC ghi đầy đủ vào sheet LSC_PHAT_SINH, KHÔNG ghi đè công việc cũ.
 * - Worker hiện hữu xem xét 1 yêu cầu/lượt, chỉ kích hoạt nếu chứng minh
 *   xe chưa ra cổng, CVDV còn Trong xưởng và việc cũ đã hoàn thành.
 * - Không tạo Data EM/SCC Hoàn thành, không tự sinh giờ nghiệp vụ.
 * - Thiếu chứng cứ: để nguyên hàng đợi, có lý do để người quản lý xem.
 * - Không thêm trigger. Auto OFF trước khi quản lý chạy preflight rồi bật.
 **************************************************************************/
const P053_AUTO_CFG = Object.freeze({
  ENABLE_KEY:'TPMH_P053_LSC_AUTO_PROMOTE_ENABLED',
  TAB:'LSC_PHAT_SINH',
  BUDGET_GUARD_MS:6500,
  SKIP_SECONDS:120,
  MAX_CHECK:12
});
function p053AutoEnabled_() {
  return PropertiesService.getScriptProperties().getProperty(P053_AUTO_CFG.ENABLE_KEY)==='ON';
}
function p053QueueReason_(sh, row, reason, nextStatus) {
  if(!sh || row<2)return;
  if(nextStatus)sh.getRange(row,9).setValue(nextStatus);
  sh.getRange(row,12,1,2).setValues([[new Date(),String(reason||'').slice(0,490)]]);
}
function p053QueueOwner_(ss, plate) {
  const owner=p053DispatchOwnerWeb_(ss,plate);
  const wait=p050WaitingOwnerForWeb_(ss,plate);
  return {owner:owner,wait:wait};
}
function p053RawContextForProof_(ss, item, wait, nextId) {
  const payload=item.payload;
  const dp=ss.getSheetByName(CFG.SHEET_DISPATCH);
  const wd=ss.getSheetByName(CFG.SHEET_WAITING_DELIVERY);
  const em=ss.getSheetByName(CFG.SHEET_DATA_EM);
  const scc=ss.getSheetByName(CFG.SHEET_DATA_SCC);
  if(!dp||!wd||!em||!scc)throw new Error('P053_REQUIRED_SHEET_MISSING');
  if(!payload || !payload.ngay || !payload.gioNhan || !(payload.cv || payload.inspectionType))
    throw new Error('P053_NEW_LSC_INCOMPLETE');
  const newDate=parseNgayWeb_(payload.ngay);
  const newReceive=parseGioWeb_(payload.gioNhan);
  if(!newDate || !newReceive)throw new Error('P053_NEW_DATE_OR_CLOCK_INVALID');
  const rec={date:newDate,receiveTime:newReceive,advisor:item.advisor,
    plate:item.plate,model:item.model,visitId:nextId};
  const oldRec={visitId:wait.visitId,plate:item.plate};
  const oldEmRow=findDataRow_(em,oldRec);
  const oldSccRow=findDataRow_(scc,oldRec);
  if(!oldEmRow&&!oldSccRow)throw new Error('P053_OLD_DATA_NOT_FOUND');
  [[em,oldEmRow,'EM'],[scc,oldSccRow,'SCC']].forEach(function(v){
    if(!v[1])return;
    if(normalizePlate_(v[0].getRange(v[1],2).getDisplayValue())!==item.plate ||
      !isComplete_(v[0].getRange(v[1],13).getDisplayValue()))
      throw new Error('P053_OLD_JOB_NOT_COMPLETED_'+v[2]);
  });
  const ctx={oldId:wait.visitId,newId:nextId,plate:item.plate,rec:rec,
    oldWaiting:wd.getRange(wait.row,1,1,8).getValues()[0],
    em:em,scc:scc,oldEmRow:oldEmRow,oldSccRow:oldSccRow};
  return p049UndeliveredSourceProof_(ss,ctx);
}
function p053FormatStagedRow_(ss, sh, row, nextId, item) {
  const p=item.payload;
  const date=parseNgayWeb_(p.ngay), delivery=parseNgayWeb_(p.ngayGiao);
  const received=parseGioWeb_(p.gioNhan), promised=parseGioWeb_(p.gioGiao);
  if(!date||!delivery||!received||!promised)throw new Error('P053_LSC_DATE_TIME_INVALID');
  const scan=scanDispatchForInsertWeb_(sh,item.plate);
  if(scan.duplicateRow)throw new Error('P053_DISPATCH_OWNER_REAPPEARED');
  if(row!==scan.emptyRow)throw new Error('P053_EMPTY_ROW_MOVED');
  if(row>sh.getMaxRows())sh.insertRowAfter(sh.getMaxRows());
  sh.getRange(row,WEBAPP_CONFIG.COL.VISIT_ID).setNumberFormat('@').setValue(nextId);
  applyFastFormatsWeb_(sh,row);
  const work=p.cv||p.inspectionType;
  const lscExtra=[p.ghiChu||'',p.gtgt?('GTGT LSC: '+p.gtgt):'',
    p.soMon?('Số món LSC: '+p.soMon):''].filter(Boolean).join(' | ');
  const note=(lscExtra ? lscExtra+' | ':'')+
    '[P053 LSC PHÁT SINH ĐÃ ĐỐI CHIẾU | YÊU CẦU='+item.requestId+
    ' | chủ cũ='+item.ownerId+' | Chờ phân công thực tế]';
  // Giao việc theo đúng LUỒNG đã được phân loại: 1K/Hỗ trợ SCC là EM,
  // SC cơ bản/3D là SCC, còn BD chưa phân công thì chờ giao việc chung.
  // KHÔNG ghi 'Hoàn thành' hoặc tạo bất kỳ giờ kết thúc nào.
  const route=classifyWork_({workRaw:work,inspectionType:p.inspectionType||'',
    emGroup:'',inspector:'',ktv1:'',ktv2:'',itemCount:''});
  const emPending=route.em.active?'Chờ giao việc':'';
  const sccPending=(route.scc.active||(!route.em.active&&!route.scc.active))
    ?'Chờ giao việc':'';
  // GTGT, số món, chi tiết vẫn giữ đủ trong ledger; người điều phối chỉ
  // cần hoàn thiện phân công thực tế, không làm rơi thông tin chụp LSC.
  const values=[[date,item.plate,item.advisor,item.model,'','','','',work,
    p.capBD||'','','',received,delivery,promised,emPending,sccPending,'',note]];
  setDispatchRowValuesWeb_(sh,row,sh.getRange(row,1,1,19),values,work,'');
  if(p.inspectionType||p.qcLevel||p.noiDungKT){
    sh.getRange(row,9).setNote([
      p.inspectionType?'LOẠI KIỂM TRA: '+p.inspectionType:'',
      p.qcLevel?'CẤP QC: '+p.qcLevel:'',
      p.noiDungKT?'NỘI DUNG KIỂM TRA:\n'+p.noiDungKT:''
    ].filter(Boolean).join('\n'));
  }
  SpreadsheetApp.flush();
  const verify=sh.getRange(row,1,1,22).getDisplayValues()[0];
  if(normalizePlate_(verify[1])!==item.plate || clean_(verify[21])!==nextId ||
     (emPending && normalizeKey_(verify[15])!=='cho giao viec') ||
     (sccPending && normalizeKey_(verify[16])!=='cho giao viec'))
    throw new Error('P053_NEW_DISPATCH_NOT_VERIFIED');
}
function p053OneQueueTickUnderLock_(ss, deadline) {
  const out={enabled:p053AutoEnabled_(),status:'OFF',checked:0,activated:0};
  if(!out.enabled)return out;
  if(typeof p053QueueSheetWeb_!=='function' || typeof p053DispatchOwnerWeb_!=='function'){
    out.status='WEBAPP_HELPERS_NOT_INSTALLED';return out;
  }
  if(isWebAppConfirmPriorityActive_()) {out.status='WEBAPP_PRIORITY';return out;}
  if(Date.now()+P053_AUTO_CFG.BUDGET_GUARD_MS>=Number(deadline||0)){
    out.status='TIME_BUDGET_SKIP';return out;
  }
  const ledger=p053QueueSheetWeb_(ss,false);
  if(!ledger||ledger.getLastRow()<2){out.status='EMPTY';return out;}
  const end=Math.min(ledger.getLastRow(),1501);
  const rows=ledger.getRange(2,1,end-1,14).getDisplayValues();
  const candidates=[];
  rows.forEach(function(v,i){
    if(['CHO_KET_THUC_CONG_VIEC_CU','CHO_XAC_MINH_NGUON','DANG_KICH_HOAT'].indexOf(v[8])<0)return;
    candidates.push({row:i+2,requestId:v[1],plate:normalizePlate_(v[2]),advisor:v[3],
      model:v[4],ownerId:v[5],status:v[8],raw:v[9],newId:v[10]});
  });
  if(!candidates.length){out.status='NO_PENDING';return out;}
  // P0.5.8 FAIR-QUEUE: avoid starvation if first 12 old requests remain blocked.
  // ScriptLock is already held by the existing worker; no new trigger.
  const props=PropertiesService.getScriptProperties();
  const cursorKey='TPMH_P053_LSC_SCAN_CURSOR';
  const rawCursor=Number(props.getProperty(cursorKey)||0);
  const startCursor=(Number.isFinite(rawCursor)&&rawCursor>=0?Math.floor(rawCursor):0)
    % candidates.length;
  const scanLimit=Math.min(candidates.length,P053_AUTO_CFG.MAX_CHECK);
  props.setProperty(cursorKey,String((startCursor+scanLimit)%candidates.length));
  for(let i=0;i<scanLimit;i++){
    const item=candidates[(startCursor+i)%candidates.length];out.checked++;
    // Nếu transaction đã COMMITTED nhưng lần ghi trạng thái ledger bị ngắt,
    // đọc audit P046 để tự kết thúc idempotent (kể cả lượt mới đã chuyển Chờ giao).
    if(item.newId){
      const completedAudit=p046Audit_(ss,false);
      if(completedAudit && p046AlreadyCommitted_(completedAudit,item.newId)){
        p053QueueReason_(ledger,item.row,'Đã xác minh P046 COMMITTED / '+item.newId,'DA_KICH_HOAT');
        out.status='RECOVERED_FROM_AUDIT';out.activated=1;return out;
      }
    }
    const owner=p053QueueOwner_(ss,item.plate);
    if(owner.owner){
      if(owner.owner.visitId===item.ownerId){out.status='WAIT_OLD_DISPATCH';continue;}
      if(owner.owner.visitId!==item.newId){out.status='WAIT_OTHER_DISPATCH';continue;}
    }
    if(!owner.wait){
      out.status=item.newId?'WAIT_AUDIT_COMMIT_REQUIRED':'WAIT_OLD_WAITING_NOT_FOUND';
      continue;
    }
    if(owner.wait.visitId!==item.ownerId){out.status='WAIT_VISIT_CHANGED';continue;}
    if(normalizeKey_(owner.wait.advisor)!==normalizeKey_(item.advisor) ||
       normalizeKey_(owner.wait.model)!==normalizeKey_(item.model)){
      out.status='WAIT_IDENTITY_MISMATCH';continue;
    }
    const cooldown='P053_BLOCK_'+item.requestId;
    if(CacheService.getScriptCache().get(cooldown)){out.status='COOLDOWN';continue;}
    if(Date.now()+P053_AUTO_CFG.BUDGET_GUARD_MS>=deadline){out.status='TIME_BUDGET_SKIP';break;}
    try{
      item.payload=JSON.parse(item.raw);
      if(item.payload.ownerVisitId!==item.ownerId ||
         normalizePlate_(item.payload.bienSo)!==item.plate)
        throw new Error('P053_QUEUE_IDENTITY_CORRUPTED');
      const id=item.newId||generateUniqueVisitId_(ss); // P0.5.5: existing globally unique ID allocator
      // Chứng minh thời gian và trạng thái liên tục trước khi tạo Điều phối mới.
      const proof=p053RawContextForProof_(ss,item,owner.wait,id);
      if(isWebAppConfirmPriorityActive_() || Date.now()+2000>=deadline){
        out.status='TIME_OR_PRIORITY_YIELD';return out;
      }
      if(!item.newId){
        ledger.getRange(item.row,11).setNumberFormat('@').setValue(id);
        ledger.getRange(item.row,9).setValue('DANG_KICH_HOAT');
        SpreadsheetApp.flush();
      }
      const dispatch=ss.getSheetByName(CFG.SHEET_DISPATCH);
      let row=owner.owner&&owner.owner.visitId===id ? owner.owner.row:0;
      if(!row){
        const scan=scanDispatchForInsertWeb_(dispatch,item.plate);
        if(scan.duplicateRow)throw new Error('P053_NEW_OWNER_CONFLICT');
        row=scan.emptyRow;
        p053FormatStagedRow_(ss,dispatch,row,id,item);
      }
      // P046 xác minh lại toàn bộ nguồn + rollback Chờ giao trong cùng ScriptLock.
      const committed=p046Commit_(ss,id,item.plate,proof);
      if(!committed || committed.status!=='COMMITTED')
        throw new Error('P053_P046_COMMIT_NOT_CONFIRMED');
      p053QueueReason_(ledger,item.row,
        'Đã tạo công việc phát sinh mới '+id+'; giữ lịch sử cũ. Chờ KTV phân công.',
        'DA_KICH_HOAT');
      out.status='COMMITTED';out.activated=1;out.newId=id;out.plate=item.plate;
      console.log('[P053_AUTO_COMMITTED] '+JSON.stringify(out));
      return out;
    }catch(err){
      out.status='WAIT_VERIFICATION';out.reason=String(err.message||err);
      console.warn('[P053_AUTO_NOT_COMMITTED] '+JSON.stringify({plate:item.plate,reason:out.reason}));
      // Giữ nguyên ledger để có thể thử lại. Không xóa dữ liệu công việc cũ.
      // Nếu dòng mới đã tạo mà giao dịch chưa chốt, cần đối soát trước thao tác tay.
      const nextStatus=item.newId?'DANG_KICH_HOAT':'CHO_XAC_MINH_NGUON';
      p053QueueReason_(ledger,item.row,out.reason,nextStatus);
      CacheService.getScriptCache().put(cooldown,'1',P053_AUTO_CFG.SKIP_SECONDS);
      return out;
    }
  }
  return out;
}
function p053ProductionPreflightReadOnly() {
  const ss=laySpreadsheetWeb_();
  const s=ss.getSheetByName('LSC_PHAT_SINH');
  const triggers=ScriptApp.getProjectTriggers().filter(function(t){
    return t.getHandlerFunction()==='runPendingVisitWorker';}).length;
  const r={readOnly:true,version:'P0.5.3',autoEnabled:p053AutoEnabled_(),
    workerTriggerCount:triggers,ledgerExists:!!s,
    queued:s?Math.max(0,s.getLastRow()-1):0,
    requiredHelpers:['p053QueueSheetWeb_','p053DispatchOwnerWeb_',
      'p049UndeliveredSourceProof_','p046Commit_'].filter(function(x){
      return typeof ({p053QueueSheetWeb_:p053QueueSheetWeb_,p053DispatchOwnerWeb_:p053DispatchOwnerWeb_,
       p049UndeliveredSourceProof_:p049UndeliveredSourceProof_,p046Commit_:p046Commit_})[x]!=='function';})};
  r.safeToEnable=triggers===1 && r.requiredHelpers.length===0;
  console.log('[P053_PREFLIGHT_READ_ONLY] '+JSON.stringify(r));
  return r;
}
function p053EnableAutoQueueOnce() {
  const check=p053ProductionPreflightReadOnly();
  if(!check.safeToEnable)throw new Error('P053_PREFLIGHT_NOT_READY');
  PropertiesService.getScriptProperties().setProperty(P053_AUTO_CFG.ENABLE_KEY,'ON');
  console.log('[P053_ENABLED] '+JSON.stringify({workers:check.workerTriggerCount}));
  return {enabled:true,workerTriggerCount:check.workerTriggerCount};
}
function p053DisableAutoQueue() {
  PropertiesService.getScriptProperties().setProperty(P053_AUTO_CFG.ENABLE_KEY,'OFF');
  return {enabled:false};
}

/***************************************************************
 * TPMH P0.5.5 – IDENTITY / DATA WRITE HARDENING (CANDIDATE)
 * Production-safe ADDITIVE guards: no triggers, no new sheets,
 * no data modification from Preflight; no invented completion/exit.
 * This source was user-provided; deployment and live-script parity
 * are NOT established until the owner verifies Apps Script version.
 ***************************************************************/
/** Validate active dispatch ownership for a visitId before writing any EM/SCC.
 * Returns the exact active owner count. Historical completed work may no
 * longer be in Dispatch, so 0 owners is allowed; any conflicting owner is NOT.
 * Caller must already hold the existing worker/onEdit/WebApp ScriptLock.
 */
function p055AssertDataWriteOwner_(ss, rec, destination) {
  if (!ss || !rec) throw new Error('P055_BAD_WRITE_CONTEXT');
  const id = clean_(rec.visitId), plate = normalizePlate_(rec.plate);
  const advisor = normalizeKey_(rec.advisor);
  if (!id || !plate || !advisor) throw new Error('P055_VISIT_IDENTITY_REQUIRED');
  const dp = ss.getSheetByName(CFG.SHEET_DISPATCH);
  if (!dp) throw new Error('P055_DISPATCH_SHEET_MISSING');
  const first = CFG.FIRST_DATA_ROW;
  const last = dp.getLastRow();
  if (last < first) return 0;
  const values = dp.getRange(first, CFG.DP.BIEN_SO, last-first+1,
    CFG.DP.VISIT_ID-CFG.DP.BIEN_SO+1).getDisplayValues();
  let owners = 0;
  for (let i=0; i<values.length; i++) {
    const v = values[i], rowId=clean_(v[CFG.DP.VISIT_ID-CFG.DP.BIEN_SO]);
    if (rowId!==id) continue;
    owners++;
    const ownerPlate=normalizePlate_(v[0]);
    const ownerAdvisor=normalizeKey_(v[CFG.DP.CVDV-CFG.DP.BIEN_SO]);
    if (ownerPlate!==plate || ownerAdvisor!==advisor) {
      throw new Error('P055_CROSS_VEHICLE_ID_BLOCKED: ' + destination +
        ' visitId=' + id + ' dispatchRow=' + (first+i) +
        ' actual=' + ownerPlate + '/' + ownerAdvisor +
        ' requested=' + plate + '/' + advisor);
    }
  }
  if (owners>1) throw new Error('P055_DUPLICATE_ACTIVE_VISIT_ID_BLOCKED: '+id);
  return owners;
}
/** Called ONLY when a new Data row is about to be inserted. Stops cross-side
 * collisions before append, including if old Dispatch owner was already moved.
 */
function p055AssertOtherDataOwnerBeforeInsert_(ss,rec,destination) {
  const otherName=destination===CFG.SHEET_DATA_EM ? CFG.SHEET_DATA_SCC : CFG.SHEET_DATA_EM;
  const sh=ss.getSheetByName(otherName);
  if (!sh) throw new Error('P055_OTHER_DATA_SHEET_MISSING: '+otherName);
  const last=sh.getLastRow(); if (last<2) return true;
  const ids=sh.getRange(2,15,last-1,1).getDisplayValues();
  const id=clean_(rec.visitId),p=normalizePlate_(rec.plate),a=normalizeKey_(rec.advisor);
  let matched=0;
  for(let i=0;i<ids.length;i++){
    if(clean_(ids[i][0])!==id)continue;
    matched++;
    const identity=sh.getRange(i+2,2,1,2).getDisplayValues()[0];
    if(normalizePlate_(identity[0])!==p || normalizeKey_(identity[1])!==a){
      throw new Error('P055_CROSS_SIDE_IDENTITY_BLOCKED: '+otherName+'!'+(i+2)+
        ' / visitId='+id+'; refuse '+destination+' append');
    }
  }
  if(matched>1)throw new Error('P055_OTHER_DATA_DUPLICATE_ID: '+id);
  return true;
}
/** Read-only, targeted operational report; no status, time, ID or queue writes.
 * Includes completed-without-Data cases as BLOCKED; never fabricates an EM row.
 */
function TPMH_P055_IdentityPreflightReadOnly(){
  const ss=laySpreadsheetWeb_();
  const dp=ss.getSheetByName(CFG.SHEET_DISPATCH);
  const em=ss.getSheetByName(CFG.SHEET_DATA_EM);
  const scc=ss.getSheetByName(CFG.SHEET_DATA_SCC);
  if(!dp||!em||!scc)throw new Error('P055_REQUIRED_SHEET_MISSING');
  const first=CFG.FIRST_DATA_ROW,last=dp.getLastRow();
  const data=last>=first?dp.getRange(first,1,last-first+1,CFG.DP.VISIT_ID).getDisplayValues():[];
  function index_(sh){
    const result=Object.create(null);const l=sh.getLastRow();
    if(l<2)return result;
    sh.getRange(2,1,l-1,15).getDisplayValues().forEach(function(r,i){
      const id=clean_(r[14]);if(!id)return;
      (result[id]||(result[id]=[])).push({row:i+2,plate:normalizePlate_(r[1]),
        advisor:normalizeKey_(r[2]),status:clean_(r[12])});
    });
    return result;
  }
  const EM=index_(em),SCC=index_(scc);
  const byId=Object.create(null),report={version:'P0.5.5_IDENTITY_GUARD',
    readOnly:true,duplicateDispatch:[],crossVehicleData:[],
    missingEM:[],missingSCC:[],completedWithoutEM:[],completedWithoutSCC:[]};
  data.forEach(function(v,i){
    const plate=normalizePlate_(v[CFG.DP.BIEN_SO-1]);if(!plate)return;
    const visitId=clean_(v[CFG.DP.VISIT_ID-1]);
    const advisor=normalizeKey_(v[CFG.DP.CVDV-1]);
    const row=first+i;
    if(visitId)(byId[visitId]||(byId[visitId]=[])).push({row:row,plate:plate,advisor:advisor});
    if(!visitId)return;
    [['EM',EM,CFG.DP.NHOM_EM,CFG.DP.TD_EM],
     ['SCC',SCC,CFG.DP.KTV1,CFG.DP.TD_SCC]].forEach(function(spec){
      const kind=spec[0],index=spec[1],status=clean_(v[spec[3]-1]);
      const assigned=kind==='EM' ? !!clean_(v[CFG.DP.NHOM_EM-1]) :
        !!(clean_(v[CFG.DP.NGUOI_KT-1]) || clean_(v[CFG.DP.KTV1-1]) ||
           clean_(v[CFG.DP.KTV2-1]));
      const matches=index[visitId]||[];
      const foreign=matches.filter(z=>z.plate!==plate||z.advisor!==advisor);
      if(foreign.length)report.crossVehicleData.push({row:row,plate:plate,visitId:visitId,kind:kind,foreign:foreign});
      if(assigned && status && !matches.some(z=>z.plate===plate&&z.advisor===advisor)){
        const item={row:row,plate:plate,visitId:visitId,status:status};
        report['missing'+kind].push(item);
        if(isComplete_(status))report['completedWithout'+kind].push(item);
      }
    });
  });
  Object.keys(byId).forEach(id=>{if(byId[id].length>1)
    report.duplicateDispatch.push({visitId:id,owners:byId[id]});});
  console.log('[TPMH_P055_IDENTITY_PREFLIGHT] '+JSON.stringify(report));
  return report;
}



/*********************************************************
 * TPMH P0.5.7 - COMPLETION RECEIPT & SOURCE-BOUND RECOVERY
 * PATCH 11/11/2026 - NO NEW TRIGGER / NO NEW SHEET.
 *
 * - onEdit P/Q preserves a small, visit-bound second receipt independent
 *   of dirty queue. It stores the observed onEdit event time only.
 * - Existing one-minute worker can replay a still-valid receipt ONLY
 *   while the same visit / plate / advisor / status is present.
 * - Historical Completed without event: remain BLOCKED; a human may
 *   use the dedicated menu after verifying the REAL completion time.
 * - No invented timestamp, no physical-exit decision, no status downgrade,
 *   no blind cleanup of historical rows, no additional triggers.
 *********************************************************/
const TPMH_P057_CFG = Object.freeze({
  RECEIPT_PREFIX: 'TPMH_P057_COMPLETE_',
  MAX_EDIT_ROWS: 12,
  WORKER_MAX_RECEIPTS: 8,
  WORKER_MAX_REPLAYS: 1,
  WORKER_BUDGET_MS: 2500,
  RECEIPT_MAX_AGE_DAYS: 14,
  // Independent scheduling metadata only; never changes a completion receipt.
  WORKER_CURSOR_KEY: 'TPMH_P057_RECEIPT_SCAN_CURSOR_V1',
});
function p057ReceiptKey_(id, side) {
  if (!/^[0-9a-f-]{36}$/i.test(clean_(id)) ||
      (side !== 'EM' && side !== 'SCC')) return '';
  return TPMH_P057_CFG.RECEIPT_PREFIX + side + '_' + clean_(id);
}
/** Called synchronously INSIDE a true P/Q onEdit event, before ScriptLock.
 * Never fabricates an event for WebApp/script edits (Apps Script does not
 * fire onEdit for programmatic setValue). Every saved receipt is 1 visit.
 */
function p057CaptureStatusReceiptFromEdit_(e) {
  if (!e || !e.range || e.range.getSheet().getName() !== CFG.SHEET_DISPATCH)
    return {stored:0,cleared:0,skipped:0};
  const sh=e.range.getSheet();
  const first=Math.max(CFG.FIRST_DATA_ROW,e.range.getRow());
  const last=e.range.getLastRow();
  const out={stored:0,cleared:0,skipped:0};
  if (last<first) return out;
  const count=Math.min(last-first+1,TPMH_P057_CFG.MAX_EDIT_ROWS);
  const props=PropertiesService.getScriptProperties();
  const at=clean_(e.p044CapturedEventAt||'');
  const t=new Date(at);
  if (!at || isNaN(t.getTime()) || t.getTime()>Date.now()+60000)
    throw Error('P057_DIRECT_EVENT_INVALID');
  const columns=[{col:CFG.DP.TD_EM,side:'EM'},
                 {col:CFG.DP.TD_SCC,side:'SCC'}]
    .filter(function(x){return editTouchesColumn_(e.range,x.col);});
  for(let i=0;i<count;i++) {
    const row=first+i;
    const cells=sh.getRange(row,1,1,CFG.DP.VISIT_ID).getDisplayValues()[0];
    const plate=normalizePlate_(cells[CFG.DP.BIEN_SO-1]);
    const visitId=clean_(cells[CFG.DP.VISIT_ID-1]);
    const advisor=normalizeKey_(cells[CFG.DP.CVDV-1]);
    const rawDate=sh.getRange(row,CFG.DP.NGAY).getValue();
    const dateKey=cvdvDateKey_(rawDate,sh.getParent());
    if(!plate||!advisor||!dateKey||!visitId) {out.skipped++;continue;}
    columns.forEach(function(x){
      const key=p057ReceiptKey_(visitId,x.side);
      if(!key){out.skipped++;return;}
      const status=cells[x.col-1];
      if(!isComplete_(status)) {
        const old=props.getProperty(key);
        if(old) {
          try {
            const oldProof=JSON.parse(old);
            if(oldProof.plate!==plate||oldProof.visitId!==visitId) {
              out.skipped++; return;
            }
            props.deleteProperty(key);out.cleared++;
          } catch(err) {out.skipped++;}
        }
        return;
      }
      const receipt={version:'P0.5.7',source:'USER_STATUS_ONEDIT',
        side:x.side,visitId:visitId,plate:plate,advisor:advisor,
        dateKey:dateKey,eventAt:at,createdAt:new Date().toISOString()};
      props.setProperty(key,JSON.stringify(receipt));out.stored++;
    });
  }
  if(last-first+1>count) out.skipped+=last-first+1-count;
  console.log('[P057_COMPLETION_RECEIPT] '+JSON.stringify(out));
  return out;
}
/** Delete a receipt only AFTER a Completed Data row is successfully written.
 * A failure to clean up a receipt must never fail the original operation.
 */
function p057ClearCompletedReceiptAfterDataCommit_(rec,isEM) {
  const side=isEM?'EM':'SCC',key=p057ReceiptKey_(rec.visitId,side);
  if(!key)return;
  try {
    const props=PropertiesService.getScriptProperties();
    const raw=props.getProperty(key);if(!raw)return;
    const proof=JSON.parse(raw);
    if(proof.visitId===clean_(rec.visitId) &&
       proof.plate===normalizePlate_(rec.plate) &&
       proof.advisor===normalizeKey_(rec.advisor)) props.deleteProperty(key);
  } catch(err) {
    console.warn('[P057_RECEIPT_CLEANUP_DEFERRED] '+String(err.message||err));
  }
}
/** Existing worker ScriptLock MUST be held by caller. Exactly one bounded
 * event replay, guarded by source-bound receipt and original pending-queue
 * processor. No direct new Completed Data append here.
 */
function p057RecoverCompletedReceiptUnderLock_(ss,deadline) {
  const startedAt=Date.now();
  const out={checked:0,replayed:0,cleared:0,blocked:0,routeMismatches:0,
    routeMismatchDetails:[],errors:[],items:[],yieldedForBudget:false,
    cursorAdvanced:false,elapsedMs:0};
  if(Date.now()>=deadline||isWebAppConfirmPriorityActive_())return out;
  const props=PropertiesService.getScriptProperties();
  const all=props.getProperties();
  // Parse once. Sort by ORIGINAL confirmation event time, oldest first.
  // Invalid receipt metadata sorts last; the original evidence is unchanged.
  const entries=Object.keys(all).filter(function(k){
    return k.indexOf(TPMH_P057_CFG.RECEIPT_PREFIX)===0;
  }).map(function(key){
    let proof=null,sortAt='9999-12-31T23:59:59.999Z';
    try {
      proof=JSON.parse(all[key]);
      const date=new Date(proof.eventAt||proof.createdAt||'');
      if(!isNaN(date.getTime()))sortAt=date.toISOString();
    } catch(ignored) {}
    return {key:key,proof:proof,raw:all[key],sortAt:sortAt};
  }).sort(function(a,b){
    return a.sortAt.localeCompare(b.sortAt)||a.key.localeCompare(b.key);
  });
  if(!entries.length)return out;
  const dp=ss.getSheetByName(CFG.SHEET_DISPATCH);
  if(!dp){out.errors.push('P057_DISPATCH_MISSING');return out;}
  // Round-robin AFTER an oldest-first initial pass. If the oldest proof is
  // permanently mismatched, the next 1-minute tick still reaches other visits.
  // One short Script Property stores only scheduling metadata, not a new proof.
  let startIndex=0;
  try {
    const cursor=JSON.parse(props.getProperty(TPMH_P057_CFG.WORKER_CURSOR_KEY)||'null');
    if(cursor&&typeof cursor.sortAt==='string'&&typeof cursor.key==='string') {
      const next=entries.findIndex(function(e){
        return e.sortAt>cursor.sortAt || (e.sortAt===cursor.sortAt&&e.key>cursor.key);
      });
      startIndex=next<0?0:next;
    }
  } catch(ignored) {} // A corrupt cursor may not suppress valid receipts.
  let lastAttempt=null;
  for(let j=0;j<entries.length && out.checked<TPMH_P057_CFG.WORKER_MAX_RECEIPTS &&
      out.replayed<TPMH_P057_CFG.WORKER_MAX_REPLAYS &&
      Date.now()<deadline && !isWebAppConfirmPriorityActive_();j++) {
    const e=entries[(startIndex+j)%entries.length];
    const p=e.proof, side=p&&p.side;
    const item={side:String(side||''),plate:String(p&&p.plate||''),
      visitId:String(p&&p.visitId||''),eventAt:String(p&&p.eventAt||''),
      result:'',ms:0};
    const itemStarted=Date.now();
    // Do not mark a receipt inspected if no time remains to do any work.
    if(deadline-Date.now()<200){out.yieldedForBudget=true;break;}
    out.checked++;
    const previousAttempt=lastAttempt;
    lastAttempt={sortAt:e.sortAt,key:e.key};
    try {
      if(!p||p.source!=='USER_STATUS_ONEDIT'||
         e.key!==p057ReceiptKey_(p.visitId,side)) {
        out.blocked++;item.result='INVALID_RECEIPT';continue;
      }
      const stamp=new Date(p.eventAt);
      if(isNaN(stamp.getTime()) || stamp.getTime()>Date.now()+60000 ||
         Date.now()-stamp.getTime()>TPMH_P057_CFG.RECEIPT_MAX_AGE_DAYS*86400000) {
        out.blocked++;item.result='INVALID_EVENT_TIME';continue;
      }
      const rows=resolveDispatchRowsByVisitId_(dp,p.visitId);
      const data=ss.getSheetByName(side==='EM'?CFG.SHEET_DATA_EM:CFG.SHEET_DATA_SCC);
      if(!data||rows.length!==1){out.blocked++;item.result='MISSING_OR_NONUNIQUE_OWNER';continue;}
      const rec=readDispatchRow_(dp,rows[0]);
      if(clean_(rec.visitId)!==p.visitId || normalizePlate_(rec.plate)!==p.plate ||
         normalizeKey_(rec.advisor)!==p.advisor ||
         cvdvDateKey_(rec.date,ss)!==p.dateKey ||
         !isComplete_(side==='EM'?rec.emStatus:rec.sccStatus)) {
        out.blocked++;item.result='SOURCE_CHANGED';continue;
      }
      // A receipt proves a status edit, not a business assignment.
      // Mr. Khoi without a separate SCC team is EM ONLY (P0.5.3.1).
      const route=classifyWork_(rec);
      const activeSide=side==='EM'?route.em:route.scc;
      if(!activeSide||activeSide.active!==true) {
        out.blocked++;out.routeMismatches++;item.result='ROUTE_MISMATCH';
        if(out.routeMismatchDetails.length<5) {
          const expected=[route.em.active?'EM':'',route.scc.active?'SCC':'']
            .filter(Boolean).join('+')||'NONE';
          out.routeMismatchDetails.push('P057_ROUTE_MISMATCH '+side+' '+
            p.visitId+' '+p.plate+' actual='+expected);
        }
        continue; // Preserve original receipt as evidence, try next visit.
      }
      const stored=findDataRow_(data,rec);
      if(stored && isComplete_(data.getRange(stored,13).getDisplayValue())) {
        if(props.getProperty(e.key)===e.raw){props.deleteProperty(e.key);out.cleared++;}
        item.result='ALREADY_COMPLETED';continue;
      }
      if(stamp.getTime()<rec.date.getTime()) {
        out.blocked++;item.result='EVENT_BEFORE_INTAKE';continue;
      }
      // Retain the original P/Q proof; the worker clock never supplies a stamp.
      const sideCol=side==='EM'?CFG.DP.TD_EM:CFG.DP.TD_SCC;
      const eventMap={};eventMap[String(sideCol)]=p.eventAt;
      const otherSide=side==='EM'?'SCC':'EM';
      const otherKey=p057ReceiptKey_(p.visitId,otherSide);
      const otherRaw=otherKey&&props.getProperty(otherKey);
      if(otherRaw) {
        try {
          const second=JSON.parse(otherRaw);
          const secondAt=new Date(second.eventAt);
          if(second.source==='USER_STATUS_ONEDIT' && second.side===otherSide &&
             second.visitId===p.visitId && second.plate===p.plate &&
             second.advisor===p.advisor && second.dateKey===p.dateKey &&
             !isNaN(secondAt.getTime()) && secondAt.getTime()>=rec.date.getTime() &&
             secondAt.getTime()<=Date.now()+60000 &&
             Date.now()-secondAt.getTime()<=TPMH_P057_CFG.RECEIPT_MAX_AGE_DAYS*86400000 &&
             (otherSide==='EM'?route.em.active:route.scc.active)===true &&
             isComplete_(otherSide==='EM'?rec.emStatus:rec.sccStatus)) {
            const secondCol=otherSide==='EM'?CFG.DP.TD_EM:CFG.DP.TD_SCC;
            eventMap[String(secondCol)]=second.eventAt;
          }
        } catch(ignored) {}
      }
      const eventCols=Object.keys(eventMap).map(Number).sort(function(a,b){return a-b;});
      const job={sheet:CFG.SHEET_DISPATCH,visitId:p.visitId,plate:p.plate,
        rowHint:rows[0],colStart:eventCols[0],numCols:eventCols[eventCols.length-1]-eventCols[0]+1,
        completionEventAt:p.eventAt,completionEventSides:eventCols,
        p043CompletionEvents:eventMap};
      // Receipt writer runs BEFORE ScriptLock on true user P/Q edits. Recheck
      // the original snapshot so a concurrent newer click cannot be replayed
      // with an older confirmation timestamp from our worker snapshot.
      if(props.getProperty(e.key)!==e.raw) {
        out.blocked++;item.result='RECEIPT_SUPERSEDED';continue;
      }
      if(Date.now()>=deadline||isWebAppConfirmPriorityActive_()) {
        out.yieldedForBudget=true;item.result='YIELD_BEFORE_REPLAY';
        lastAttempt=previousAttempt;break;
      }
      // Existing original processor owns all write validation and status.
      // Apps Script synchronous calls may overrun the requested budget; the
      // deadline is a cooperative check, NOT a guaranteed hard timeout.
      const result=processDirtyVisit_(ss,job);
      if(!result.done)throw result.error||Error('P057_REPLAY_NOT_DONE');
      SpreadsheetApp.flush();
      const check=findDataRow_(data,rec);
      if(!check || !isComplete_(data.getRange(check,13).getDisplayValue()))
        throw Error('P057_REPLAY_READBACK_FAILED '+side+' '+p.visitId);
      out.replayed++;item.result='REPLAYED';item.dataRow=check;
      // Source processor can independently clear the receipt after a commit.
      if(props.getProperty(e.key)===e.raw){props.deleteProperty(e.key);out.cleared++;}
    } catch(err) {
      const message=String(err.message||err).slice(0,220);
      out.errors.push(message);out.blocked++;
      item.result='ERROR';item.error=message;
    } finally {
      item.ms=Date.now()-itemStarted;
      out.items.push(item);
    }
  }
  // Cursor is only fairness bookkeeping. Failure must not hide replay result
  // or interfere with queue / WebApp / ScriptLock ownership.
  if(lastAttempt) {
    try {
      props.setProperty(TPMH_P057_CFG.WORKER_CURSOR_KEY,
        JSON.stringify(lastAttempt));
      out.cursorAdvanced=true;
    } catch(err) {
      console.warn('[P057_CURSOR_SAVE_FAILED] '+String(err.message||err).slice(0,140));
    }
  }
  out.elapsedMs=Date.now()-startedAt;
  if(out.checked||out.errors.length)
    console.log('[P057_RECEIPT_RECOVERY] '+JSON.stringify(out));
  return out;
}
/** Read-only: completed jobs missing a same-visit Data record. */
function p057CompletionGapAuditReadOnly() {
  const ss=laySpreadsheetWeb_();
  const dp=ss.getSheetByName(CFG.SHEET_DISPATCH);
  const em=ss.getSheetByName(CFG.SHEET_DATA_EM);
  const scc=ss.getSheetByName(CFG.SHEET_DATA_SCC);
  if(!dp||!em||!scc)throw Error('P057_REQUIRED_SHEET_MISSING');
  function ids_(sh){
    const n=sh.getLastRow()-1;
    return new Set(n>0?sh.getRange(2,15,n,1).getDisplayValues().map(function(r){
      return clean_(r[0]);
    }).filter(Boolean):[]);
  }
  const knownEM=ids_(em),knownSCC=ids_(scc);
  const rows=dp.getLastRow()>=CFG.FIRST_DATA_ROW?
    dp.getRange(CFG.FIRST_DATA_ROW,1,dp.getLastRow()-CFG.FIRST_DATA_ROW+1,
      CFG.DP.VISIT_ID).getDisplayValues():[];
  const props=PropertiesService.getScriptProperties();
  const out={version:'P0.5.7',readOnly:true,missingCompleted:[],openMissing:[],
    checked:0,sourceWrites:0,boardWrites:0,triggersChanged:0};
  rows.forEach(function(v,i){
    const id=clean_(v[CFG.DP.VISIT_ID-1]);
    const plate=normalizePlate_(v[CFG.DP.BIEN_SO-1]);
    if(!id||!plate)return;
    out.checked++;
    const inspector=clean_(v[CFG.DP.NGUOI_KT-1]);
    const ktv1=clean_(v[CFG.DP.KTV1-1]);
    const ktv2=clean_(v[CFG.DP.KTV2-1]);
    const khoiOnly=/^mr\.?\s*khoi$/.test(normalizeKey_(inspector)) && !ktv1 && !ktv2;
    [{side:'EM',active:!!clean_(v[CFG.DP.NHOM_EM-1])||khoiOnly,
      status:v[CFG.DP.TD_EM-1],ids:knownEM},
     {side:'SCC',active:!!(inspector||ktv1||ktv2)&&!khoiOnly,
      status:v[CFG.DP.TD_SCC-1],ids:knownSCC}].forEach(function(x){
      if(!x.active||x.ids.has(id))return;
      const info={row:CFG.FIRST_DATA_ROW+i,plate:plate,visitId:id,
        side:x.side,status:clean_(x.status),hasReceipt:
        !!props.getProperty(p057ReceiptKey_(id,x.side))};
      (isComplete_(x.status)?out.missingCompleted:out.openMissing).push(info);
    });
  });
  console.log('[P057_READ_ONLY] '+JSON.stringify(out));return out;
}
/** Admin-only intentional intervention for an already-Completed dispatch
 * row whose Data side is absent and whose original event was lost.
 * The operator MUST type a real observed completion date/time and confirm.
 */
function p057RecoverSelectedCompletedWithOperatorProof() {
  const ui=SpreadsheetApp.getUi();
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const active=ss&&ss.getActiveRange();
  if(!active||active.getSheet().getName()!==CFG.SHEET_DISPATCH||
     active.getNumRows()!==1||active.getRow()<CFG.FIRST_DATA_ROW)
    throw Error('P057_SELECT_ONE_DISPATCH_ROW');
  const selectedRow=active.getRow();
  const dp=active.getSheet();
  const preliminary=readDispatchRow_(dp,selectedRow);
  const plate=normalizePlate_(preliminary.plate),visitId=clean_(preliminary.visitId);
  if(!plate||!visitId)throw Error('P057_SOURCE_IDENTITY_REQUIRED');
  // A manually typed Date is parsed in the Apps Script project timezone.
  // Refuse ambiguous timezone setups rather than shifting 11:16 silently.
  const scriptTz=Session.getScriptTimeZone();
  const sheetTz=ss.getSpreadsheetTimeZone();
  const accepted=['Asia/Ho_Chi_Minh','Asia/Saigon','Asia/Bangkok'];
  if(accepted.indexOf(scriptTz)===-1||accepted.indexOf(sheetTz)===-1)
    throw Error('P057_TIMEZONE_MISMATCH: script='+scriptTz+' sheet='+sheetTz);
  const route=classifyWork_(preliminary);
  const candidates=[{side:'EM',active:route.em.active,status:preliminary.emStatus,
      name:CFG.SHEET_DATA_EM},
    {side:'SCC',active:route.scc.active,status:preliminary.sccStatus,
      name:CFG.SHEET_DATA_SCC}]
    .filter(function(x){
      if(!x.active||!isComplete_(x.status))return false;
      const data=ss.getSheetByName(x.name);
      if(!data)return false;
      const existing=findDataRow_(data,preliminary);
      if(!existing)return true;
      // Resume only an OPEN row staged by this same P057 recovery.
      const staged=data.getRange(existing,15).getNote();
      return !isComplete_(data.getRange(existing,13).getDisplayValue()) &&
        String(staged||'').indexOf('P057_OPEN_STAGE | ')===0 &&
        staged.indexOf('visitId='+visitId)!==-1 &&
        staged.indexOf(' | '+plate+' | '+x.side+' | ')!==-1;
    });
  if(candidates.length!==1)
    throw Error('P057_EXPECT_EXACTLY_ONE_COMPLETED_MISSING_DATA_SIDE; found='+candidates.length);
  const side=candidates[0].side;
  const answer=ui.prompt('P0.5.7 | XÁC NHẬN GIỜ HOÀN THÀNH',
    'Xe '+plate+' / '+side+' / visitId='+visitId+
    '\nNhập NGÀY GIỜ HOÀN THÀNH THỰC TẾ theo mẫu dd/MM/yyyy HH:mm.\n'+
    'Không được lấy giờ chạy worker hoặc giờ hẹn giao làm giờ hoàn thành.',
    ui.ButtonSet.OK_CANCEL);
  if(answer.getSelectedButton()!==ui.Button.OK)return {cancelled:true};
  const m=/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})$/.exec(answer.getResponseText().trim());
  if(!m)throw Error('P057_COMPLETION_TIME_EXPECTED_DDMMYYYY_HHMM');
  const day=+m[1],mon=+m[2],year=+m[3],hour=+m[4],minute=+m[5];
  const stamp=new Date(year,mon-1,day,hour,minute,0,0);
  if(stamp.getFullYear()!==year||stamp.getMonth()!==mon-1||
     stamp.getDate()!==day||stamp.getHours()!==hour||stamp.getMinutes()!==minute||
     stamp.getTime()>Date.now()+60000)
    throw Error('P057_COMPLETION_TIME_INVALID_OR_FUTURE');
  if(!(preliminary.date instanceof Date)||stamp.getTime()<preliminary.date.getTime())
    throw Error('P057_COMPLETION_BEFORE_DISPATCH_DATE');
  const confirmation=ui.alert('P0.5.7 | XÁC NHẬN NGHIỆP VỤ',
    'Bạn xác nhận '+side+' của xe '+plate+' thực tế đã Hoàn thành lúc '+
    answer.getResponseText().trim()+'?\n'+
    'Chỉ tạo Data đúng visitId từ bằng chứng này, KHÔNG ghi giờ ra cổng.',
    ui.ButtonSet.YES_NO);
  if(confirmation!==ui.Button.YES)return {cancelled:true};
  if(isWebAppConfirmPriorityActive_())throw Error('P057_WEBAPP_PRIORITY_ACTIVE');
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(3000))throw Error('P057_SCRIPT_LOCK_BUSY');
  try {
    if(isWebAppConfirmPriorityActive_())throw Error('P057_WEBAPP_PRIORITY_ACTIVE');
    const rows=resolveDispatchRowsByVisitId_(dp,visitId);
    if(rows.length!==1||rows[0]!==selectedRow)
      throw Error('P057_OWNER_MOVED_OR_DUPLICATED');
    const rec=readDispatchRow_(dp,selectedRow);
    if(clean_(rec.visitId)!==visitId||normalizePlate_(rec.plate)!==plate||
       normalizeKey_(rec.advisor)!==normalizeKey_(preliminary.advisor)||
       cvdvDateKey_(rec.date,ss)!==cvdvDateKey_(preliminary.date,ss))
      throw Error('P057_SOURCE_CHANGED_RECHECK');
    p04AssertPhysicalVisitOwnership_(dp,selectedRow,plate);
    const actualRoute=classifyWork_(rec);
    if(!(side==='EM'?actualRoute.em.active:actualRoute.scc.active)||
       !isComplete_(side==='EM'?rec.emStatus:rec.sccStatus))
      throw Error('P057_STATUS_OR_ASSIGNMENT_CHANGED');
    const data=ss.getSheetByName(side==='EM'?CFG.SHEET_DATA_EM:CFG.SHEET_DATA_SCC);
    if(!data)throw Error('P057_DATA_SHEET_MISSING');
    const existing=findDataRow_(data,rec);
    if(existing) {
      const staged=data.getRange(existing,15).getNote();
      if(isComplete_(data.getRange(existing,13).getDisplayValue()) ||
         String(staged||'').indexOf('P057_OPEN_STAGE | ')!==0 ||
         staged.indexOf('visitId='+visitId)===-1 ||
         staged.indexOf(' | '+plate+' | '+side+' | ')===-1 ||
         staged.indexOf(answer.getResponseText().trim())===-1)
        throw Error('P057_EXISTING_DATA_NOT_SAME_STAGED_PROOF');
    }
    const targetStamp=side==='EM'?rec.emConfirmationStamp:rec.sccConfirmationStamp;
    if(targetStamp instanceof Date&&!isNaN(targetStamp.getTime())&&
       Math.abs(targetStamp.getTime()-stamp.getTime())>60000)
      throw Error('P057_EXISTING_COMPLETION_STAMP_CONFLICT');
    const receiveText=dp.getRange(selectedRow,CFG.DP.GIO_NHAN).getDisplayValue();
    const re=/^(\d{1,2}):(\d{2})$/.exec(receiveText);
    if(re && +re[1]<=23 && +re[2]<=59 &&
       cvdvDateKey_(stamp,ss)===cvdvDateKey_(rec.date,ss) &&
       hour*60+minute < +re[1]*60 + +re[2])
      throw Error('P057_COMPLETION_BEFORE_RECEIVE_TIME');
    const completed=Object.assign({},rec,side==='EM'?{
      emConfirmationStamp:stamp}:{sccConfirmationStamp:stamp});
    const problems=side==='EM'?validateEMCompletion_(completed,actualRoute):
      validateSCCCompletion_(completed,actualRoute);
    if(problems.length)throw Error('P057_COMPLETION_VALIDATION_BLOCKED: '+problems.join(', '));
    p042ResolvedCompletionStamp_(completed,side==='EM');
    p055AssertDataWriteOwner_(ss,rec,data.getName());
    if(!existing)p055AssertOtherDataOwnerBeforeInsert_(ss,rec,data.getName());
    const workText=side==='EM'?actualRoute.em.workText:actualRoute.scc.workText;
    if(!existing) {
      const open=Object.assign({},completed,side==='EM'?{emStatus:'Đang làm'}:
        {sccStatus:'Đang làm'});
      syncDataForVisit_(open,workText,side==='EM');
      SpreadsheetApp.flush();
    }
    const dataRow=findDataRow_(data,completed);
    if(!dataRow)throw Error('P057_BOOTSTRAP_OPEN_FAILED');
    const provenance='P057_OPERATOR_VERIFIED | '+plate+' | '+side+
      ' | visitId='+visitId+' | Ngày giờ hoàn thành người vận hành xác nhận: '+
      answer.getResponseText().trim()+
      ' | KHÔNG phải giờ worker, giờ hẹn giao hay giờ ra cổng.';
    if(!existing)data.getRange(dataRow,15).setNote('P057_OPEN_STAGE | '+provenance);
    syncDataForVisit_(completed,workText,side==='EM');
    data.getRange(dataRow,12).setNote(provenance);
    data.getRange(dataRow,13).setNote(provenance);
    data.getRange(dataRow,15).setNote(provenance);
    const stampCol=side==='EM'?CFG.DP.MOC_TD_EM:CFG.DP.MOC_TD_SCC;
    const lastCol=side==='EM'?CFG.DP.LAST_EM_STATUS:CFG.DP.LAST_SCC_STATUS;
    dp.getRange(selectedRow,stampCol).setValue(stamp)
      .setNumberFormat('dd/MM/yyyy HH:mm:ss').setNote(provenance);
    dp.getRange(selectedRow,lastCol).setValue('Hoàn thành').setNote(provenance);
    SpreadsheetApp.flush();
    const check=findDataRow_(data,completed);
    if(check!==dataRow||!isComplete_(data.getRange(check,13).getDisplayValue()))
      throw Error('P057_FINAL_DATA_READBACK_FAILED');
    const result={version:'P0.5.7',committed:true,plate:plate,side:side,
      dispatchRow:selectedRow,dataRow:check,visitId:visitId,
      confirmedAt:answer.getResponseText().trim(),sourceWrites:1,
      exitTimeWritten:false};
    console.log('[P057_OPERATOR_RESTORED] '+JSON.stringify(result));
    return result;
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

/** Read-only, exact-visit audit for the two outstanding P04 cases. */
function p0574AuditP04TwoReadOnly() {
  const ss=laySpreadsheetWeb_();
  const dp=ss.getSheetByName(CFG.SHEET_DISPATCH);
  if(!dp)throw Error('P0574_DISPATCH_MISSING');
  const props=PropertiesService.getScriptProperties();
  const targets=[
    {plate:'51K02960',id:'993b3654-d45a-4f6d-8755-1e0836bc8d29',side:'EM'},
    {plate:'51G20232',id:'29fcf2f8-d30a-4ee9-9c75-82e1587abb1a',side:'SCC'}
  ];
  const out={version:'P0.5.7.4',readOnly:true,sourceWrites:0,
    triggerChanges:0,items:[]};
  targets.forEach(function(t) {
    const v={plate:t.plate,visitId:t.id,side:t.side};
    try {
      const rows=resolveDispatchRowsByVisitId_(dp,t.id);
      v.dispatchRows=rows;
      if(rows.length!==1){v.reason='MISSING_OR_MULTIPLE_DISPATCH';out.items.push(v);return;}
      const rec=readDispatchRow_(dp,rows[0]);
      v.identityMatches=normalizePlate_(rec.plate)===t.plate &&
        clean_(rec.visitId)===t.id;
      v.advisor=clean_(rec.advisor);
      v.dispatchStatus=clean_(t.side==='EM'?rec.emStatus:rec.sccStatus);
      v.lastCommittedStatus=clean_(t.side==='EM'?rec.lastEMStatus:rec.lastSCCStatus);
      const route=classifyWork_(rec);
      v.routeActive=t.side==='EM'?route.em.active:route.scc.active;
      const sh=ss.getSheetByName(t.side==='EM'?CFG.SHEET_DATA_EM:CFG.SHEET_DATA_SCC);
      if(sh){const row=findDataRow_(sh,rec);v.dataRow=row;
        if(row){v.dataStatus=sh.getRange(row,13).getDisplayValue();
          v.confirmationTime=sh.getRange(row,12).getDisplayValue();}}
      const key=p057ReceiptKey_(t.id,t.side);
      const raw=key&&props.getProperty(key);
      v.hasReceipt=!!raw;
      if(raw){try {
        const proof=JSON.parse(raw);
        v.receiptSource=String(proof.source||'');
        v.receiptEventAt=String(proof.eventAt||'');
        v.receiptIdentityMatches=proof.visitId===t.id&&
          proof.plate===t.plate;
      }catch(ignored){v.receiptParseError=true;}}
    }catch(err){v.error=String(err.message||err).slice(0,240);}
    out.items.push(v);
  });
  console.log('[P0574_P04_READ_ONLY] '+JSON.stringify(out));
  return out;
}
