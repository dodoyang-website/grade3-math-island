/**
 * ============================================================================
 * 🏆 國小三上數學期中總複習（1～4單元）運動會大闖關 — 雲端成績與迷思概念成長分析後端
 * ============================================================================
 * 【老師 3 步驟超簡單啟用教學】
 * 1. 開啟一個全新的空白 Google 試算表，點選上方選單「擴充功能」➔「Apps Script」。
 * 2. 將編輯器原本的內容刪掉，把這整份程式碼完整貼上，按 Ctrl+S 儲存。
 * 3. 在上方函式下拉選單選擇「setupSpreadsheet」，點擊「▶ 執行」（首次執行請點「審查權限 ➔ 進階 ➔ 前往（不安全）➔ 允許」）。
 *    ✨ 執行完畢後，試算表會自動產生 5 張精美彩色分析報表！
 * 4. 點擊右上角藍色按鈕「部署」➔「新增部署作業」：
 *    - 左側齒輪選取類型：「網頁應用程式」
 *    - 執行身分：「我」
 *    - 誰可以存取：務必選擇「所有人」（這樣學生平板免登入 Google 帳號也能自動上傳成績！）
 *    - 點擊「部署」，複製產生的「網頁應用程式網址 (https://script.google.com/macros/s/.../exec)」，
 *      貼回闖關網頁右上角的「👨‍🏫 教師後台」即可！
 * ============================================================================
 */

const MISCONCEPTION_META = {
  'M-01': { name: '位值、零的讀寫與錢幣換算', unit: '單元1：10000以內的數', advice: '建議搭配位值表定位練習中間有0的讀寫，並用圖卡操作500元/50元換鈔。' },
  'M-02': { name: '萬以內連續進位與跨零借位', unit: '單元2：10000以內的加減', advice: '整千或中間有0退位時，請學生在直式上方標註借位後的「9」與「10」。' },
  'M-03': { name: '毫米公分換算與尺規非零起點', unit: '單元3：毫米與數線', advice: '強化「1公分＝10毫米」量感，斷尺測量務必口訣提醒「終點減起點」。' },
  'M-04': { name: '數線間距判讀與左右移動', unit: '單元3：毫米與數線', advice: '先圈出相鄰兩刻度差確認「1格代表多少」，再判斷往右加、往左減及間隔數。' },
  'M-05': { name: '乘數進位、中間有零乘法與估算', unit: '單元4：乘法', advice: '注意如408×6十位進位占位、尾數補0個數，以及付款估算需採無條件進位。' },
  'M-06': { name: '跨單元兩步驟逆推與素養資訊整合', unit: '單元1～4綜合素養', advice: '引導圈出關鍵問題與單位，畫線段圖釐清「原有/剩下/相差」並排除無關數字。' }
};

const LEVEL_NAMES = {
  'level1': '第1關：初級（選手熱身場）',
  'level2': '第2關：中級（校園選拔賽）',
  'level3': '第3關：高級（分組錦標賽）',
  'level4': '第4關：超高級（運動會素養決賽）'
};

/**
 * 一鍵初始化全套 5 張雲端分析工作表
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename('🏆 三上數學期中總複習（1-4單元）全班成績與迷思概念成長分析表');

  // 1. 建立或取得 5 張工作表
  const sheetNames = [
    '1_作答明細',
    '2_全班首次闖關迷思診斷',
    '3_重測進步與迷思消除追蹤',
    '4_個人迷思與成長報告單',
    '5_全班四關進度總表'
  ];

  const sheets = {};
  sheetNames.forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sheets[name] = sh;
  });

  // 移除預設空白工作表（若存在）
  ['工作表1', 'Sheet1'].forEach(defaultName => {
    const dsh = ss.getSheetByName(defaultName);
    if (dsh && ss.getSheets().length > 1) {
      ss.deleteSheet(dsh);
    }
  });

  // 初始化表 1：作答明細
  const rawSh = sheets['1_作答明細'];
  if (rawSh.getLastRow() === 0) {
    const headers = [
      '提交時間', '學生唯一碼', '班級', '座號', '姓名',
      '關卡代號', '關卡名稱', '第幾次挑戰', '得分', '是否過關(>=90)', '作答時間(秒)',
      'M-01錯題數', 'M-02錯題數', 'M-03錯題數', 'M-04錯題數', 'M-05錯題數', 'M-06錯題數',
      '本次觸發迷思代碼', '本次已克服迷思代碼', '錯題編號與詳情JSON', '完整概念統計JSON'
    ];
    rawSh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#1E293B').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
    rawSh.setFrozenRows(1);
    rawSh.setColumnWidth(1, 155);
    rawSh.setColumnWidth(7, 210);
    rawSh.setColumnWidth(18, 160);
    rawSh.setColumnWidth(19, 160);
    rawSh.setColumnWidth(20, 260);
  }

  // 初始化表 4：個人迷思與成長報告單版面與公式
  buildIndividualReportTemplate(sheets['4_個人迷思與成長報告單']);

  // 執行一次全表儀表板更新
  refreshAllDashboards(ss);

  // 將「2_全班首次闖關迷思診斷」設為開啟預設頁
  ss.setActiveSheet(sheets['2_全班首次闖關迷思診斷']);
}

/**
 * 接收網頁前端 POST 送出的闖關資料
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000); // 防止全班同時交卷衝突
  try {
    const rawText = e.postData ? e.postData.contents : '{}';
    const data = JSON.parse(rawText);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 若尚未執行過 setupSpreadsheet，自動補建工作表
    if (!ss.getSheetByName('1_作答明細')) {
      setupSpreadsheet();
    }

    // 支援教師後台「查詢雲端摘要」請求
    if (data.action === 'ping') {
      return jsonResponse({ ok: true, message: '雲端試算表連線成功！', title: ss.getName() });
    }

    const rawSh = ss.getSheetByName('1_作答明細');
    const studentKey = `${data.className || '未填班級'}_${String(data.seatNo || '0').padStart(2, '0')}_${data.studentName || '匿名'}`;

    // 計算該生在「該關卡」是第幾次挑戰（以雲端真實筆數與前端傳入取較大值確保準確）
    const existingData = rawSh.getDataRange().getValues();
    let cloudAttemptCount = 0;
    let firstAttemptWrongMap = null;
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][1] === studentKey && existingData[i][5] === data.levelId) {
        cloudAttemptCount++;
        if (Number(existingData[i][7]) === 1) {
          firstAttemptWrongMap = {
            'M-01': Number(existingData[i][11]) || 0,
            'M-02': Number(existingData[i][12]) || 0,
            'M-03': Number(existingData[i][13]) || 0,
            'M-04': Number(existingData[i][14]) || 0,
            'M-05': Number(existingData[i][15]) || 0,
            'M-06': Number(existingData[i][16]) || 0
          };
        }
      }
    }
    const attemptNo = cloudAttemptCount + 1;

    const mcWrong = data.misconceptionWrongCounts || {};
    const m01 = Number(mcWrong['M-01'] || 0);
    const m02 = Number(mcWrong['M-02'] || 0);
    const m03 = Number(mcWrong['M-03'] || 0);
    const m04 = Number(mcWrong['M-04'] || 0);
    const m05 = Number(mcWrong['M-05'] || 0);
    const m06 = Number(mcWrong['M-06'] || 0);

    const activeMisconceptions = [];
    ['M-01', 'M-02', 'M-03', 'M-04', 'M-05', 'M-06'].forEach(code => {
      if (Number(mcWrong[code] || 0) > 0) activeMisconceptions.push(code);
    });

    // 比對第 1 次卡關但本次 0 錯題的概念（成功克服迷思！）
    const overcomeList = [];
    if (attemptNo >= 2 && firstAttemptWrongMap) {
      ['M-01', 'M-02', 'M-03', 'M-04', 'M-05', 'M-06'].forEach(code => {
        if (firstAttemptWrongMap[code] > 0 && Number(mcWrong[code] || 0) === 0) {
          overcomeList.push(code);
        }
      });
    }

    const nowStr = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
    const row = [
      nowStr,
      studentKey,
      data.className || '',
      String(data.seatNo || '').padStart(2, '0'),
      data.studentName || '',
      data.levelId || 'level1',
      LEVEL_NAMES[data.levelId] || data.levelName || data.levelId,
      attemptNo,
      Number(data.score || 0),
      Number(data.score || 0) >= 90 ? '✅ 過關' : '❌ 未過關',
      Number(data.timeSeconds || 0),
      m01, m02, m03, m04, m05, m06,
      activeMisconceptions.length > 0 ? activeMisconceptions.join(', ') : '無（全對穩固）',
      overcomeList.length > 0 ? overcomeList.join(', ') : (attemptNo === 1 ? '首次基準測驗' : '-'),
      JSON.stringify(data.wrongDetails || []),
      JSON.stringify(data.conceptStats || {})
    ];

    rawSh.appendRow(row);

    // 即時更新所有統計報表與學生下拉名單
    refreshAllDashboards(ss);

    return jsonResponse({
      ok: true,
      attemptNo: attemptNo,
      overcomeList: overcomeList,
      message: '已成功寫入雲端並更新全班與個人迷思分析表！'
    });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 支援瀏覽器開啟網址測試連線與讀取全班雲端統計
 */
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'status';
  if (action === 'summary') {
    const rawSh = ss.getSheetByName('1_作答明細');
    const rows = rawSh ? rawSh.getDataRange().getValues() : [];
    return jsonResponse({
      ok: true,
      title: ss.getName(),
      totalSubmissions: Math.max(0, rows.length - 1),
      records: rows.slice(1).map(r => ({
        timestamp: r[0],
        studentKey: r[1],
        className: r[2],
        seatNo: r[3],
        studentName: r[4],
        levelId: r[5],
        levelName: r[6],
        attemptNo: r[7],
        score: r[8],
        passed: r[9],
        timeSeconds: r[10],
        mcWrong: { 'M-01': r[11], 'M-02': r[12], 'M-03': r[13], 'M-04': r[14], 'M-05': r[15], 'M-06': r[16] },
        activeMisconceptions: r[17],
        overcome: r[18]
      }))
    });
  }
  return jsonResponse({
    ok: true,
    status: 'READY',
    message: '🏆 三上數學運動會闖關雲端後端運作正常！請將此網址貼入闖關網頁的「教師後台」。'
  });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 自動重新計算並繪製：
 *   - 2_全班首次闖關迷思診斷
 *   - 3_重測進步與迷思消除追蹤
 *   - 5_全班四關進度總表
 *   - 更新 4_個人迷思與成長報告單 的學生選單
 */
function refreshAllDashboards(ss) {
  const rawSh = ss.getSheetByName('1_作答明細');
  if (!rawSh) return;
  const allValues = rawSh.getDataRange().getValues();
  const records = [];
  for (let i = 1; i < allValues.length; i++) {
    const r = allValues[i];
    if (!r[1]) continue;
    records.push({
      time: r[0],
      studentKey: String(r[1]),
      className: String(r[2]),
      seatNo: String(r[3]),
      name: String(r[4]),
      levelId: String(r[5]),
      levelName: String(r[6]),
      attempt: Number(r[7]) || 1,
      score: Number(r[8]) || 0,
      timeSec: Number(r[10]) || 0,
      wrong: {
        'M-01': Number(r[11]) || 0,
        'M-02': Number(r[12]) || 0,
        'M-03': Number(r[13]) || 0,
        'M-04': Number(r[14]) || 0,
        'M-05': Number(r[15]) || 0,
        'M-06': Number(r[16]) || 0
      }
    });
  }

  // 按學生+關卡分組
  const studentMap = {}; // studentKey -> { className, seatNo, name, levels: { levelId: [rec1, rec2...] } }
  records.forEach(rec => {
    if (!studentMap[rec.studentKey]) {
      studentMap[rec.studentKey] = {
        studentKey: rec.studentKey,
        className: rec.className,
        seatNo: rec.seatNo,
        name: rec.name,
        levels: {}
      };
    }
    if (!studentMap[rec.studentKey].levels[rec.levelId]) {
      studentMap[rec.studentKey].levels[rec.levelId] = [];
    }
    studentMap[rec.studentKey].levels[rec.levelId].push(rec);
  });

  // 排序每位學生的歷次挑戰
  Object.values(studentMap).forEach(st => {
    Object.keys(st.levels).forEach(lv => {
      st.levels[lv].sort((a, b) => a.attempt - b.attempt);
    });
  });

  const sortedStudents = Object.values(studentMap).sort((a, b) => {
    if (a.className !== b.className) return a.className.localeCompare(b.className, 'zh-TW');
    return (parseInt(a.seatNo, 10) || 0) - (parseInt(b.seatNo, 10) || 0);
  });

  // =========================================================================
  // 更新表 2：2_全班首次闖關迷思診斷（專門分析第 1 次作答原始基準）
  // =========================================================================
  const sh2 = ss.getSheetByName('2_全班首次闖關迷思診斷');
  sh2.clear();
  sh2.getRange('A1:H1').merge().setValue('📊 全班「第一次闖關」原始成績與六大迷思概念診斷總覽（未看詳解前真實學情）')
    .setBackground('#1E3A8A').setFontColor('#FFFFFF').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');

  // 區塊 A：各關第一次闖關成績統計
  sh2.getRange('A3:F3').setValues([['關卡名稱', '已測學生人數', '首次平均分數', '首次即過關人數(>=90分)', '首次過關率', '首次最高 / 最低分']])
    .setBackground('#DBEAFE').setFontWeight('bold').setHorizontalAlignment('center');

  const levelKeys = ['level1', 'level2', 'level3', 'level4'];
  const lvSummaryRows = [];
  const firstAttemptsAll = [];

  levelKeys.forEach(lv => {
    const firstList = [];
    sortedStudents.forEach(st => {
      if (st.levels[lv] && st.levels[lv].length > 0) {
        firstList.push(st.levels[lv][0]);
        firstAttemptsAll.push(st.levels[lv][0]);
      }
    });
    const cnt = firstList.length;
    if (cnt === 0) {
      lvSummaryRows.push([LEVEL_NAMES[lv], 0, '-', 0, '0%', '-']);
    } else {
      const avg = (firstList.reduce((s, x) => s + x.score, 0) / cnt).toFixed(1);
      const passCnt = firstList.filter(x => x.score >= 90).length;
      const passRate = Math.round((passCnt / cnt) * 100) + '%';
      const maxS = Math.max(...firstList.map(x => x.score));
      const minS = Math.min(...firstList.map(x => x.score));
      lvSummaryRows.push([LEVEL_NAMES[lv], cnt, Number(avg), passCnt, passRate, `${maxS} 分 / ${minS} 分`]);
    }
  });
  sh2.getRange(4, 1, 4, 6).setValues(lvSummaryRows).setHorizontalAlignment('center');

  // 區塊 B：全班第一次闖關「六大迷思概念」踩雷人數排行與教學建議
  sh2.getRange('A9:H9').merge().setValue('🔍 全班「第一次闖關」六大迷思概念踩雷統計（統計每位學生首次遇到該概念即答錯之人數）')
    .setBackground('#7C2D12').setFontColor('#FFFFFF').setFontSize(12).setFontWeight('bold');

  sh2.getRange('A10:H10').setValues([[
    '迷思代碼', '迷思概念名稱', '對應單元', '首次踩雷學生數', '佔已測學生比例', '首次踩雷視覺長條圖', '首次踩雷座號名單', '系統自動教學補救建議'
  ]]).setBackground('#FFEDD5').setFontWeight('bold').setHorizontalAlignment('center');

  const totalUniqueStudents = sortedStudents.length;
  const mcCodes = ['M-01', 'M-02', 'M-03', 'M-04', 'M-05', 'M-06'];

  mcCodes.forEach((code, idx) => {
    const rowIdx = 11 + idx;
    const hitStudents = [];
    sortedStudents.forEach(st => {
      // 只要該生在任一關的「第 1 次挑戰」中，此迷思答錯題數 > 0，即列入首次具有此迷思
      let hasFirstError = false;
      levelKeys.forEach(lv => {
        if (st.levels[lv] && st.levels[lv][0] && st.levels[lv][0].wrong[code] > 0) {
          hasFirstError = true;
        }
      });
      if (hasFirstError) {
        hitStudents.push(`${st.seatNo}號${st.name}`);
      }
    });
    const hitCnt = hitStudents.length;
    const ratio = totalUniqueStudents > 0 ? Math.round((hitCnt / totalUniqueStudents) * 100) : 0;
    sh2.getRange(rowIdx, 1, 1, 8).setValues([[
      code,
      MISCONCEPTION_META[code].name,
      MISCONCEPTION_META[code].unit,
      hitCnt,
      `${ratio}%`,
      '',
      hitStudents.length > 0 ? hitStudents.join('、') : '🎉 全班無人踩雷',
      MISCONCEPTION_META[code].advice
    ]]);
    const maxVal = Math.max(1, totalUniqueStudents);
    sh2.getRange(rowIdx, 6).setFormula(`=SPARKLINE(D${rowIdx},{"charttype","bar";"max",${maxVal};"color1","#EF4444"})`);
  });

  sh2.setColumnWidth(1, 90);
  sh2.setColumnWidth(2, 230);
  sh2.setColumnWidth(3, 165);
  sh2.setColumnWidth(4, 120);
  sh2.setColumnWidth(5, 120);
  sh2.setColumnWidth(6, 160);
  sh2.setColumnWidth(7, 240);
  sh2.setColumnWidth(8, 340);

  // =========================================================================
  // 更新表 3：3_重測進步與迷思消除追蹤（比較第 1 次 vs 第 2、3+ 次是否學會！）
  // =========================================================================
  const sh3 = ss.getSheetByName('3_重測進步與迷思消除追蹤');
  sh3.clear();
  sh3.getRange('A1:H1').merge().setValue('🚀 全班「重測進步與迷思概念消除成效」追蹤表（比較第 1 次 vs 第 2、3 次重測是否真正學會）')
    .setBackground('#065F46').setFontColor('#FFFFFF').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');

  sh3.getRange('A3:H3').setValues([[
    '迷思代碼', '迷思概念名稱', '第1次卡關人次', '重測後已全對學會人次', '重測後仍卡關人次', '尚未重測人次', '迷思修復學會率', '目前仍需老師個別指導名單（重測仍錯）'
  ]]).setBackground('#D1FAE5').setFontWeight('bold').setHorizontalAlignment('center');

  mcCodes.forEach((code, idx) => {
    const rowIdx = 4 + idx;
    let firstFailCount = 0;
    let masteredAfterRetry = 0;
    let stillFailAfterRetry = 0;
    let notRetriedYet = 0;
    const needHelpNames = [];

    sortedStudents.forEach(st => {
      levelKeys.forEach(lv => {
        const attempts = st.levels[lv];
        if (!attempts || attempts.length === 0) return;
        const firstRec = attempts[0];
        const latestRec = attempts[attempts.length - 1];
        if (firstRec.wrong[code] > 0) {
          firstFailCount++;
          if (attempts.length === 1) {
            notRetriedYet++;
          } else {
            if (latestRec.wrong[code] === 0) {
              masteredAfterRetry++;
            } else {
              stillFailAfterRetry++;
              needHelpNames.push(`${st.seatNo}${st.name}(${LEVEL_NAMES[lv].slice(0, 4)})`);
            }
          }
        }
      });
    });

    const retriedPool = masteredAfterRetry + stillFailAfterRetry;
    const masteryRateStr = retriedPool > 0
      ? `${Math.round((masteredAfterRetry / retriedPool) * 100)}% 已學會`
      : (firstFailCount === 0 ? '100% 穩固' : '等待學生重測');

    sh3.getRange(rowIdx, 1, 1, 8).setValues([[
      code,
      MISCONCEPTION_META[code].name,
      firstFailCount,
      masteredAfterRetry,
      stillFailAfterRetry,
      notRetriedYet,
      masteryRateStr,
      needHelpNames.length > 0 ? needHelpNames.join('、') : '✅ 無持續卡關學生'
    ]]);
  });

  // 下方列出每位有進行第 2、3 次重測學生的前後對照表
  sh3.getRange('A12:I12').merge().setValue('📈 學生個別「第 1 次 ➔ 第 2 次 ➔ 第 3+ 次」重測進步與迷思突破明細')
    .setBackground('#1E293B').setFontColor('#FFFFFF').setFontWeight('bold');
  sh3.getRange('A13:I13').setValues([[
    '班級', '座號', '姓名', '關卡', '第1次分數', '第2次分數', '第3次(或最新)分數', '進步總分', '迷思概念修復診斷結果'
  ]]).setBackground('#E2E8F0').setFontWeight('bold').setHorizontalAlignment('center');

  const growthRows = [];
  sortedStudents.forEach(st => {
    levelKeys.forEach(lv => {
      const att = st.levels[lv];
      if (!att || att.length === 0) return;
      const s1 = att[0].score;
      const s2 = att.length >= 2 ? att[1].score : '-';
      const s3 = att.length >= 3 ? att[att.length - 1].score : '-';
      const latest = att[att.length - 1];
      const diff = latest.score - s1;

      const fixed = [];
      const stuck = [];
      mcCodes.forEach(c => {
        if (att[0].wrong[c] > 0) {
          if (att.length >= 2 && latest.wrong[c] === 0) fixed.push(c);
          else if (att.length >= 2 && latest.wrong[c] > 0) stuck.push(c);
        }
      });

      let statusText = '';
      if (att.length === 1) {
        statusText = s1 >= 90 ? '🌟 首次即高分過關' : '⏳ 僅測1次（尚未重測）';
      } else {
        const parts = [];
        if (fixed.length > 0) parts.push(`🎉 已學會修復：${fixed.join(',')}`);
        if (stuck.length > 0) parts.push(`⚠️ 仍待加強：${stuck.join(',')}`);
        if (parts.length === 0) parts.push('✅ 觀念維持穩固');
        statusText = parts.join(' ｜ ');
      }

      growthRows.push([
        st.className, st.seatNo, st.name, LEVEL_NAMES[lv],
        s1, s2, s3,
        att.length >= 2 ? (diff >= 0 ? `+${diff} 分` : `${diff} 分`) : '-',
        statusText
      ]);
    });
  });

  if (growthRows.length > 0) {
    sh3.getRange(14, 1, growthRows.length, 9).setValues(growthRows).setHorizontalAlignment('center');
    sh3.getRange(14, 9, growthRows.length, 1).setHorizontalAlignment('left');
  }
  sh3.setColumnWidth(2, 230);
  sh3.setColumnWidth(4, 210);
  sh3.setColumnWidth(8, 300);
  sh3.setColumnWidth(9, 340);

  // =========================================================================
  // 更新表 5：5_全班四關進度總表
  // =========================================================================
  const sh5 = ss.getSheetByName('5_全班四關進度總表');
  sh5.clear();
  sh5.getRange('A1:L1').merge().setValue('🏅 全班四關總複習闖關進度與成績總表（格式：首次分 ➔ 最高分 / 次數）')
    .setBackground('#7E22CE').setFontColor('#FFFFFF').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');

  sh5.getRange('A3:L3').setValues([[
    '班級', '座號', '姓名', '目前已解鎖最高關卡',
    '初級：首次分', '初級：最高分(次數)',
    '中級：首次分', '中級：最高分(次數)',
    '高級：首次分', '高級：最高分(次數)',
    '超高級：首次分', '超高級：最高分(次數)'
  ]]).setBackground('#F3E8FF').setFontWeight('bold').setHorizontalAlignment('center');

  const progressRows = sortedStudents.map(st => {
    const lvCells = [];
    let highestBadge = '尚未過關';
    levelKeys.forEach((lv, idx) => {
      const att = st.levels[lv];
      if (!att || att.length === 0) {
        lvCells.push('-', '-');
      } else {
        const firstS = att[0].score;
        const bestS = Math.max(...att.map(x => x.score));
        if (bestS >= 90) {
          highestBadge = ['🥉 初級過關', '🥈 中級過關', '🥇 高級過關', '🏆 四關全制霸總冠軍！'][idx];
        }
        lvCells.push(firstS, `${bestS}分 (${att.length}次)`);
      }
    });
    return [st.className, st.seatNo, st.name, highestBadge, ...lvCells];
  });

  if (progressRows.length > 0) {
    sh5.getRange(4, 1, progressRows.length, 12).setValues(progressRows).setHorizontalAlignment('center');
  }
  sh5.setColumnWidth(4, 190);

  // =========================================================================
  // 更新表 4 的學生下拉選單驗證名單
  // =========================================================================
  const sh4 = ss.getSheetByName('4_個人迷思與成長報告單');
  if (sh4) {
    const studentKeys = sortedStudents.map(s => s.studentKey);
    if (studentKeys.length > 0) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(studentKeys, true).setAllowInvalid(true).build();
      sh4.getRange('C3').setDataValidation(rule);
      if (!sh4.getRange('C3').getValue()) {
        sh4.getRange('C3').setValue(studentKeys[0]);
      }
    }
  }
}

/**
 * 建立表 4：「個人迷思與成長報告單」動態公式版面（老師切換下拉選單即自動跳出該生第1/2/3次成長分析！）
 */
function buildIndividualReportTemplate(sh) {
  sh.clear();
  sh.getRange('A1:G1').merge().setValue('📋 學生個人「歷次闖關成績 ＆ 迷思概念修復成長」診斷報告單')
    .setBackground('#0F172A').setFontColor('#FFFFFF').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');

  sh.getRange('A3').setValue('👤 選擇學生：').setFontWeight('bold').setHorizontalAlignment('right').setBackground('#FEF3C7');
  sh.getRange('B3').setValue('（請點右側下拉選單）').setFontColor('#64748B').setBackground('#FEF3C7');
  sh.getRange('C3').setBackground('#FDE68A').setFontWeight('bold').setHorizontalAlignment('center');

  sh.getRange('E3').setValue('🎯 選擇關卡：').setFontWeight('bold').setHorizontalAlignment('right').setBackground('#FEF3C7');
  const lvRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['level1', 'level2', 'level3', 'level4'], true).build();
  sh.getRange('F3').setDataValidation(lvRule).setValue('level1').setBackground('#FDE68A').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('G3').setFormula('=IF(F3="level1","初級：選手熱身場",IF(F3="level2","中級：校園選拔賽",IF(F3="level3","高級：分組錦標賽","超高級：運動會素養決賽")))')
    .setFontWeight('bold').setFontColor('#1D4ED8');

  // 成績成長比較表頭
  sh.getRange('A5:G5').setValues([[
    '評量指標 / 迷思概念', '對應單元', '第 1 次挑戰（基準）', '第 2 次挑戰', '第 3 次挑戰', '總挑戰次數', '迷思修復與學習成長最終判定'
  ]]).setBackground('#1E293B').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

  // 第 6 列：總得分演變（使用 FILTER/INDEX 抓取 1_作答明細）
  sh.getRange('A6').setValue('🏆 闖關總得分（滿分100）').setFontWeight('bold').setBackground('#F1F5F9');
  sh.getRange('B6').setValue('第 1～4 單元綜合').setHorizontalAlignment('center').setBackground('#F1F5F9');
  sh.getRange('C6').setFormula('=IFERROR(INDEX(FILTER(\'1_作答明細\'!I:I, \'1_作答明細\'!B:B=$C$3, \'1_作答明細\'!F:F=$F$3, \'1_作答明細\'!H:H=1),1)&" 分", "尚未作答")').setHorizontalAlignment('center').setFontWeight('bold');
  sh.getRange('D6').setFormula('=IFERROR(INDEX(FILTER(\'1_作答明細\'!I:I, \'1_作答明細\'!B:B=$C$3, \'1_作答明細\'!F:F=$F$3, \'1_作答明細\'!H:H=2),1)&" 分", "-")').setHorizontalAlignment('center').setFontWeight('bold');
  sh.getRange('E6').setFormula('=IFERROR(INDEX(FILTER(\'1_作答明細\'!I:I, \'1_作答明細\'!B:B=$C$3, \'1_作答明細\'!F:F=$F$3, \'1_作答明細\'!H:H=3),1)&" 分", "-")').setHorizontalAlignment('center').setFontWeight('bold');
  sh.getRange('F6').setFormula('=COUNTIFS(\'1_作答明細\'!B:B, $C$3, \'1_作答明細\'!F:F, $F$3)&" 次"').setHorizontalAlignment('center');
  sh.getRange('G6').setFormula('=IF(COUNTIFS(\'1_作答明細\'!B:B,$C$3,\'1_作答明細\'!F:F,$F$3)=0,"尚未測驗",IF(MAXIFS(\'1_作答明細\'!I:I,\'1_作答明細\'!B:B,$C$3,\'1_作答明細\'!F:F,$F$3)>=90,"🎉 已達 90 分成功晉級！","💪 最高 "&MAXIFS(\'1_作答明細\'!I:I,\'1_作答明細\'!B:B,$C$3,\'1_作答明細\'!F:F,$F$3)&" 分，請繼續挑戰"))').setFontWeight('bold');

  // 第 7～12 列：M-01 ～ M-06 各項迷思第 1、2、3 次錯題演變與是否學會判定
  const colLetters = ['L', 'M', 'N', 'O', 'P', 'Q']; // 對應 1_作答明細 中 M-01 ~ M-06 的欄位
  const codes = ['M-01', 'M-02', 'M-03', 'M-04', 'M-05', 'M-06'];

  codes.forEach((code, idx) => {
    const r = 7 + idx;
    const col = colLetters[idx];
    sh.getRange(r, 1).setValue(`${code} ${MISCONCEPTION_META[code].name}`).setFontWeight('bold');
    sh.getRange(r, 2).setValue(MISCONCEPTION_META[code].unit).setHorizontalAlignment('center');

    [1, 2, 3].forEach((attNo, cIdx) => {
      const targetCol = 3 + cIdx;
      const formula = `=IFERROR(IF(INDEX(FILTER('1_作答明細'!${col}:${col}, '1_作答明細'!B:B=$C$3, '1_作答明細'!F:F=$F$3, '1_作答明細'!H:H=${attNo}),1)=0, "🟢 全對穩固", "🔴 答錯 "&INDEX(FILTER('1_作答明細'!${col}:${col}, '1_作答明細'!B:B=$C$3, '1_作答明細'!F:F=$F$3, '1_作答明細'!H:H=${attNo}),1)&" 題"), "-")`;
      sh.getRange(r, targetCol).setFormula(formula).setHorizontalAlignment('center');
    });

    sh.getRange(r, 6).setFormula(`=IF(C${r}="-","-",SUMIFS('1_作答明細'!${col}:${col},'1_作答明細'!B:B,$C$3,'1_作答明細'!F:F,$F$3)&" 題累計錯")`).setHorizontalAlignment('center');

    // 第 7 欄：比較第 1 次與最新一次，自動判定是否已學會該迷思概念！
    const judgeFormula = `=IF(C${r}="-", "尚未作答", IF(C${r}="🟢 全對穩固", "🌟 首次即完全掌握", IF(COUNTIFS('1_作答明細'!B:B,$C$3,'1_作答明細'!F:F,$F$3)=1, "🔍 首次發現迷思（待重測驗證）", IF(INDEX(FILTER('1_作答明細'!${col}:${col}, '1_作答明細'!B:B=$C$3, '1_作答明細'!F:F=$F$3), COUNTIFS('1_作答明細'!B:B,$C$3,'1_作答明細'!F:F,$F$3))=0, "🎉 觀念突破！重測已學會此概念", "⚠️ 重測仍答錯，建議老師個別指導"))))`;
    sh.getRange(r, 7).setFormula(judgeFormula).setFontWeight('bold');
  });

  // 教學建議說明框
  sh.getRange('A14:G14').merge().setValue('💡 老師與家長輔導參考指南（六大迷思對照說明）')
    .setBackground('#DBEAFE').setFontWeight('bold');
  codes.forEach((code, idx) => {
    sh.getRange(15 + idx, 1, 1, 7).merge()
      .setValue(`• ${code}【${MISCONCEPTION_META[code].name}】：${MISCONCEPTION_META[code].advice}`);
  });

  sh.setColumnWidth(1, 260);
  sh.setColumnWidth(2, 165);
  sh.setColumnWidth(3, 145);
  sh.setColumnWidth(4, 145);
  sh.setColumnWidth(5, 145);
  sh.setColumnWidth(6, 125);
  sh.setColumnWidth(7, 290);
}
