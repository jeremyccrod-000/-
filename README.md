// ==================== 【設定區】 ====================
// 1. 11508 範本：活動介入 初次評估單 (Google Doc ID)
const TEMPLATE_DOC_ID = '1DUiwhItDARFAGTPjuscM6s9P-b48QAy9gHbq0v4a5ic';

// 2. 產出的評估單想要存入的【指定資料夾 ID】
const SPECIFIED_FOLDER_ID = '1hzPos2XAbKlKcrBnszyezzROLeZwEbs4';
// ===================================================

// 一、診斷對應表 (d1 ~ d13)
const DIAGNOSIS_MAP = {
  '無': 'd1', '痛風': 'd2', '高血壓': 'd3', '糖尿病': 'd4', '心血管疾病': 'd5',
  '呼吸系統疾病': 'd6', '失智症': 'd7', '腎臟疾病': 'd8', '肝膽腸胃疾病': 'd9',
  '精神科疾病': 'd10', '帕金森氏症': 'd11', '腦血管疾病': 'd12', '癌症': 'd13'
};

// 二、重視問題對應表 (q1 ~ q20)
const PROBLEM_MAP = {
  '家屬衛教擺位及關節運動': 'q1', '疼痛': 'q2', '手部功能不佳': 'q3',
  '自行轉位(翻身、上下床)': 'q4', '溝通功能': 'q5', '寫字': 'q6',
  '站姿/坐姿不平衡': 'q7', '室內行走': 'q8', '室外行走': 'q9', '上下樓梯': 'q10',
  '進食': 'q11', '盥洗': 'q12', '穿脫衣服': 'q13', '如廁': 'q14',
  '大小便控制': 'q15', '洗澡': 'q16', '休閒娛樂': 'q17',
  '提升從事活動的動機': 'q18', '養成良好日常作息': 'q19', '情緒及心理支持及穩定': 'q20'
};

// 三、生理功能單選對應表 (s1 ~ s28)
const SINGLE_CHOICE_MAP = {
  '清醒': 's1', '嗜睡': 's2', '混亂': 's3', '昏迷': 's4',
  '口語': 's5', '手勢或紙筆': 's6', '無': 's7',
  '複雜或抽象語句': 's8', '簡單句子': 's9', '單字': 's10', '無法理解': 's11',
  '佳': 's12', '普通': 's13', '被動': 's14', '差': 's15',
  '視覺_正常': 's16', '視覺_受損': 's17', '視覺_N/A': 's18',
  '聽覺_正常': 's19', '聽覺_受損': 's20', '聽覺_N/A': 's21',
  '步行': 's22', '輔助器': 's23', '輪椅': 's24', '臥床': 's25',
  '自己吃': 's26', '餵食': 's27', '管灌': 's28'
};

// 四、11508 標準輔具對應表 (簡化回大項名稱)
const AIDS_EXACT_MAP = {
  '移位板': 1, '止滑墊': 2, '助行器': 3, '電話輔助器': 4,
  '床邊扶手': 5, '特殊餐具(湯匙/筷子/碗盤/杯子)': 6, '四腳拐': 7, '語音溝通板': 8,
  '特殊擺位輔具': 9, '撥桿式/感光式水龍頭': 10, '一般輪椅': 11, '特殊滑鼠/鍵盤': 12,
  '副木': 13, '穿衣/褲輔具': 14, '特製輪椅': 15, '握筆輔具': 16,
  '便盆椅': 17, '繫鈕釦輔具': 18,
  '洗澡座椅': 20, '長柄海綿刷': 21
};

// 五、主要問題文字清單
const MAIN_PROBLEMS = [
  "關節活動受限", "疼痛", "意識/警醒度不佳", "姿勢控制不佳",
  "肌肉張力異常", "肌力不足", "耐力不足",
  "翻身/躺到坐/轉位困難", "坐姿/站姿平衡困難", "精細動作不佳", "協調不佳",
  "步行/上下樓梯困難", "認知功能不佳", "自我照顧能力不佳"
];

// 六、短期目標文字清單
const SHORT_GOALS = [
  "改善關節活動度", "改善疼痛", "改善警醒度", "改善姿勢控制",
  "正常化張力", "加強肌力", "加強心肺耐力",
  "強化翻身/轉位/躺到坐", "增進坐姿/站姿平衡", "改善精細動作", "改善協調功能",
  "強化功能性移行", "促進認知功能", "促進ADL功能"
];

// 六、長期目標文字清單
const LONG_GOALS = [
  "維持現有功能", "部分需他人照顧", "完全獨立",
  "減少臥床時間", "自行使用輪椅", "協助/使用輔具下步行", "自行步行"
];

// 七、11508 介入項目標準清單 (去除團/個前綴)
const INTERVENTION_ITEMS = [
  "關節運動指導及正確姿勢擺位衛教指導",
  "至少1-2小時改變姿勢，以避免長時間維持相同姿勢造成骨凸處易受壓迫",
  "深壓覺、本體覺、聽覺等感官刺激輸入",
  "輔具評估建議與指導",
  "照顧者照顧技巧衛教與指導",
  "上、下肢肌力訓練",
  "心肺耐力訓練",
  "功能性活動介入(翻身/躺到坐/轉位)",
  "坐、站姿平衡能力訓練",
  "精細動作，手功能訓練",
  "協調能力訓練",
  "推輪椅/行走能力訓練",
  "認知功能訓練",
  "ADL訓練"
];

// ==================== 【工具函式】 ====================

/**
 * 清洗舊表單項目文字：自動拿掉「團：」或「個：」前綴，並去除多餘空白
 */
function cleanItemText(text) {
  if (!text) return '';
  return text.toString().replace(/^[團個]：/, '').trim();
}

/**
 * 逸出正則表達式特殊字元
 */
function escapeRegex(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * 核取方塊替換核心函式
 */
function replaceCheckmark(body, itemText, isChecked) {
  if (!itemText) return;
  const mark = isChecked ? '■' : '□';
  // 匹配 ☐、□ 或 ■ 後接該文字，並替換為目標狀態
  body.replaceText(`[☐□■]\\s*${escapeRegex(itemText)}`, `${mark} ${itemText}`);
}

/**
 * 解析目標儲存資料夾
 */
function resolveDestinationFolder(folderId, fallbackFile) {
  try {
    if (folderId && folderId !== '請貼上新帳號中資料夾的ID') {
      return DriveApp.getFolderById(folderId);
    }
  } catch (err) {
    Logger.log('⚠️ 資料夾 ID 無效，改存至範本同層目錄：' + err.message);
  }
  const parents = fallbackFile.getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}

// ==================== 【表單提交主執行函式】 ====================

function onFormSubmitAutoFillNew(e) {
  if (!e || !e.response) {
    Logger.log('⚠️ 本函式需由 Google 表單提交事件觸發。');
    return;
  }

  const itemResponses = e.response.getItemResponses();
  const templateFile = DriveApp.getFileById(TEMPLATE_DOC_ID);
  const targetFolder = resolveDestinationFolder(SPECIFIED_FOLDER_ID, templateFile);

  // 1. 建立個案資訊字典
  const data = {};
  itemResponses.forEach(res => {
    const title = res.getItem().getTitle().trim();
    const val = res.getResponse();
    data[title] = val;
  });

  const caseName = data['姓名'] || '未知姓名';
  const caseNo = data['案號'] || '';
  const evalDate = data['評估日期'] || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');

  // 2. 複製 11508 範本文件並命名
  const newFileName = `${evalDate}_活動介入初次評估單_${caseName}`;
  const newDocFile = templateFile.makeCopy(newFileName, targetFolder);
  const newDoc = DocumentApp.openById(newDocFile.getId());
  const body = newDoc.getBody();

  // 3. 基本文字欄位填入
  body.replaceText('案號：\\s*', `案號：${caseNo}    `);
  body.replaceText('姓名：\\s*', `姓名：${caseName}    `);
  body.replaceText('年齡：\\s*', `年齡：${data['年齡'] || ''}    `);
  body.replaceText('評估日期:\\s*', `評估日期: ${evalDate}    `);
  body.replaceText('受託日：\\s*', `受託日：${data['受託日'] || ''}    `);
  body.replaceText('注意事項:\\s*', `注意事項: ${data['注意事項'] || ''}    `);
  body.replaceText('指導者簽章:\\s*', `指導者簽章: ${data['指導者簽章'] || ''}    `);
  body.replaceText('評估者簽章:\\s*', `評估者簽章: ${data['評估者簽章'] || ''}    `);

  // 性別勾選
  const gender = data['性別'] || '';
  replaceCheckmark(body, '男', gender === '男');
  replaceCheckmark(body, '女', gender === '女');

  // 4. 一、診斷勾選
  const userDiagnoses = [].concat(data['診斷'] || []);
  Object.keys(DIAGNOSIS_MAP).forEach(diag => {
    replaceCheckmark(body, diag, userDiagnoses.includes(diag));
  });

  // 5. 二、重視問題勾選
  const userProblems = [].concat(data['個案或家屬最重視的問題'] || data['最重視的問題（可複選）'] || []);
  Object.keys(PROBLEM_MAP).forEach(prob => {
    replaceCheckmark(body, prob, userProblems.includes(prob));
  });

  // 6. 三、生理功能勾選 (單選)
  ['意識狀態', '表達方式', '理解能力', '個案動機和配合度', '視覺', '聽覺', '移動方式', '進食方式'].forEach(item => {
    const val = data[item];
    if (val) replaceCheckmark(body, val, true);
  });

  // 認知功能分數填入
  if (data['認知功能及需協助程度 (評估分數)']) {
    const cogScore = data['認知功能及需協助程度 (評估分數)'].toString().replace(/[^0-9]/g, '');
    body.replaceText('評估分數:\\s*', `評估分數: ${cogScore}`);
  }

  // 7. 四、輔具需求 (包含文字清洗與還原大項)
  const currentAids = [].concat(data['目前使用輔具'] || []);
  currentAids.forEach(aid => {
    const cleanAid = aid.split('(')[0].trim(); // 去掉括號子項目
    replaceCheckmark(body, cleanAid, true);
  });

  // 8. 五、主要問題、六、目標勾選
  const mainProblems = [].concat(data['主要問題'] || []);
  MAIN_PROBLEMS.forEach(prob => replaceCheckmark(body, prob, mainProblems.includes(prob)));

  const shortGoals = [].concat(data['短期目標'] || []);
  SHORT_GOALS.forEach(goal => replaceCheckmark(body, goal, shortGoals.includes(goal)));

  const longGoals = [].concat(data['長期目標'] || []);
  LONG_GOALS.forEach(goal => replaceCheckmark(body, goal, longGoals.includes(goal)));

  // 9. 七、介入項目 (核心：自動清洗「團：/個：」並白名單比對)
  const userInterventions = [].concat(data['介入項目'] || data['介入項目（可複選）'] || []);
  const cleanedUserInterventions = userInterventions.map(item => cleanItemText(item));

  INTERVENTION_ITEMS.forEach(item => {
    // 只要表單提交的內容清洗後有符合 11508 項目就勾選，不在 11508 裡面的多餘欄位則完全不處理
    const isChecked = cleanedUserInterventions.includes(item);
    replaceCheckmark(body, item, isChecked);
  });

  // 儲存並關閉新檔案
  newDoc.saveAndClose();
  Logger.log(`✅ 成功產出 11508 評估單：${newFileName}`);
}# -
