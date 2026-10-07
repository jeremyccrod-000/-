// ==================== 【設定區與對照表】 ====================
const TEMPLATE_DOC_ID = '1qL66kN63cbxsJrAaLp53S-D6GQP8H3MmCJRG5UVQUvY';
const SPECIFIED_FOLDER_ID = '1hzPos2XAbKlKcrBnszyezzROLeZwEbs4';

const DIAGNOSIS_MAP = {
  '無': 'd1', '痛風': 'd2', '高血壓': 'd3', '糖尿病': 'd4', '心血管疾病': 'd5',
  '呼吸系統疾病': 'd6', '失智症': 'd7', '腎臟疾病': 'd8', '肝膽腸胃疾病': 'd9',
  '精神科疾病': 'd10', '帕金森氏症': 'd11', '腦血管疾病': 'd12', '癌症': 'd13'
};

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

const AIDS_KEY_MAP = {
  '移位板': 1, '止滑墊': 2, '助行器': 3, '電話輔助器': 4, '床邊扶手': 5,
  '特殊餐具': 6, '四腳拐': 7, '語音溝通板': 8, '特殊擺位輔具': 9, '水龍頭': 10,
  '一般輪椅': 11, '特殊': 12, '滑鼠': 12, '鍵盤': 12, '副木': 13,
  '穿衣': 14, '穿褲': 14, '特製輪椅': 15, '握筆輔具': 16, '便盆椅': 17,
  '繫鈕釦輔具': 18, '洗澡座椅': 20, '長柄海綿刷': 21
};

const PROBLEM_MAP = {
  '關節活動受限': 'p1', '疼痛': 'p2', '意識/警醒度不佳': 'p3', '姿勢控制不佳': 'p4',
  '肌肉張力異常': 'p5', '肌力不足': 'p6', '耐力不足': 'p7', '翻身困難': 'p8',
  '躺到坐困難': 'p8', '轉位困難': 'p8', '坐姿平衡困難': 'p9', '站姿平衡困難': 'p9',
  '精細動作不佳': 'p10', '協調不佳': 'p11', '步行困難': 'p12', '上下樓梯困難': 'p12',
  '認知功能不佳': 'p13', '自我照顧能力不佳': 'p14', '其他': 'p15'
};

const SHORT_TERM_MAP = {
  '改善關節活動度': 'st1', '改善疼痛': 'st2', '改善警醒度': 'st3', '改善姿勢控制': 'st4',
  '正常化張力': 'st5', '加強肌力': 'st6', '加強心肺耐力': 'st7', '強化翻身': 'st8',
  '強化轉位': 'st8', '強化躺到坐': 'st8', '增進坐姿平衡': 'st9', '增進站姿平衡': 'st9',
  '改善精細動作': 'st10', '改善協調功能': 'st11', '強化功能性移行': 'st12',
  '促進認知功能': 'st13', '促進ADL功能': 'st14', '其他': 'st15'
};

const LONG_TERM_MAP = {
  '維持現有功能': 'lt1', '部分需他人照顧': 'lt2', '完全獨立': 'lt3', '減少臥床時間': 'lt4',
  '自行使用輪椅': 'lt5', '協助輔具下步行': 'lt6', '使用輔具下步行': 'lt6', '自行步行': 'lt7', '其他': 'lt8'
};

// ==================== 【輔助函式】 ====================
function cleanText(text) {
  return (text === null || text === undefined) ? '' : String(text).replace(/^[團個][：:]/, '').trim();
}

function escapeRegex(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function replaceTag(body, tag, val) {
  body.replaceText('\\{\\{' + escapeRegex(tag) + '\\}\\}', (val !== undefined && val !== null) ? String(val) : '');
}

function markBoxInElement(el, word) {
  if (!el || !word) return;
  el.replaceText('[□☐]\\s*' + escapeRegex(word), '■ ' + word);
}

// ==================== 【單格單肢體精準替換：徹底防止多選】 ====================
function updateGridLimbText(cell, limbName, rawVal, type) {
  if (!cell || !rawVal) return;
  const val = String(rawVal);
  
  // 取得儲存格內完整文字並分行處理
  let text = cell.getText();
  const lines = text.split('\n');
  
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].includes(limbName)) {
      let line = lines[i];

      if (type === 'ROM') {
        if (val.includes('正常')) {
          line = line.replace(/[□☐]\s*正常/, '■ 正常');
        } else if (val.includes('受限')) {
          let circleGrading = '受限(輕/中/重)';
          if (val.includes('輕')) circleGrading = '受限(○輕/中/重)';
          else if (val.includes('中')) circleGrading = '受限(輕/○中/重)';
          else if (val.includes('重')) circleGrading = '受限(輕/中/○重)';
          line = line.replace(/[□☐]\s*受限\s*\([輕/中/重]+\)/, '■ ' + circleGrading);
        }
      } else {
        // 張力 (正常/高張/低張) 或 肌力 (正常/尚可/不足)
        const opts = (type === 'TONE') ? ['正常', '高張', '低張'] : ['正常', '尚可', '不足'];
        opts.forEach(opt => {
          if (val.includes(opt)) {
            // 使用正則嚴格只替換當前選項前面的方框
            const reg = new RegExp('[□☐]\\s*' + escapeRegex(opt));
            line = line.replace(reg, '■ ' + opt);
          }
        });
      }
      lines[i] = line;
      break; // 處理完該肢體行即結束
    }
  }

  // 將修改後的文字回填給儲存格
  cell.setText(lines.join('\n'));
}

// ==================== 【主執行函式】 ====================
function onFormSubmitAutoFillNew(e) {
  if (!e || !e.response) {
    Logger.log('⚠️ 本函式需由 Google 表單提交事件觸發。');
    return;
  }

  const templateFile = DriveApp.getFileById(TEMPLATE_DOC_ID);
  let targetFolder;
  try { targetFolder = DriveApp.getFolderById(SPECIFIED_FOLDER_ID); }
  catch (err) { targetFolder = templateFile.getParents().next(); }

  // 1. 蒐集表單資料
  const formData = {}, allValues = [];
  e.response.getItemResponses().forEach(res => {
    const title = res.getItem().getTitle().trim();
    const resp = res.getResponse();
    formData[title] = resp;
    const addVal = v => { if (v !== null && v !== undefined) allValues.push(cleanText(v)); };
    if (Array.isArray(resp)) {
      resp.forEach(item => {
        if (Array.isArray(item)) item.forEach(addVal);
        else if (typeof item === 'string') item.split(',').forEach(addVal);
        else addVal(item);
      });
    } else if (resp !== null && resp !== undefined) {
      String(resp).split(',').forEach(addVal);
    }
  });

  const getFormDataByKey = k => {
    for (let key in formData) { if (key.includes(k)) return formData[key]; }
    return null;
  };

  // 2. 基本資料與複製範本
  const caseName = cleanText(getFormDataByKey('姓名')) || '長輩';
  const caseNo = cleanText(getFormDataByKey('案號')) || '';
  const evalDate = cleanText(getFormDataByKey('評估日期')) || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  const newFileName = `${evalDate}_活動介入初次評估單_${caseName}`;
  const newDoc = DocumentApp.openById(templateFile.makeCopy(newFileName, targetFolder).getId());
  const body = newDoc.getBody();

  // 填寫評估日期
  replaceTag(body, '評估日期', evalDate);
  body.replaceText('評估日期:\\s*', '評估日期: ' + evalDate + '   ');
  body.replaceText('評估日期：\\s*', '評估日期: ' + evalDate + '   ');

  replaceTag(body, '案號', caseNo);
  replaceTag(body, '姓名', caseName);
  replaceTag(body, '年齡', getFormDataByKey('年齡') || '');
  replaceTag(body, '受託日', getFormDataByKey('受託日') || '');
  replaceTag(body, '注意事項', getFormDataByKey('注意事項') || '');
  replaceTag(body, 'supervisor_name', String(getFormDataByKey('指導者') || '').trim());
  replaceTag(body, 'assessor_name', String(getFormDataByKey('評估者') || '').trim());

  const gender = getFormDataByKey('性別') || '';
  replaceTag(body, '性別_男', gender === '男' ? '■' : '□');
  replaceTag(body, '性別_女', gender === '女' ? '■' : '□');

  // 3. 一、診斷
  const userDiag = [].concat(getFormDataByKey('診斷') || []);
  let hasOtherDiag = false;
  Object.keys(DIAGNOSIS_MAP).forEach(d => replaceTag(body, DIAGNOSIS_MAP[d], userDiag.includes(d) ? '■' : '□'));
  userDiag.forEach(d => { if (!DIAGNOSIS_MAP[d] && d !== '其他') hasOtherDiag = true; });
  replaceTag(body, 'd14', (hasOtherDiag || userDiag.includes('其他')) ? '■' : '□');

  // 4. 二、生理功能評估
  const visualVal = getFormDataByKey('視覺'), auditVal = getFormDataByKey('聽覺');
  Object.keys(SINGLE_CHOICE_MAP).forEach(k => {
    let isChecked = k.startsWith('視覺_') ? (visualVal === k.replace('視覺_', ''))
      : k.startsWith('聽覺_') ? (auditVal === k.replace('聽覺_', ''))
      : allValues.includes(k);
    replaceTag(body, SINGLE_CHOICE_MAP[k], isChecked ? '■' : '□');
  });
  replaceTag(body, 's29', '□');

  // 認知分數
  let scoreVal = '';
  allValues.forEach(v => {
    if (/(Ind\.|Stand-by|Min\.|Mod\.|Max\.|Total\.)/.test(v)) {
      const m = v.match(/\d+/);
      if (m) scoreVal = m[0];
    }
  });
  replaceTag(body, 'score', scoreVal);

  // 5. 三、輔具
  const currentAidsVal = getFormDataByKey('目前使用輔具') || '';
  const currentAidsText = Array.isArray(currentAidsVal) ? currentAidsVal.join('、') : String(currentAidsVal);
  replaceTag(body, '目前使用輔具', currentAidsText);
  replaceTag(body, 'current_aids', currentAidsText);

  for (let i = 1; i <= 21; i++) replaceTag(body, 'a' + i, '□');
  ['a6_1','a6_2','a6_3','a6_4','a10_1','a10_2','a12_1','a12_2','a14_1','a14_2'].forEach(s => replaceTag(body, s, '□'));
  replaceTag(body, 'a_other', '');

  const suggestAids = [].concat(getFormDataByKey('建議增加') || [], getFormDataByKey('輔具') || []);
  suggestAids.forEach(aid => {
    const aidStr = String(aid);
    Object.keys(AIDS_KEY_MAP).forEach(k => {
      if (aidStr.includes(k)) { replaceTag(body, 'a' + AIDS_KEY_MAP[k], '■'); markBoxInElement(body, k); }
    });
    const subItems = { '湯匙':'a6_1', '筷子':'a6_2', '碗盤':'a6_3', '杯子':'a6_4', '撥桿':'a10_1', '感光':'a10_2', '滑鼠':'a12_1', '鍵盤':'a12_2', '穿衣':'a14_1', '穿褲':'a14_2' };
    Object.keys(subItems).forEach(sub => {
      if (aidStr.includes(sub)) { replaceTag(body, subItems[sub], '■'); markBoxInElement(body, sub); }
    });
  });

  // 6. 四、主要問題
  for (let i = 1; i <= 15; i++) replaceTag(body, 'p' + i, '□');
  replaceTag(body, 'p_other', '');
  [].concat(getFormDataByKey('主要問題') || []).forEach(p => {
    const pStr = String(p).trim();
    Object.keys(PROBLEM_MAP).forEach(k => {
      if (pStr.includes(k)) { replaceTag(body, PROBLEM_MAP[k], '■'); markBoxInElement(body, k); }
    });
    if (/(翻身|躺到坐|轉位)/.test(pStr)) replaceTag(body, 'p8', '■');
    if (/(坐姿平衡|站姿平衡)/.test(pStr)) replaceTag(body, 'p9', '■');
    if (/(步行|樓梯)/.test(pStr)) replaceTag(body, 'p12', '■');
  });

  // 7. 五、短期目標
  for (let i = 1; i <= 15; i++) replaceTag(body, 'st' + i, '□');
  replaceTag(body, 'st_other', '');
  [].concat(getFormDataByKey('短期目標') || []).forEach(g => {
    const gStr = String(g).trim();
    Object.keys(SHORT_TERM_MAP).forEach(k => {
      if (gStr.includes(k)) { replaceTag(body, SHORT_TERM_MAP[k], '■'); markBoxInElement(body, k); }
    });
    if (/(翻身|轉位|躺到坐)/.test(gStr)) replaceTag(body, 'st8', '■');
    if (/(坐姿平衡|站姿平衡)/.test(gStr)) replaceTag(body, 'st9', '■');
    if (gStr.includes('精細動作')) replaceTag(body, 'st10', '■');
    if (gStr.includes('ADL')) replaceTag(body, 'st14', '■');
  });

  // 8. 五、長期目標
  for (let i = 1; i <= 8; i++) replaceTag(body, 'lt' + i, '□');
  replaceTag(body, 'lt_other', '');
  [].concat(getFormDataByKey('長期目標') || []).forEach(item => {
    const itemStr = String(item).trim();
    Object.keys(LONG_TERM_MAP).forEach(k => {
      if (itemStr.includes(k)) { replaceTag(body, LONG_TERM_MAP[k], '■'); markBoxInElement(body, k); }
    });
  });

  // 9. 六、介入項目
  let indList = [], grpList = [];
  [].concat(getFormDataByKey('介入項目') || []).forEach(item => {
    const s = String(item).trim();
    if (/^個[：:]/.test(s)) indList.push('• ' + s.replace(/^個[：:]/, '').trim());
    else if (/^團[：:]/.test(s)) grpList.push('• ' + s.replace(/^團[：:]/, '').trim());
    else if (s) indList.push('• ' + s);
  });
  let outputParts = [];
  if (indList.length) outputParts.push('【個別介入】\n' + indList.join('\n'));
  if (grpList.length) outputParts.push('【團體介入】\n' + grpList.join('\n'));
  replaceTag(body, 'intervention_items', outputParts.length ? outputParts.join('\n\n') : '無');

  // 10. 生理功能網格表格（嚴格逐行隔離：左上/左下/右上/右下）
  try {
    body.getTables().forEach(table => {
      for (let r = 0; r < table.getNumRows(); r++) {
        const row = table.getRow(r), rowText = row.getText();
        if (row.getNumCells() < 5) continue;
        const leftCell = row.getCell(2), rightCell = row.getCell(4);

        if (rowText.includes('關節活動度')) {
          const d = [].concat(getFormDataByKey('關節活動度') || []);
          if (d[0]) updateGridLimbText(leftCell, '上肢', d[0], 'ROM');
          if (d[1]) updateGridLimbText(leftCell, '下肢', d[1], 'ROM');
          if (d[2]) updateGridLimbText(rightCell, '上肢', d[2], 'ROM');
          if (d[3]) updateGridLimbText(rightCell, '下肢', d[3], 'ROM');
        } else if (rowText.includes('肌肉張力')) {
          const d = [].concat(getFormDataByKey('肌肉張力') || []);
          if (d[0]) updateGridLimbText(leftCell, '上肢', d[0], 'TONE');
          if (d[1]) updateGridLimbText(leftCell, '下肢', d[1], 'TONE');
          if (d[2]) updateGridLimbText(rightCell, '上肢', d[2], 'TONE');
          if (d[3]) updateGridLimbText(rightCell, '下肢', d[3], 'TONE');
        } else if (rowText.includes('肌肉力量')) {
          const d = [].concat(getFormDataByKey('肌肉力量') || []);
          if (d[0]) updateGridLimbText(leftCell, '上肢', d[0], 'POWER');
          if (d[1]) updateGridLimbText(leftCell, '下肢', d[1], 'POWER');
          if (d[2]) updateGridLimbText(rightCell, '上肢', d[2], 'POWER');
          if (d[3]) updateGridLimbText(rightCell, '下肢', d[3], 'POWER');
        }
      }
    });
  } catch (err) {
    Logger.log('⚠️ 網格表格處理失敗：' + err.message);
  }

  // 11. 清除未用標籤並儲存
  body.replaceText('\\{\\{[^}]*\\}\\}', '');
  newDoc.saveAndClose();
  Logger.log('✅ 成功產出評估單：' + newFileName);
}
