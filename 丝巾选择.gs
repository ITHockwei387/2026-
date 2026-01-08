// ============================================================
// STANDALONE SCARF & SHAMPOO SELECTION FORM FOR 风水大会 - FINAL VERSION
// WITH RED GRADIENT BACKGROUND
// ============================================================

function smartSplit(str) {
  const parts = [];
  let currentPart = '';
  let depth = 0;
  
  for (let i = 0; i < str.length; i++) {
    const char = str[i];
    
    if (char === '（' || char === '(') {
      depth++;
      currentPart += char;
    } else if (char === '）' || char === ')') {
      depth--;
      currentPart += char;
    } else if (char === '+' && depth === 0) {
      if (currentPart.trim()) {
        parts.push(currentPart.trim());
      }
      currentPart = '';
    } else {
      currentPart += char;
    }
  }
  
  if (currentPart.trim()) {
    parts.push(currentPart.trim());
  }
  
  return parts;
}

const SCARF_CONFIG = {
  SHEET_NAME: '风水大会',
  START_ROW: 2,
  
  COLUMNS: {
    TIMESTAMP: 1,
    ORDER_ID: 2,
    NAME: 3,
    EMAIL: 4,
    PHONE: 5,
    MAIN_PRODUCT: 6,
    QUANTITY: 7,
    ORDER_SUMMARY: 8,
    TOTAL_PRICE: 9,
    STATUS: 10,
    SCARF_SELECTION: 11,      // Column K - Scarf selection
    SHAMPOO_SELECTION: 12,    // Column L - Shampoo selection
    SCARF_LINK: 13,           // Column M - Result link
    COMPLETION_STATUS: 14     // Column N - Complete status
  },
  
  SCARF_OPTIONS: [
    '财富之光（冰丝）',
    '成就之光（冰丝）',
    '挚爱之光（冰丝）',
    '丰盛满钞（冰丝）',
    '生命之花（冰丝）',
    '财富版图（冰丝）',
    '财富版图（黑白款）（冰丝）',
    '财富十二宫（冰丝）'
  ],
  
  SHAMPOO_OPTIONS: [
    '爆爽洗发水',
    '爆顺洗发水',
    '爆发洗发水'
  ],
  
  PRODUCTS: {
    SET_A: { name: 'Set A 配套', scarfQty: 1, shampooQty: 1 },
    SET_B: { name: 'Set B 配套', scarfQty: 2, shampooQty: 2 }
  }
};

function doGet(e) {
  try {
    Logger.log('========== doGet START ==========');
    Logger.log('Parameters: ' + JSON.stringify(e.parameter));
    
    const action = e.parameter.action || 'form';
    const email = e.parameter.email || '';
    
    Logger.log('Action: ' + action);
    Logger.log('Email: ' + email);
    
    if (action === 'view') {
      if (!email) {
        Logger.log('❌ No email provided for view action');
        return HtmlService.createHtmlOutput(createErrorPage('缺少电子邮箱参数'));
      }
      Logger.log('📋 Viewing results for: ' + email);
      return viewResultsByEmail(email);
    }
    
    if (action === 'form') {
      if (email) {
        Logger.log('🔍 Checking if email already completed: ' + email);
        const completionStatus = checkIfAlreadyCompleted(email);
        Logger.log('Completion status: ' + JSON.stringify(completionStatus));
        
        if (completionStatus.completed) {
          Logger.log('✅ Email already completed, showing results');
          return viewResultsByEmail(email);
        }
      }
      Logger.log('📝 Showing form');
      return createScarfSelectionForm();
    }
    
    Logger.log('📝 Default: showing form');
    return createScarfSelectionForm();
    
  } catch (error) {
    Logger.log('❌ ERROR in doGet: ' + error);
    Logger.log('Stack: ' + error.stack);
    return HtmlService.createHtmlOutput(createErrorPage('系统错误: ' + error.message));
  }
}

function checkIfAlreadyCompleted(email) {
  try {
    Logger.log('🔍 checkIfAlreadyCompleted for: ' + email);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SCARF_CONFIG.SHEET_NAME);
    
    if (!sh) {
      Logger.log('❌ Sheet not found: ' + SCARF_CONFIG.SHEET_NAME);
      return { completed: false, error: 'Sheet not found' };
    }
    
    const lastRow = sh.getLastRow();
    Logger.log('📊 Last row: ' + lastRow);
    
    for (let i = SCARF_CONFIG.START_ROW; i <= lastRow; i++) {
      const emailCell = sh.getRange(i, SCARF_CONFIG.COLUMNS.EMAIL).getValue();
      
      if (emailCell && emailCell.toString().toLowerCase().trim() === email.toLowerCase().trim()) {
        Logger.log('✅ Found email at row ' + i);
        
        const completionStatus = sh.getRange(i, SCARF_CONFIG.COLUMNS.COMPLETION_STATUS).getValue();
        Logger.log('Completion status value: "' + completionStatus + '"');
        
        if (completionStatus === 'Complete') {
          Logger.log('✅ Email has completed selection');
          return { completed: true, row: i, email: email };
        } else {
          Logger.log('⏳ Email found but not completed yet');
        }
      }
    }
    
    Logger.log('❌ Email not found in sheet');
    return { completed: false };
    
  } catch (error) {
    Logger.log('❌ Error in checkIfAlreadyCompleted: ' + error);
    return { completed: false, error: error.toString() };
  }
}

function viewResultsByEmail(email) {
  try {
    Logger.log('📋 viewResultsByEmail for: ' + email);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SCARF_CONFIG.SHEET_NAME);
    
    if (!sh) {
      Logger.log('❌ Sheet not found');
      return HtmlService.createHtmlOutput(createErrorPage('系统错误：找不到数据表'));
    }
    
    const lastRow = sh.getLastRow();
    let targetRow = -1;
    
    for (let i = SCARF_CONFIG.START_ROW; i <= lastRow; i++) {
      const emailCell = sh.getRange(i, SCARF_CONFIG.COLUMNS.EMAIL).getValue();
      if (emailCell && emailCell.toString().toLowerCase().trim() === email.toLowerCase().trim()) {
        targetRow = i;
        Logger.log('✅ Found email at row ' + i);
        break;
      }
    }
    
    if (targetRow === -1) {
      Logger.log('❌ Email not found in records');
      return HtmlService.createHtmlOutput(createErrorPage('找不到此邮箱的记录'));
    }
    
    const completionStatus = sh.getRange(targetRow, SCARF_CONFIG.COLUMNS.COMPLETION_STATUS).getValue();
    Logger.log('Completion status: ' + completionStatus);
    
    if (completionStatus !== 'Complete') {
      Logger.log('❌ Not completed yet');
      return HtmlService.createHtmlOutput(createErrorPage('您尚未完成选择'));
    }
    
    const scarfSelection = sh.getRange(targetRow, SCARF_CONFIG.COLUMNS.SCARF_SELECTION).getValue();
    const shampooSelection = sh.getRange(targetRow, SCARF_CONFIG.COLUMNS.SHAMPOO_SELECTION).getValue();
    const name = sh.getRange(targetRow, SCARF_CONFIG.COLUMNS.NAME).getValue();
    const orderSummary = sh.getRange(targetRow, SCARF_CONFIG.COLUMNS.ORDER_SUMMARY).getValue();
    
    Logger.log('✅ Creating results page');
    return createResultsPage(email, name, scarfSelection, shampooSelection, orderSummary);
    
  } catch (error) {
    Logger.log('❌ Error in viewResultsByEmail: ' + error);
    return HtmlService.createHtmlOutput(createErrorPage('加载结果时出错: ' + error.message));
  }
}

function getCustomerOrderInfo(email) {
  try {
    Logger.log('📦 getCustomerOrderInfo for: ' + email);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SCARF_CONFIG.SHEET_NAME);
    
    if (!sh) {
      Logger.log('❌ Sheet not found');
      return { success: false, error: 'Sheet not found' };
    }
    
    const lastRow = sh.getLastRow();
    Logger.log('📊 Last row: ' + lastRow);
    
    for (let i = SCARF_CONFIG.START_ROW; i <= lastRow; i++) {
      const emailCell = sh.getRange(i, SCARF_CONFIG.COLUMNS.EMAIL).getValue();
      
      if (emailCell && emailCell.toString().toLowerCase().trim() === email.toLowerCase().trim()) {
        const name = sh.getRange(i, SCARF_CONFIG.COLUMNS.NAME).getValue();
        const orderSummary = sh.getRange(i, SCARF_CONFIG.COLUMNS.ORDER_SUMMARY).getValue();
        
        Logger.log('📧 Email matched at row ' + i);
        Logger.log('📦 Order Summary: ' + orderSummary);
        
        let setAQty = 0;
        let setBQty = 0;
        
        const parts = smartSplit((orderSummary || '').toString());
        Logger.log('📋 Parsed parts: ' + JSON.stringify(parts));
        
        for (let j = 0; j < parts.length; j++) {
          const part = parts[j].trim();
          
          if (part.indexOf('Set A') !== -1 && part.indexOf('Set B') === -1) {
            const matches = part.match(/[xX×]\s*(\d+)/);
            setAQty = matches && matches[1] ? parseInt(matches[1]) : 1;
            Logger.log('✅ Set A x' + setAQty + ' from: ' + part);
          }
          else if (part.indexOf('Set B') !== -1 && part.indexOf('Set A') === -1) {
            const matches = part.match(/[xX×]\s*(\d+)/);
            setBQty = matches && matches[1] ? parseInt(matches[1]) : 1;
            Logger.log('✅ Set B x' + setBQty + ' from: ' + part);
          }
        }
        
        const scarfQty = (setAQty * 1) + (setBQty * 2);
        const shampooQty = (setAQty * 1) + (setBQty * 2);
        Logger.log('🎯 Total scarves: ' + scarfQty);
        Logger.log('🧴 Total shampoos: ' + shampooQty);
        
        if (scarfQty === 0) {
          Logger.log('❌ No valid set found');
          return { success: false, error: '无法识别您购买的配套类型。请联系客服。' };
        }
        
        let setTypeDisplay = '';
        if (setAQty > 0 && setBQty > 0) setTypeDisplay = 'Set A x' + setAQty + ' + Set B x' + setBQty;
        else if (setAQty > 0) setTypeDisplay = 'Set A x' + setAQty;
        else if (setBQty > 0) setTypeDisplay = 'Set B x' + setBQty;
        
        Logger.log('✅ Success: ' + setTypeDisplay + ' = ' + scarfQty + ' scarves + ' + shampooQty + ' shampoos');
        
        return {
          success: true,
          row: i,
          name: name || 'N/A',
          setType: setTypeDisplay,
          scarfQty: scarfQty,
          shampooQty: shampooQty,
          orderSummary: orderSummary,
          phone: sh.getRange(i, SCARF_CONFIG.COLUMNS.PHONE).getValue() || ''
        };
      }
    }
    
    Logger.log('❌ Email not found in records');
    return { success: false, error: '找不到此邮箱的订单记录' };
    
  } catch (error) {
    Logger.log('❌ Error in getCustomerOrderInfo: ' + error);
    return { success: false, error: error.toString() };
  }
}

function processScarfSelection(formData) {
  try {
    Logger.log('💾 processScarfSelection');
    Logger.log('Data: ' + JSON.stringify(formData));
    
    if (!formData.email || !formData.selectedScarves || formData.selectedScarves.length === 0) {
      Logger.log('❌ Missing email or scarves');
      return { success: false, error: '请选择丝巾款式' };
    }
    
    if (!formData.selectedShampoos || formData.selectedShampoos.length === 0) {
      Logger.log('❌ Missing shampoos');
      return { success: false, error: '请选择洗发水款式' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SCARF_CONFIG.SHEET_NAME);
    
    if (!sh) {
      Logger.log('❌ Sheet not found');
      return { success: false, error: 'Sheet not found' };
    }
    
    const orderInfo = getCustomerOrderInfo(formData.email);
    if (!orderInfo.success) {
      Logger.log('❌ Order info failed: ' + orderInfo.error);
      return { success: false, error: orderInfo.error };
    }
    
    if (formData.selectedScarves.length !== orderInfo.scarfQty) {
      Logger.log('❌ Wrong scarf quantity: ' + formData.selectedScarves.length + ' vs ' + orderInfo.scarfQty);
      return { success: false, error: '您需要选择 ' + orderInfo.scarfQty + ' 条丝巾' };
    }
    
    if (formData.selectedShampoos.length !== orderInfo.shampooQty) {
      Logger.log('❌ Wrong shampoo quantity: ' + formData.selectedShampoos.length + ' vs ' + orderInfo.shampooQty);
      return { success: false, error: '您需要选择 ' + orderInfo.shampooQty + ' 瓶洗发水' };
    }
    
    const targetRow = orderInfo.row;
    const currentCompletionStatus = sh.getRange(targetRow, SCARF_CONFIG.COLUMNS.COMPLETION_STATUS).getValue();
    
    if (currentCompletionStatus === 'Complete') {
      Logger.log('❌ Already submitted');
      return { success: false, error: '您已经提交过选择了，无法重复提交' };
    }
    
    // Format scarf selection
    const scarfCount = {};
    for (let i = 0; i < formData.selectedScarves.length; i++) {
      const scarf = formData.selectedScarves[i];
      scarfCount[scarf] = (scarfCount[scarf] || 0) + 1;
    }
    
    const formattedScarfParts = [];
    for (const scarf in scarfCount) {
      const qty = scarfCount[scarf];
      formattedScarfParts.push(qty > 1 ? scarf + ' x' + qty : scarf);
    }
    
    const scarfSelectionText = formattedScarfParts.join(' + ');
    
    // Format shampoo selection
    const shampooCount = {};
    for (let i = 0; i < formData.selectedShampoos.length; i++) {
      const shampoo = formData.selectedShampoos[i];
      shampooCount[shampoo] = (shampooCount[shampoo] || 0) + 1;
    }
    
    const formattedShampooParts = [];
    for (const shampoo in shampooCount) {
      const qty = shampooCount[shampoo];
      formattedShampooParts.push(qty > 1 ? shampoo + ' x' + qty : shampoo);
    }
    
    const shampooSelectionText = formattedShampooParts.join(' + ');
    
    const scriptUrl = ScriptApp.getService().getUrl();
    const resultLink = scriptUrl + '?action=view&email=' + encodeURIComponent(formData.email);
    
    Logger.log('🔗 Result link: ' + resultLink);
    
    sh.getRange(targetRow, SCARF_CONFIG.COLUMNS.SCARF_SELECTION).setValue(scarfSelectionText);
    sh.getRange(targetRow, SCARF_CONFIG.COLUMNS.SHAMPOO_SELECTION).setValue(shampooSelectionText);
    sh.getRange(targetRow, SCARF_CONFIG.COLUMNS.SCARF_LINK).setValue(resultLink);
    sh.getRange(targetRow, SCARF_CONFIG.COLUMNS.COMPLETION_STATUS).setValue('Complete');
    
    Logger.log('✅ Updated row ' + targetRow);
    Logger.log('🧣 Scarves (Column K): ' + scarfSelectionText);
    Logger.log('🧴 Shampoos (Column L): ' + shampooSelectionText);
    Logger.log('🔗 Link (Column M): ' + resultLink);
    Logger.log('✅ Set completion status to Complete in column N');
    
    return {
      success: true,
      scarves: formattedScarfParts,
      shampoos: formattedShampooParts,
      name: orderInfo.name,
      setType: orderInfo.setType,
      resultLink: resultLink
    };
    
  } catch (error) {
    Logger.log('❌ Error in processScarfSelection: ' + error);
    return { success: false, error: error.toString() };
  }
}

function createErrorPage(message) {
  return '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>错误</title><style>body{font-family:"Microsoft YaHei",Arial,sans-serif;background:#cca983;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:10px}.error-container{background:white;border-radius:20px;padding:40px;max-width:500px;width:100%;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,0.3)}h2{color:#9e0b0f;margin-bottom:20px;font-size:32px}p{color:#333;font-size:18px;line-height:1.6}</style></head><body><div class="error-container"><h2>❌ 错误</h2><p>' + message + '</p></div></body></html>';
}

function createResultsPage(email, name, scarfSelection, shampooSelection, orderSummary) {
  const scarves = scarfSelection.split(' + ');
  const shampoos = shampooSelection ? shampooSelection.split(' + ') : [];
  
  // Get phone number from sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SCARF_CONFIG.SHEET_NAME);
  let phone = '';
  
  for (let i = SCARF_CONFIG.START_ROW; i <= sh.getLastRow(); i++) {
    const emailCell = sh.getRange(i, SCARF_CONFIG.COLUMNS.EMAIL).getValue();
    if (emailCell && emailCell.toString().toLowerCase().trim() === email.toLowerCase().trim()) {
      phone = sh.getRange(i, SCARF_CONFIG.COLUMNS.PHONE).getValue() || '';
      break;
    }
  }
  
  const imageMap = {
    '财富之光（冰丝）': 'https://lh3.googleusercontent.com/d/1GChEXSN5Mf8yabm7TvUh-UWx3ZMsWG0p',
    '成就之光（冰丝）': 'https://lh3.googleusercontent.com/d/1hdzn5yMw7mLv0EUm67-_4q2lTmdP4Z9c',
    '挚爱之光（冰丝）': 'https://lh3.googleusercontent.com/d/12YlxaNRSGn5inIhijfE6UnrwBUtYrsVx',
    '丰盛满钞（冰丝）': 'https://lh3.googleusercontent.com/d/1bibLK0oQINwaop-jJeEiTHvVTUGUeRWy',
    '生命之花（冰丝）': 'https://lh3.googleusercontent.com/d/1szujBgUcDoOu-77QchWcOOWu8JaFcIzG',
    '财富版图（冰丝）': 'https://lh3.googleusercontent.com/d/19hgGBvMVu9PKqF5s9PPbn5ty8Q_7HgTD',
    '财富版图（黑白款）（冰丝）': 'https://lh3.googleusercontent.com/d/18c8QDQPn83h6B2EtMgfzepRr0CdLDzqD',
    '财富十二宫（冰丝）': 'https://lh3.googleusercontent.com/d/1eTA3tuGYy89avF_Co9mroYZfZE6uNfQs',
    '爆爽洗发水': 'https://lh3.googleusercontent.com/d/1A8S9f-06nFsbGKncp662BLOUiU3Yhkvr',
    '爆顺洗发水': 'https://lh3.googleusercontent.com/d/1Jw-PXpObED7oxg2ggtvdlo08H3MJntjU',
    '爆发洗发水': 'https://lh3.googleusercontent.com/d/1SceOlEfmMdQ2ZjI2h_T5Etxlkb9x04L9'
  };
  
  let scarvesHtml = '';
  for (let i = 0; i < scarves.length; i++) {
    const scarfText = scarves[i].trim();
    const parts = scarfText.split(' x');
    const scarfName = parts[0].trim();
    
    Logger.log('🔍 Looking up scarf image for: "' + scarfName + '"');
    
    const imageUrl = imageMap[scarfName];
    
    if (imageUrl) {
      Logger.log('✅ Found image: ' + imageUrl);
    } else {
      Logger.log('❌ No image found for: ' + scarfName);
    }
    
    const imageHtml = imageUrl ? 
      '<img src="' + imageUrl + '" style="width:100px;height:100px;object-fit:cover;border-radius:12px;margin-bottom:15px;border:3px solid #9e0b0f;" onerror="this.style.display=\'none\'">' :
      '<div class="scarf-icon">🧣</div>';
    
    scarvesHtml += '<div class="scarf-item">' + imageHtml + '<div class="scarf-name">' + scarfText + '</div></div>';
  }
  
  let shampoosHtml = '';
  for (let i = 0; i < shampoos.length; i++) {
    const shampooText = shampoos[i].trim();
    const parts = shampooText.split(' x');
    const shampooName = parts[0].trim();
    
    Logger.log('🔍 Looking up shampoo image for: "' + shampooName + '"');
    
    const imageUrl = imageMap[shampooName];
    
    if (imageUrl) {
      Logger.log('✅ Found image: ' + imageUrl);
    } else {
      Logger.log('❌ No image found for: ' + shampooName);
    }
    
    const imageHtml = imageUrl ? 
      '<img src="' + imageUrl + '" style="width:100px;height:100px;object-fit:cover;border-radius:12px;margin-bottom:15px;border:3px solid #9e0b0f;" onerror="this.style.display=\'none\'">' :
      '<div class="scarf-icon">🧴</div>';
    
    shampoosHtml += '<div class="scarf-item">' + imageHtml + '<div class="scarf-name">' + shampooText + '</div></div>';
  }
  
  return HtmlService.createHtmlOutput('<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>选择结果</title><style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:"Microsoft YaHei",Arial,sans-serif;background:#cca983;min-height:100vh;padding:10px}.container{max-width:600px;width:100%;margin:0 auto;background:white;border-radius:20px;box-shadow:0 20px 60px rgba(0,0,0,0.3);overflow:hidden}.header{background:linear-gradient(135deg,#9e0b0f 0%,#670000 100%);color:white;padding:40px 30px;text-align:center}.header h1{font-size:48px;margin:0;font-weight:bold;letter-spacing:8px;text-transform:uppercase}.header p{margin:12px 0 0 0;font-size:24px;letter-spacing:3px;font-weight:600}.customer-info{background:#ffe6e6;padding:20px;text-align:center;border-bottom:3px solid #9e0b0f}.customer-info h2{color:#000000;margin-bottom:10px;font-size:24px}.customer-info p{color:#000000;font-size:18px;margin:8px 0;font-weight:500}.info-notice{background:#fff9e6;border-left:4px solid #9e0b0f;padding:15px;margin:20px;border-radius:6px}.info-notice p{margin:8px 0;font-size:16px;color:#333;line-height:1.6}.info-notice strong{color:#9e0b0f}.results-content{padding:30px}.section-title{text-align:center;color:#9e0b0f;margin-bottom:20px;font-size:24px;font-weight:bold;padding:10px;background:#ffe6e6;border-radius:10px}.scarf-item{background:white;border:2px solid #9e0b0f;border-radius:12px;padding:25px;margin-bottom:20px;text-align:center;transition:transform 0.3s}.scarf-item:hover{transform:scale(1.02)}.scarf-icon{font-size:64px;margin-bottom:15px}.scarf-name{color:#9e0b0f;font-size:28px;font-weight:bold;letter-spacing:2px}.section-divider{height:3px;background:linear-gradient(to right,transparent,#9e0b0f,transparent);margin:30px 0}.footer{background:#9e0b0f;color:white;padding:20px;text-align:center;font-size:15px}.footer p{margin:8px 0;font-weight:500}.footer-phones{display:flex;gap:15px;justify-content:center;margin-top:10px;font-size:16px}@media (max-width:480px){.header h1{font-size:36px;letter-spacing:4px}.header p{font-size:20px}.customer-info h2{font-size:20px}.customer-info p{font-size:16px}.scarf-name{font-size:24px}}</style></head><body><div class="container"><div class="header"><h1>New Year Package</h1><p>选择结果</p></div><div class="customer-info"><h2>👤 ' + name + '</h2><p>📧 ' + email + '</p>' + (phone ? '<p>📞 ' + phone + '</p>' : '') + '<p>📦 ' + orderSummary + '</p></div><div class="info-notice"><p><strong>💡 温馨提示：</strong></p><p>您可以随时使用此链接查看您的选择记录。请保存或收藏此页面以便日后查看。</p><p>⚠️ 建议您截屏保存，以备不时之需。</p></div><div class="results-content"><h2 class="section-title">🧴 您选择的洗发水款式</h2>' + shampoosHtml + '<div class="section-divider"></div><h2 class="section-title">🧣 您选择的丝巾款式</h2>' + scarvesHtml + '</div><div class="footer"><p><strong>恭喜你！您的选择已确认！</strong></p><p><strong>洗发水和丝巾将会和配套一起寄出。如需更换可以联系客服更换哦。</strong></p><div class="footer-phones"><span>📞 +6013-928 4699</span><span>📞 +6013-530 8863</span></div></div></div></body></html>')
    .setTitle('New Year Package - 选择结果')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function createScarfSelectionForm() {
  const htmlTemplate = HtmlService.createHtmlOutputFromFile('scarf_form_template');
  return htmlTemplate
    .setTitle('New Year Package - 丝巾 & 洗发水选择')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
