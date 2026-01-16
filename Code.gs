/**
 * OpenRouter Google Apps Script for QA Bug Rewriter
 * 
 * MiMo-V2-Flash (free)
 * xiaomi/mimo-v2-flash:free
 * Xiaomi_API_KEY
 * 
 * TNG: DeepSeek R1T2 Chimera (free)
 * tngtech/deepseek-r1t2-chimera:free
 * Deepseek_API_KEY
 * 
 * Devstral 2 2512 (free)
 * mistralai/devstral-2512:free
 * Devstral_API_KEY
 */

const CONFIG = {
  OPENROUTER: {
    url: 'https://openrouter.ai/api/v1/chat/completions',
    model: 'xiaomi/mimo-v2-flash:free',
    keyName: 'Xiaomi_API_KEY'
  },
  BATCH_SIZE: 50,
  MAX_TOKENS: 2000,
  TEMPERATURE: 0.1
};

let apiKey = null;
const processedTexts = new Map();

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI Tools')
    .addItem('Rewrite cells', 'rewriteSelectedCell')
    .addItem('Generate titles', 'generateTitles')
    .addItem('Show AI Panel', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: 'Google Sans', Arial, sans-serif;
            margin: 0;
            padding: 10px;
            background: #f8f9fa;
          }
          
          .panel {
            background: white;
            border-radius: 8px;
            padding: 16px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
          }
          
          h3 {
            margin: 0 0 16px 0;
            color: #202124;
            font-size: 16px;
            font-weight: 500;
          }
          
          .btn {
            width: 100%;
            padding: 12px 16px;
            margin-bottom: 12px;
            border: none;
            border-radius: 4px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
          }
          
          .btn-primary {
            background: #1a73e8;
            color: white;
          }
          
          .btn-primary:hover {
            background: #1557b0;
            box-shadow: 0 2px 4px rgba(0,0,0,0.2);
          }
          
          .btn-secondary {
            background: #34a853;
            color: white;
          }
          
          .btn-secondary:hover {
            background: #2d8e47;
            box-shadow: 0 2px 4px rgba(0,0,0,0.2);
          }
          
          .status {
            margin-top: 16px;
            padding: 12px;
            border-radius: 4px;
            font-size: 13px;
            display: none;
          }
          
          .status.success {
            background: #e6f4ea;
            color: #137333;
            border: 1px solid #34a853;
          }
          
          .status.error {
            background: #fce8e6;
            color: #c5221f;
            border: 1px solid #ea4335;
          }
          
          .status.info {
            background: #e8f0fe;
            color: #1967d2;
            border: 1px solid #1a73e8;
          }
          
          .icon {
            width: 18px;
            height: 18px;
          }
          
          .divider {
            height: 1px;
            background: #e8eaed;
            margin: 16px 0;
          }
          
          .info-text {
            font-size: 12px;
            color: #5f6368;
            margin-top: 8px;
            line-height: 1.5;
          }
        </style>
      </head>
      <body>
        <div class="panel">
          <h3>🤖 AI Tools</h3>
          
          <button class="btn btn-primary" onclick="rewriteCells()">
            <svg class="icon" viewBox="0 0 24 24" fill="currentColor">
              <path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/>
            </svg>
            Rewrite Selected Cells
          </button>
          
          <button class="btn btn-secondary" onclick="generateTitles()">
            <svg class="icon" viewBox="0 0 24 24" fill="currentColor">
              <path d="M5 4v3h5.5v12h3V7H19V4z"/>
            </svg>
            Generate Titles
          </button>
          
          <div class="divider"></div>
          
          <div class="info-text">
            <strong>Rewrite:</strong> Improves clarity and correctness of selected cells.<br><br>
            <strong>Generate Titles:</strong> Creates short titles in the column to the left of selected cells.
          </div>
          
          <div id="status" class="status"></div>
        </div>
        
        <script>
          function showStatus(message, type) {
            const status = document.getElementById('status');
            status.textContent = message;
            status.className = 'status ' + type;
            status.style.display = 'block';
            
            if (type === 'success') {
              setTimeout(() => {
                status.style.display = 'none';
              }, 3000);
            }
          }
          
          function rewriteCells() {
            showStatus('Processing...', 'info');
            google.script.run
              .withSuccessHandler(() => showStatus('✓ Cells rewritten successfully!', 'success'))
              .withFailureHandler((error) => showStatus('✗ Error: ' + error.message, 'error'))
              .rewriteSelectedCell();
          }
          
          function generateTitles() {
            showStatus('Generating titles...', 'info');
            google.script.run
              .withSuccessHandler(() => showStatus('✓ Titles generated successfully!', 'success'))
              .withFailureHandler((error) => showStatus('✗ Error: ' + error.message, 'error'))
              .generateTitles();
          }
        </script>
      </body>
    </html>
  `)
    .setTitle('AI Tools')
    .setWidth(280);
  
  SpreadsheetApp.getUi().showSidebar(html);
}

function validateApiKey() {
  if (!apiKey) {
    apiKey = PropertiesService.getScriptProperties().getProperty(CONFIG.OPENROUTER.keyName);
  }
  
  if (!apiKey || apiKey.trim() === '') {
    throw new Error(`API Key not found. Please set ${CONFIG.OPENROUTER.keyName} in Script Properties (Extensions > Apps Script > Project Settings > Script Properties)`);
  }
  
  return apiKey;
}

function validateSelection(range) {
  if (!range) {
    throw new Error('No cells selected. Please select cells to process.');
  }
  
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();
  
  if (numRows === 0 || numCols === 0) {
    throw new Error('Invalid selection. Please select at least one cell.');
  }
  
  return true;
}

function rewriteSelectedCell() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getActiveRange();
    
    validateSelection(range);
    validateApiKey();

    const values = range.getValues();
    const cellsToProcess = [];
    
    for (let i = 0; i < values.length; i++) {
      for (let j = 0; j < values[i].length; j++) {
        const text = values[i][j] ? values[i][j].toString().trim() : '';
        if (text && !processedTexts.has(text)) {
          cellsToProcess.push({ row: i, col: j, text });
        }
      }
    }

    if (cellsToProcess.length === 0) {
      throw new Error('No new content to rewrite. All selected cells are either empty or already processed.');
    }

    SpreadsheetApp.getActive().toast(`Processing ${cellsToProcess.length} cells...`, 'AI Rewrite', 1);

    for (let i = 0; i < cellsToProcess.length; i += CONFIG.BATCH_SIZE) {
      const batch = cellsToProcess.slice(i, Math.min(i + CONFIG.BATCH_SIZE, cellsToProcess.length));
      const rewritten = processBatch(batch, 'rewrite');
      
      batch.forEach((cell, idx) => {
        if (rewritten[idx]) {
          values[cell.row][cell.col] = rewritten[idx];
          processedTexts.set(cell.text, rewritten[idx]);
        }
      });
    }
    
    range.setValues(values);
    SpreadsheetApp.getActive().toast('Done', 'Success', 1);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    throw error;
  }
}

function generateTitles() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getActiveRange();
    
    validateSelection(range);
    validateApiKey();

    const startRow = range.getRow();
    const startCol = range.getColumn();
    const values = range.getValues();
    const cellsToProcess = [];
    
    if (startCol === 1) {
      throw new Error('Cannot generate titles for column A (no column to the left). Please select cells in column B or beyond.');
    }
    
    for (let i = 0; i < values.length; i++) {
      for (let j = 0; j < values[i].length; j++) {
        const text = values[i][j] ? values[i][j].toString().trim() : '';
        const actualCol = startCol + j;
        
        if (text) {
          cellsToProcess.push({ 
            row: i, 
            col: j, 
            text: text,
            sheetRow: startRow + i,
            sheetCol: actualCol
          });
        }
      }
    }

    if (cellsToProcess.length === 0) {
      throw new Error('No content to generate titles for. All selected cells are empty.');
    }

    SpreadsheetApp.getActive().toast(`Generating titles for ${cellsToProcess.length} cells...`, 'AI Title Generation', 1);

    for (let i = 0; i < cellsToProcess.length; i += CONFIG.BATCH_SIZE) {
      const batch = cellsToProcess.slice(i, Math.min(i + CONFIG.BATCH_SIZE, cellsToProcess.length));
      const titles = processBatch(batch, 'title');
      
      batch.forEach((cell, idx) => {
        if (titles[idx]) {
          const titleCol = cell.sheetCol - 1;
          sheet.getRange(cell.sheetRow, titleCol).setValue(titles[idx]);
        }
      });
    }
    
    SpreadsheetApp.getActive().toast('Titles generated successfully', 'Success', 1);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    throw error;
  }
}

function processBatch(batch, type) {
  if (!batch || batch.length === 0) {
    throw new Error('No data to process in batch.');
  }
  
  const batchText = batch.map((cell, i) => `${i + 1}. ${cell.text}`).join('\n\n');
  
  const systemPrompt = type === 'title'
    ? 'Generate a short, concise title (2-5 words) for each numbered description. The title should capture the main topic or subject. Return only the numbered titles in the same order, nothing else.'
    : 'Rewrite each numbered bug description from a QA perspective. Make the description clear, specific, and easy for developers to understand. Only improve the description text itself - do not add sections like "Expected:", "Actual:", "Steps:", or any formatting. Keep the same core issue but express it professionally and clearly. Return only the numbered rewritten descriptions in the same order, nothing else.';

  const payload = JSON.stringify({
    model: CONFIG.OPENROUTER.model,
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: batchText }
    ],
    max_tokens: CONFIG.MAX_TOKENS,
    temperature: CONFIG.TEMPERATURE
  });

  const headers = {
    'Content-Type': 'application/json',
    'Authorization': `Bearer ${apiKey}`
  };

  let response;
  try {
    response = UrlFetchApp.fetch(CONFIG.OPENROUTER.url, {
      method: 'post',
      contentType: 'application/json',
      headers: headers,
      payload: payload,
      muteHttpExceptions: true
    });
  } catch (error) {
    throw new Error(`Network error: ${error.message}. Please check your internet connection.`);
  }

  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  if (responseCode !== 200) {
    let errorMessage = 'API request failed';
    
    try {
      const error = JSON.parse(responseText);
      errorMessage = error.error?.message || error.message || `HTTP ${responseCode}: ${responseText}`;
    } catch (parseError) {
      errorMessage = `HTTP ${responseCode}: ${responseText}`;
    }
    
    if (responseCode === 401) {
      throw new Error(`Authentication failed. Please check your API key in Script Properties.`);
    } else if (responseCode === 429) {
      throw new Error(`Rate limit exceeded. Please wait a moment and try again.`);
    } else if (responseCode === 500 || responseCode === 502 || responseCode === 503) {
      throw new Error(`OpenRouter server error (${responseCode}). Please try again later.`);
    }
    
    throw new Error(errorMessage);
  }

  let data;
  try {
    data = JSON.parse(responseText);
  } catch (error) {
    throw new Error(`Invalid response from API. Could not parse JSON: ${error.message}`);
  }

  if (!data.choices || !data.choices[0] || !data.choices[0].message || !data.choices[0].message.content) {
    throw new Error('Unexpected API response structure. Missing content in response.');
  }

  const content = data.choices[0].message.content;

  return parseNumberedResponse(content, batch.length);
}

function parseNumberedResponse(content, expectedCount) {
  if (!content || typeof content !== 'string') {
    throw new Error('Invalid response content from API.');
  }
  
  const lines = content.split('\n').filter(line => line.trim());
  const results = [];
  
  for (let i = 0; i < expectedCount; i++) {
    const pattern = new RegExp(`^${i + 1}\\.?\\s*(.+)$`, 'm');
    const match = content.match(pattern);
    results.push(match ? match[1].trim() : lines[i]?.replace(/^\d+\.\s*/, '').trim() || '');
  }
  
  return results;
}