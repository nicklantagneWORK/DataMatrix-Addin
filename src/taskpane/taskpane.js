/* ─────────────────────────────────────────────────────
   Data Matrix Excel Add-in — taskpane.js
   Offline ECC 200 generation via bwip-js + Office.js
   ───────────────────────────────────────────────────── */

'use strict';

// ════════════════════════════════════════════════════════
//  ANSI MH10.8.2-2016 DATA IDENTIFIER DEFINITIONS
// ════════════════════════════════════════════════════════
const ANSI_IDENTIFIERS = [
  { di: '1P',  label: 'Part Number',              placeholder: 'e.g. PN-12345' },
  { di: 'S',   label: 'Serial Number',             placeholder: 'e.g. SN-98765' },
  { di: 'L',   label: 'Lot / Batch Number',        placeholder: 'e.g. LOT-A01'  },
  { di: 'D',   label: 'Date of Manufacture',       placeholder: 'YYYYMMDD'       },
  { di: '9D',  label: 'Expiration Date',           placeholder: 'YYYYMMDD'       },
  { di: 'V',   label: 'Supplier / Vendor Code',   placeholder: 'e.g. VEN001'    },
  { di: 'Q',   label: 'Quantity',                  placeholder: 'e.g. 100'       },
  { di: '11K', label: 'Customer Order Number',     placeholder: 'e.g. PO-55500'  },
  { di: '6P',  label: 'Customer Part Number',      placeholder: 'e.g. CPN-00X'  },
  { di: '4L',  label: 'Country of Origin',         placeholder: 'e.g. US'        },
  { di: 'W',   label: 'Weight (kg)',               placeholder: 'e.g. 1.254'     },
  { di: '2P',  label: 'Alternative Part Number',   placeholder: 'e.g. APN-9'     },
  { di: 'T',   label: 'Traceability Code',         placeholder: 'e.g. TC-001'    },
  { di: '17V', label: 'Revision Level',            placeholder: 'e.g. RevA'      },
  { di: 'R',   label: 'Purchase Order Number',     placeholder: 'e.g. PO-77700'  },
];

// ════════════════════════════════════════════════════════
//  INIT
// ════════════════════════════════════════════════════════
Office.onReady(({ host }) => {
  if (host !== Office.HostType.Excel) {
    showStatus('status-single', 'error', 'This add-in requires Microsoft Excel.');
    return;
  }
  initTabs();
  initSingleTab();
  initBatchTab();
  initLiveTab();
  initAnsiTab();
});

// ════════════════════════════════════════════════════════
//  TAB NAVIGATION
// ════════════════════════════════════════════════════════
function initTabs() {
  document.querySelectorAll('.tab').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('panel-' + btn.dataset.tab).classList.add('active');
    });
  });
}

// ════════════════════════════════════════════════════════
//  SINGLE TAB
// ════════════════════════════════════════════════════════
function initSingleTab() {
  const input      = document.getElementById('single-input');
  const sizeSelect = document.getElementById('matrix-size');
  const pxInput    = document.getElementById('px-size');
  const canvas     = document.getElementById('preview-canvas');
  const errEl      = document.getElementById('preview-error');
  const injectBtn  = document.getElementById('btn-inject-single');
  const injectSize = document.getElementById('inject-size');

  // Debounced live preview
  let debounceTimer;
  const triggerPreview = () => {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => renderPreview(canvas, errEl, input.value, sizeSelect.value, parseInt(pxInput.value) || 5), 200);
  };

  input.addEventListener('input', triggerPreview);
  sizeSelect.addEventListener('change', triggerPreview);
  pxInput.addEventListener('input', triggerPreview);

  injectBtn.addEventListener('click', async () => {
    const text = input.value.trim();
    if (!text) { showStatus('status-single', 'error', 'Please enter some data to encode.'); return; }

    injectBtn.disabled = true;
    showStatus('status-single', 'info', 'Generating barcode…');

    try {
      const base64 = await generateDataMatrixBase64(text, sizeSelect.value, parseInt(pxInput.value) || 5);
      await injectImageIntoSheet(base64, parseFloat(injectSize.value) || 80);
      showStatus('status-single', 'success', '✓ Barcode injected successfully!');
    } catch (err) {
      showStatus('status-single', 'error', '✗ ' + err.message);
    } finally {
      injectBtn.disabled = false;
    }
  });

  // Initial render placeholder
  renderPreview(canvas, errEl, 'DEMO123', 'auto', 5);
}

// ════════════════════════════════════════════════════════
//  BATCH TAB
// ════════════════════════════════════════════════════════
function initBatchTab() {
  const btn         = document.getElementById('btn-inject-batch');
  const progressWrap= document.getElementById('batch-progress');
  const progressFill= document.getElementById('progress-fill');
  const progressLbl = document.getElementById('progress-label');
  const batchSize   = document.getElementById('batch-size');
  const colOffset   = document.getElementById('batch-col-offset');

  btn.addEventListener('click', async () => {
    btn.disabled = true;
    progressWrap.classList.remove('hidden');
    showStatus('status-batch', 'info', 'Reading selection…');

    try {
      await Excel.run(async (context) => {
        const selection = context.workbook.getSelectedRange();
        selection.load(['values', 'rowIndex', 'columnIndex', 'rowCount']);
        await context.sync();

        const values = selection.values;
        const total  = values.length;
        const startRow = selection.rowIndex;
        const startCol = selection.columnIndex;
        const offset   = parseInt(colOffset.value) || 1;
        const ptSize   = parseFloat(batchSize.value) || 60;

        showStatus('status-batch', 'info', `Generating ${total} barcode(s)…`);

        for (let i = 0; i < total; i++) {
          const text = String(values[i][0]).trim();
          if (!text) continue;

          try {
            const base64 = await generateDataMatrixBase64(text, 'auto', 4);
            const sheet  = context.workbook.worksheets.getActiveWorksheet();

            // Compute pixel position of the target cell
            const targetCell = sheet.getRangeByIndexes(startRow + i, startCol + offset, 1, 1);
            targetCell.load(['left', 'top']);
            await context.sync();

            const shape = sheet.shapes.addImage(base64);
            shape.left   = targetCell.left;
            shape.top    = targetCell.top;
            shape.height = ptSize;
            shape.width  = ptSize;
            shape.name   = `DM_${i}_${Date.now()}`;
            await context.sync();
          } catch (rowErr) {
            console.warn(`Row ${i} failed: ${rowErr.message}`);
          }

          // Update progress
          const pct = Math.round(((i + 1) / total) * 100);
          progressFill.style.width = pct + '%';
          progressLbl.textContent  = `${i + 1} / ${total}`;
        }

        showStatus('status-batch', 'success', `✓ ${total} barcode(s) injected!`);
      });
    } catch (err) {
      showStatus('status-batch', 'error', '✗ ' + err.message);
    } finally {
      btn.disabled = false;
      setTimeout(() => progressWrap.classList.add('hidden'), 3000);
    }
  });
}

// ════════════════════════════════════════════════════════
//  ANSI BUILDER TAB
// ════════════════════════════════════════════════════════
function initAnsiTab() {
  const container = document.getElementById('ansi-fields');
  const assembled = document.getElementById('assembled-code');
  const copyBtn   = document.getElementById('btn-ansi-copy');

  // Build toggle rows
  ANSI_IDENTIFIERS.forEach(({ di, label, placeholder }) => {
    const row = document.createElement('div');
    row.className = 'ansi-row';
    row.dataset.di = di;

    const toggle = document.createElement('button');
    toggle.className = 'ansi-toggle';
    toggle.setAttribute('aria-label', `Toggle ${label}`);

    const diLabel = document.createElement('span');
    diLabel.className = 'ansi-di';
    diLabel.textContent = di;

    const valueInput = document.createElement('input');
    valueInput.className = 'ansi-value';
    valueInput.placeholder = placeholder;
    valueInput.disabled = true;
    valueInput.setAttribute('aria-label', label);

    toggle.addEventListener('click', () => {
      const isOn = toggle.classList.toggle('on');
      valueInput.disabled = !isOn;
      row.classList.toggle('active', isOn);
      if (isOn) valueInput.focus();
      updateAssembled();
    });

    valueInput.addEventListener('input', updateAssembled);

    row.append(toggle, diLabel, valueInput);
    container.appendChild(row);
  });

  function updateAssembled() {
    const parts = [];
    container.querySelectorAll('.ansi-row').forEach(row => {
      const toggle = row.querySelector('.ansi-toggle');
      const di     = row.dataset.di;
      const val    = row.querySelector('.ansi-value').value.trim();
      if (toggle.classList.contains('on')) {
        parts.push(`[)>\x1E06\x1D${di}${val}`);
      }
    });
    const result = parts.length > 0 ? parts.join('\x1D') : '—';
    assembled.textContent = result === '—' ? '—' : result.replace(/[\x1D\x1E]/g, '·');
    assembled.dataset.raw = result;
  }

  copyBtn.addEventListener('click', () => {
    const raw = assembled.dataset.raw || '';
    if (raw === '—' || !raw) return;
    document.getElementById('single-input').value = raw;
    // Switch to Single tab
    document.querySelector('[data-tab="single"]').click();
  });
}

// ════════════════════════════════════════════════════════
//  BARCODE GENERATION (bwip-js)
// ════════════════════════════════════════════════════════

/**
 * Renders a Data Matrix barcode to an offscreen canvas and returns Base64 PNG.
 * @param {string} text - Data to encode
 * @param {string} sizeKey - e.g. 'auto', '14x14', '22x22'
 * @param {number} scale - Module size in pixels (1–20)
 * @returns {Promise<string>} Base64-encoded PNG (no data: prefix)
 */
function generateDataMatrixBase64(text, sizeKey, scale) {
  return new Promise((resolve, reject) => {
    const canvas = document.createElement('canvas');

    const opts = {
      bcid:      'datamatrix',
      text:      text,
      scale:     scale || 5,
      padding:   2,
      backgroundcolor: 'ffffff'
    };

    // Map size key to bwip-js version param
    if (sizeKey && sizeKey !== 'auto') {
      // e.g. "14x14" → version "14x14"
      opts.version = sizeKey;
    }

    try {
      bwipjs.toCanvas(canvas, opts);
      // Strip the "data:image/png;base64," prefix
      resolve(canvas.toDataURL('image/png').split(',')[1]);
    } catch (e) {
      reject(new Error(e.message || 'bwip-js encoding error'));
    }
  });
}

/**
 * Renders a preview into the visible canvas element.
 */
function renderPreview(canvasEl, errEl, text, sizeKey, scale) {
  if (!text || !text.trim()) {
    const ctx = canvasEl.getContext('2d');
    ctx.clearRect(0, 0, canvasEl.width, canvasEl.height);
    canvasEl.width = 0;
    errEl.classList.add('hidden');
    return;
  }

  const opts = { bcid: 'datamatrix', text: text.trim(), scale: scale || 5, padding: 2, backgroundcolor: 'ffffff' };
  if (sizeKey && sizeKey !== 'auto') opts.version = sizeKey;

  try {
    bwipjs.toCanvas(canvasEl, opts);
    errEl.classList.add('hidden');
  } catch (e) {
    errEl.textContent = '⚠ ' + (e.message || 'Encoding error');
    errEl.classList.remove('hidden');
    const ctx = canvasEl.getContext('2d');
    ctx.clearRect(0, 0, canvasEl.width, canvasEl.height);
    canvasEl.width = 0;
  }
}

// ════════════════════════════════════════════════════════
//  OFFICE.JS — INJECT IMAGE INTO SHEET
// ════════════════════════════════════════════════════════
async function injectImageIntoSheet(base64, sizeInPt) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const selection = context.workbook.getSelectedRange();
    selection.load(['left', 'top']);
    await context.sync();

    const shape = sheet.shapes.addImage(base64);
    shape.left   = selection.left;
    shape.top    = selection.top;
    shape.height = sizeInPt;
    shape.width  = sizeInPt;
    // Lock aspect ratio and move/size with cells
    shape.placement = Excel.Placement.twoCell;
    shape.name = 'DataMatrix_' + Date.now();

    await context.sync();
  });
}

// ════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════
// ════════════════════════════════════════════════════════
//  LIVE SYNC (Multiple Watcher Rules & In-Cell Snapping)
// ════════════════════════════════════════════════════════
let liveRules = [];
let liveSyncActive = false;
let liveSyncHandler = null;

function initLiveTab() {
  const sourceIn = document.getElementById('live-source');
  const targetIn = document.getElementById('live-target');
  const btnPickSource = document.getElementById('btn-pick-source');
  const btnPickTarget = document.getElementById('btn-pick-target');
  const addBtn   = document.getElementById('btn-add-rule');
  const listEl   = document.getElementById('rules-list');
  const toggleBtn= document.getElementById('btn-toggle-sync');
  const toggleTxt= document.getElementById('sync-text');
  const toggleIcn= document.getElementById('sync-icon');

  // Utility to fetch active selection from Excel cleanly
  const grabSelection = async (inputEl) => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('address');
        await context.sync();
        // Remove the sheet name (e.g. 'Sheet1!A1:B2' -> 'A1:B2')
        inputEl.value = range.address.includes('!') ? range.address.split('!')[1] : range.address;
      });
    } catch(err) { console.warn("Could not grab selection:", err); }
  };

  btnPickSource.addEventListener('click', () => grabSelection(sourceIn));
  btnPickTarget.addEventListener('click', () => grabSelection(targetIn));

  function renderList() {
    listEl.innerHTML = '';
    if (liveRules.length === 0) {
      listEl.innerHTML = '<span style="color:var(--text-muted);font-size:11px;">No active rules. Highlight cells and add above.</span>';
      return;
    }
    liveRules.forEach((rule, idx) => {
      const item = document.createElement('div');
      item.className = 'rule-item';
      item.innerHTML = `
        <div>
          <span class="rule-path">${rule.source}</span>
          <span class="rule-arrow">→</span>
          <span class="rule-path">${rule.target}</span>
        </div>
        <button class="btn-remove-rule" data-idx="${idx}">×</button>
      `;
      listEl.appendChild(item);
    });

    listEl.querySelectorAll('.btn-remove-rule').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const i = parseInt(e.target.dataset.idx);
        liveRules.splice(i, 1);
        renderList();
      });
    });
  }

  addBtn.addEventListener('click', () => {
    let s = sourceIn.value.trim().toUpperCase();
    let t = targetIn.value.trim().toUpperCase();
    if (!s || !t) {
      showStatus('status-live', 'error', 'Need both Source and Target ranges (e.g. A1, B1)');
      return;
    }
    liveRules.push({ source: s, target: t, activeShapeName: null });
    sourceIn.value = '';
    targetIn.value = '';
    renderList();
  });

  toggleBtn.addEventListener('click', async () => {
    if (liveRules.length === 0) { showStatus('status-live', 'error', 'Add a rule first.'); return; }

    if (!liveSyncActive) {
      try {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          liveSyncHandler = sheet.onChanged.add(onWorksheetChanged);
          await context.sync();
          
          liveSyncActive = true;
          toggleTxt.textContent = 'Stop Live Engine';
          toggleIcn.textContent = '⏹';
          toggleBtn.style.background = 'linear-gradient(135deg, #f43f5e 0%, #be123c 100%)';
          showStatus('status-live', 'info', 'Engine active. Automatically syncing...');
          
          // Force an initial sync pass to generate everything right now
          await processAllLiveRules();
        });
      } catch (err) { showStatus('status-live', 'error', 'Engine start failed: ' + err.message); }
    } else {
       try {
         await Excel.run(async (context) => {
           if (liveSyncHandler) { liveSyncHandler.remove(); await context.sync(); }
           liveSyncActive = false;
           liveSyncHandler = null;
           toggleTxt.textContent = 'Start Live Engine';
           toggleIcn.textContent = '▶';
           toggleBtn.style.background = '';
           showStatus('status-live', 'info', 'Engine paused.');
         });
       } catch (err) { showStatus('status-live', 'error', 'Engine stop failed: ' + err.message); }
    }
  });

  renderList();
}

async function onWorksheetChanged(eventArgs) {
  if (!liveSyncActive || liveRules.length === 0) return;
  // Trigger generation process safely in background without blocking UI
  setTimeout(() => processAllLiveRules(), 50);
}

async function processAllLiveRules() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    for (let rule of liveRules) {
      try {
        const sourceRange = sheet.getRange(rule.source);
        sourceRange.load(['values']);
        await context.sync();

        let textParts = [];
        for (let r = 0; r < sourceRange.values.length; r++) {
          for (let c = 0; c < sourceRange.values[r].length; c++) {
            const val = sourceRange.values[r][c];
            if (val !== null && val !== undefined && String(val).trim() !== '') {
              textParts.push(String(val).trim());
            }
          }
        }
        
        const textToEncode = textParts.join('');
        if (!textToEncode) continue;

        // Generate barcode base64 using task pane offline engine
        const base64 = await generateDataMatrixBase64(textToEncode, 'auto', 5);

        // Fetch Target Cell dimensions for perfect "In-Cell" bounding box snapping
        const targetRange = sheet.getRange(rule.target);
        targetRange.load(['left', 'top', 'width', 'height']);
        await context.sync();

        // Standardize shape placement
        if (rule.activeShapeName) {
           const oldShape = sheet.shapes.getItemOrNullObject(rule.activeShapeName);
           oldShape.load('name');
           await context.sync();
           if (!oldShape.isNullObject) { oldShape.delete(); }
        }

        const newShape = sheet.shapes.addImage(base64);
        rule.activeShapeName = 'LiveWatcher_' + Date.now();
        newShape.name = rule.activeShapeName;
        
        // Exact Cell Snapping Math (2pt margin to keep cell borders visible)
        newShape.left = targetRange.left + 2;
        newShape.top = targetRange.top + 2;
        
        // Lock to the smallest cell dimension so it doesn't overflow
        const maxDim = Math.min(targetRange.width, targetRange.height) - 4;
        newShape.width = maxDim;
        newShape.height = maxDim;
        
        // Lock Aspect Ratio & Link to Cell Resizing natively
        newShape.lockAspectRatio = true;
        newShape.placement = Excel.Placement.twoCell;
        
      } catch (err) {
        console.warn(`Rule Sync Error for ${rule.source}:`, err.message);
      }
    }
    await context.sync();
  }).catch(err => console.error("Live Sync Runtime Error:", err));
}

