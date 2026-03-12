/**
 * 国际电商工具箱 Web 版
 * 严格对齐 .exe 版本的导航与交互
 */

// ===================== 全局状态 =====================
const PYODIDE_VERSION = '0.27.0';
const DEFAULT_PYODIDE_BASE_URLS = [
  '/pyodide/v' + PYODIDE_VERSION + '/full/',
  'https://cdn.jsdelivr.net/pyodide/v' + PYODIDE_VERSION + '/full/',
  'https://fastly.jsdelivr.net/pyodide/v' + PYODIDE_VERSION + '/full/'
];

function getPyodideBaseUrls() {
  var custom = globalThis.__PYODIDE_BASE_URLS__;
  if (Array.isArray(custom) && custom.length) {
    return custom;
  }
  return DEFAULT_PYODIDE_BASE_URLS;
}
let pyodide = null;
let pyodideReady = false;
let pyodideInitPromise = null;
let currentModule = null;
let logger = null;
let rulesDBReady = false;
let rulesDBInitPromise = null;

const NAV_ITEMS = [
  { id: 'home', icon: '🏠', label: '主页' },
  { id: 'finance-audit', icon: '🏦', label: '资金对账' },
  { id: 'bill-parser', icon: '🧹', label: '快递计提' },
  { id: 'quote-calculator', icon: '🧮', label: '报价计算' },
  { id: 'dashboard', icon: '📊', label: '数据看板' },
  { id: 'rule-management', icon: '⚙️', label: '系统设置' },
];

// ===================== 规则数据库 (IndexedDB) =====================
class RulesDB {
  constructor() {
    this.dbName = 'ToolboxRulesDB';
    this.storeName = 'rules';
    this.db = null;
  }
  async init() {
    return new Promise((resolve, reject) => {
      var req = indexedDB.open(this.dbName, 1);
      req.onupgradeneeded = (e) => {
        var db = e.target.result;
        if (!db.objectStoreNames.contains(this.storeName)) {
          db.createObjectStore(this.storeName, { keyPath: 'name' });
        }
      };
      req.onsuccess = (e) => { this.db = e.target.result; resolve(); };
      req.onerror = (e) => reject(e.target.error);
    });
  }
  async saveFile(file) {
    var buffer = await file.arrayBuffer();
    return new Promise((resolve, reject) => {
      var tx = this.db.transaction(this.storeName, 'readwrite');
      tx.objectStore(this.storeName).put({ name: file.name, data: buffer, type: file.type, size: file.size, date: new Date().getTime() });
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
  }
  async getAllFiles() {
    return new Promise((resolve, reject) => {
      var tx = this.db.transaction(this.storeName, 'readonly');
      var req = tx.objectStore(this.storeName).getAll();
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  async deleteFile(name) {
    return new Promise((resolve, reject) => {
      var tx = this.db.transaction(this.storeName, 'readwrite');
      tx.objectStore(this.storeName).delete(name);
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
  }
}
const rulesDB = new RulesDB();

function initRulesDBBackground() {
  if (rulesDBInitPromise) return rulesDBInitPromise;

  rulesDBInitPromise = rulesDB.init().then(function () {
    rulesDBReady = true;
    return true;
  }).catch(function (err) {
    rulesDBInitPromise = null;
    throw err;
  });

  return rulesDBInitPromise;
}

async function ensureRulesDBReady() {
  if (rulesDBReady) return;
  await initRulesDBBackground();
}

// ===================== 日志管理 =====================
class Logger {
  constructor() {
    this.entriesEl = document.getElementById('log-entries');
    this.bodyEl = document.getElementById('log-body');
    this.consoleEl = document.getElementById('log-console');

    document.getElementById('log-clear-btn').addEventListener('click', function (e) {
      e.stopPropagation(); logger.clear();
    });
    document.getElementById('log-toggle-btn').addEventListener('click', function (e) {
      e.stopPropagation(); logger.toggle();
    });
    document.getElementById('log-header').addEventListener('click', function () { logger.toggle(); });
  }

  log(msg, level) {
    var entry = document.createElement('div');
    entry.className = 'log-entry';
    var time = new Date().toLocaleTimeString('zh-CN', { hour12: false });

    // 自动识别颜色（与 .exe 版一致）
    var autoLevel = level || 'info';
    if (!level) {
      if (msg.indexOf('失败') >= 0 || msg.indexOf('错误') >= 0 || msg.indexOf('Error') >= 0) autoLevel = 'error';
      else if (msg.indexOf('成功') >= 0 || msg.indexOf('完成') >= 0 || msg.indexOf('✅') >= 0) autoLevel = 'success';
      else if (msg.indexOf('⚠️') >= 0 || msg.indexOf('警告') >= 0) autoLevel = 'warning';
    }

    var safe = document.createElement('span');
    safe.textContent = msg;
    entry.innerHTML = '<span class="log-time">' + time + '</span><span class="log-msg ' + autoLevel + '">' + safe.innerHTML + '</span>';
    this.entriesEl.appendChild(entry);
    this.bodyEl.scrollTop = this.bodyEl.scrollHeight;
    if (this.consoleEl.classList.contains('collapsed')) {
      this.consoleEl.classList.remove('collapsed');
    }
  }
  info(msg) { this.log(msg, 'info'); }
  success(msg) { this.log(msg, 'success'); }
  warn(msg) { this.log(msg, 'warning'); }
  error(msg) { this.log(msg, 'error'); }
  clear() { this.entriesEl.innerHTML = ''; }
  toggle() {
    var c = this.consoleEl.classList.toggle('collapsed');
    document.getElementById('log-toggle-btn').textContent = c ? '▲' : '▼';
  }
}

// ===================== Pyodide 引擎 =====================
function setEngineStatus(isReady, text) {
  var dot = document.getElementById('engine-status-dot');
  var label = document.getElementById('engine-status-text');
  if (dot) {
    dot.style.background = isReady ? 'var(--color-success)' : 'var(--color-warning)';
    dot.style.boxShadow = isReady ? '0 0 6px var(--color-success)' : '0 0 6px var(--color-warning)';
  }
  if (label && text) label.textContent = text;
}

function loadScript(src, timeoutMs) {
  timeoutMs = timeoutMs || 15000;
  return new Promise(function (resolve, reject) {
    var s = document.createElement('script');
    var done = false;
    var timer = setTimeout(function () {
      if (done) return;
      done = true;
      s.remove();
      reject(new Error('加载超时: ' + src));
    }, timeoutMs);

    s.src = src;
    s.onload = function () {
      if (done) return;
      done = true;
      clearTimeout(timer);
      resolve();
    };
    s.onerror = function () {
      if (done) return;
      done = true;
      clearTimeout(timer);
      s.remove();
      reject(new Error('加载失败: ' + src));
    };

    document.head.appendChild(s);
  });
}

async function setupPyodideWithFallback(onProgress) {
  var urls = getPyodideBaseUrls();
  var errors = [];

  for (var i = 0; i < urls.length; i++) {
    var base = urls[i];
    if (!base.endsWith('/')) base += '/';

    try {
      onProgress && onProgress(10, '加载 Pyodide 运行时... (' + base + ')');
      await loadScript(base + 'pyodide.js');

      onProgress && onProgress(30, '初始化 Python 环境...');
      pyodide = await globalThis.loadPyodide({ indexURL: base });
      logger && logger.success('Pyodide 节点可用: ' + base);
      return base;
    } catch (e) {
      errors.push(base + ' => ' + e.message);
      logger && logger.warn('Pyodide 节点不可用，尝试下一个: ' + base);
      if (globalThis.loadPyodide && pyodide === null) {
        // 保持继续尝试，不中断
      }
    }
  }

  throw new Error('Pyodide 加载失败，已尝试节点: ' + errors.join(' | '));
}

async function initPyodide(onProgress) {
  if (pyodideReady) {
    onProgress && onProgress(100, '就绪');
    return;
  }
  if (pyodideInitPromise) {
    await pyodideInitPromise;
    onProgress && onProgress(100, '就绪');
    return;
  }

  pyodideInitPromise = (async function () {
    await setupPyodideWithFallback(onProgress);

    onProgress && onProgress(50, '安装 Pandas...');
    await pyodide.loadPackage('pandas');

    onProgress && onProgress(65, '安装 micropip...');
    await pyodide.loadPackage('micropip');

    onProgress && onProgress(75, '安装 Openpyxl...');
    await pyodide.runPythonAsync('import micropip\nawait micropip.install("openpyxl")');

    onProgress && onProgress(80, '拉取环境业务逻辑...');
    var scripts = ['auto_audit.py', 'bill_parser.py', 'quote_calculator.py'];
    for (var i = 0; i < scripts.length; i++) {
      try {
        var r = await fetch('../' + scripts[i]);
        if (!r.ok) r = await fetch(scripts[i]);
        if (!r.ok) throw new Error('Not found');
        var code = await r.text();
        pyodide.FS.writeFile(scripts[i], code);
      } catch (e) {
        console.warn('Failed to load ' + scripts[i], e);
      }
    }

    onProgress && onProgress(90, '注册处理引擎...');
    await registerEngines();

    onProgress && onProgress(100, '就绪');
    pyodideReady = true;
    setEngineStatus(true, '引擎就绪');
  })();

  try {
    await pyodideInitPromise;
  } finally {
    if (pyodideReady) {
      pyodideInitPromise = null;
    }
  }
}

async function registerEngines() {
  await pyodide.runPythonAsync(ENGINE_AUDIT_CODE);
  await pyodide.runPythonAsync(ENGINE_PARSER_CODE);
  await pyodide.runPythonAsync(ENGINE_CALCULATOR_CODE);
}

function setLogCallback(fn) {
  if (pyodide) pyodide.globals.set('_js_log', fn);
}

async function loadFileToVFS(file, path) {
  var buf = await file.arrayBuffer();
  pyodide.FS.writeFile(path, new Uint8Array(buf));
}

function readFileFromVFS(path) {
  return pyodide.FS.readFile(path);
}

async function ensurePyodideReady(taskName) {
  if (pyodideReady) return;

  logger && logger.info('首次使用' + taskName + '，正在初始化引擎，请稍候...');
  setEngineStatus(false, '引擎初始化中...');

  await initPyodide(function (pct, msg) {
    if (!logger) return;
    if (pct === 100 || pct % 20 === 0) {
      logger.info('[引擎初始化 ' + pct + '%] ' + msg);
    }
  });

  logger && logger.success('引擎初始化完成，可开始处理。');
}

function downloadBlob(data, filename) {
  var blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

async function ensureWorkspaceDirs() {
  await pyodide.runPythonAsync(
    "import os\nfor d in ['/workspace/accrual','/workspace/actual','/workspace/source','/workspace/output','/workspace/rules','/workspace/规则文件','/workspace/输出汇总文件夹']:\n    os.makedirs(d, exist_ok=True)"
  );
}

async function syncRulesToWorkspace() {
  await ensureRulesDBReady();
  await ensureWorkspaceDirs();
  await pyodide.runPythonAsync(
    "import os, glob\nfor d in ['/workspace/rules','/workspace/规则文件']:\n    os.makedirs(d, exist_ok=True)\n    for p in glob.glob(os.path.join(d, '*')):\n        if os.path.isfile(p):\n            os.remove(p)"
  );
  var allRules = await rulesDB.getAllFiles();
  for (var i = 0; i < allRules.length; i++) {
    var bytes = new Uint8Array(allRules[i].data);
    // 兼容旧路径，同时满足原始引擎固定读取的“规则文件”目录
    pyodide.FS.writeFile('/workspace/rules/' + allRules[i].name, bytes);
    pyodide.FS.writeFile('/workspace/规则文件/' + allRules[i].name, bytes);
  }
  return allRules.length;
}

// ===================== 通用 UI 工具 =====================
function formatSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / 1048576).toFixed(1) + ' MB';
}

/**
 * 创建文件选择区域（对齐 .exe 版的"选择文件/文件夹"样式）
 * @param {string} id - 唯一 ID
 * @param {string} label - 步骤标签
 * @param {string} btnText - 按钮文字
 * @param {boolean} multiple - 是否支持多文件
 * @param {string} accept - 接受的文件类型
 */
function createFilePickRow(id, label, btnText, multiple, accept) {
  accept = accept || '.xlsx,.xls';
  btnText = btnText || '选择文件';
  return '<div class="step-row">' +
    '<div class="step-label">' + label + '</div>' +
    '<div class="file-pick-row">' +
    '<label class="btn btn-outline" for="' + id + '-input">' +
    '<span class="btn-icon">📁</span> ' + btnText +
    '</label>' +
    '<input type="file" id="' + id + '-input" accept="' + accept + '"' + (multiple ? ' multiple' : '') + ' style="display:none;" />' +
    '<div class="file-status" id="' + id + '-status">未选择文件</div>' +
    '</div>' +
    '<div class="file-list" id="' + id + '-list"></div>' +
    '</div>';
}

/**
 * 创建拖放上传区域（多文件场景）
 */
function createUploadZone(id, label, accept, multiple) {
  accept = accept || '.xlsx,.xls,.csv';
  multiple = multiple !== false;
  return '<div class="form-group"><label class="form-label">' + label + '</label>' +
    '<div class="file-upload-zone" id="' + id + '-zone">' +
    '<div class="upload-icon">📁</div>' +
    '<div class="upload-text">拖放文件到此处，或 <strong>点击选择</strong></div>' +
    '<input type="file" id="' + id + '-input" accept="' + accept + '"' + (multiple ? ' multiple' : '') + ' />' +
    '</div><div class="file-list" id="' + id + '-list"></div></div>';
}

/**
 * 绑定文件选择事件（单文件模式，显示状态文本）
 */
function bindFilePickEvents(id) {
  var input = document.getElementById(id + '-input');
  var statusEl = document.getElementById(id + '-status');
  var listEl = document.getElementById(id + '-list');
  var files = [];

  input.addEventListener('change', function () {
    files = Array.from(input.files);
    if (files.length === 1) {
      statusEl.textContent = files[0].name;
      statusEl.className = 'file-status file-status-ok';
    } else if (files.length > 1) {
      statusEl.textContent = '已选择: ' + files.length + ' 个文件';
      statusEl.className = 'file-status file-status-ok';
    }
    // 显示文件列表
    listEl.innerHTML = files.map(function (f, i) {
      return '<div class="file-item"><span class="file-name">📄 ' + f.name + '</span>' +
        '<span class="file-size">' + formatSize(f.size) + '</span></div>';
    }).join('');
  });

  return function () { return files; };
}

/**
 * 绑定拖放上传事件（多文件模式，带删除）
 */
function bindUploadEvents(id) {
  var zone = document.getElementById(id + '-zone');
  var input = document.getElementById(id + '-input');
  var listEl = document.getElementById(id + '-list');
  var files = [];

  zone.addEventListener('dragover', function (e) { e.preventDefault(); zone.classList.add('dragover'); });
  zone.addEventListener('dragleave', function () { zone.classList.remove('dragover'); });
  zone.addEventListener('drop', function (e) { e.preventDefault(); zone.classList.remove('dragover'); addFiles(e.dataTransfer.files); });
  input.addEventListener('change', function () { addFiles(input.files); input.value = ''; });

  function addFiles(fl) {
    for (var i = 0; i < fl.length; i++) {
      var f = fl[i];
      if (!files.some(function (sf) { return sf.name === f.name && sf.size === f.size; })) {
        files.push(f);
      }
    }
    render();
  }

  function render() {
    listEl.innerHTML = files.map(function (f, i) {
      return '<div class="file-item"><span class="file-name">📄 ' + f.name + '</span>' +
        '<span class="file-size">' + formatSize(f.size) + '</span>' +
        '<span class="file-remove" data-idx="' + i + '">✕</span></div>';
    }).join('');
    listEl.querySelectorAll('.file-remove').forEach(function (btn) {
      btn.addEventListener('click', function () {
        files.splice(parseInt(btn.dataset.idx), 1);
        render();
      });
    });
  }

  return {
    getFiles: function () { return files; },
    clear: function () {
      files = [];
      render();
    }
  };
}

function showResult(id, ok, msg) {
  var el = document.getElementById(id + '-result');
  var title = document.getElementById(id + '-result-title');
  var msgEl = document.getElementById(id + '-result-msg');
  el.className = 'result-section active ' + (ok ? 'result-success' : 'result-error');
  title.textContent = ok ? '✅ 处理成功' : '❌ 处理失败';
  msgEl.innerHTML = msg;
}

function hideResult(id) {
  var el = document.getElementById(id + '-result');
  if (el) el.className = 'result-section';
}

function resultHTML(id) {
  return '<div class="result-section" id="' + id + '-result">' +
    '<div class="result-title" id="' + id + '-result-title"></div>' +
    '<div class="result-message" id="' + id + '-result-msg"></div></div>';
}

// ===================== 通用处理流程 =====================
async function runModule(opts) {
  var btn = document.getElementById(opts.btnId);
  var origLabel = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = '<div class="spinner"></div> 处理中...';
  hideResult(opts.resultId);

  try {
    await ensurePyodideReady('该模块');
    setLogCallback(function (msg) { logger.log(msg); });

    // 准备虚拟目录
    await ensureWorkspaceDirs();
    await pyodide.runPythonAsync(
      "import os, glob\nfor d in ['/workspace/accrual','/workspace/actual','/workspace/source','/workspace/output','/workspace/输出汇总文件夹']:\n    os.makedirs(d, exist_ok=True)\n    for p in glob.glob(os.path.join(d, '*')):\n        if os.path.isfile(p):\n            os.remove(p)"
    );

    // 加载文件到虚拟文件系统
    var fileSets = opts.getFileSets();
    for (var i = 0; i < fileSets.length; i++) {
      var fs = fileSets[i];
      for (var j = 0; j < fs.files.length; j++) {
        var f = fs.files[j];
        await loadFileToVFS(f, fs.dir + '/' + f.name);
        logger.info('已加载: ' + f.name);
      }
    }

    // 同步 IndexedDB 中的规则文件到虚拟文件系统
    await syncRulesToWorkspace();

    // 执行引擎
    logger.info('>>> 启动处理引擎');
    var resultJson = await pyodide.runPythonAsync(opts.engineCall);
    var result = JSON.parse(resultJson);

    if (result.success) {
      logger.success(result.message);
      showResult(opts.resultId, true, result.message);

      // 下载输出文件
      var outFiles = result.output_files || (result.output_file ? [result.output_file] : []);
      for (var k = 0; k < outFiles.length; k++) {
        var data = readFileFromVFS(outFiles[k]);
        var fname = outFiles[k].split('/').pop();
        downloadBlob(data, fname);
        logger.success('已下载: ' + fname);
      }
    } else {
      logger.error(result.message);
      showResult(opts.resultId, false, result.message);
    }
  } catch (err) {
    logger.error('处理失败: ' + err.message);
    showResult(opts.resultId, false, err.message);
  } finally {
    btn.disabled = false;
    btn.innerHTML = origLabel;
  }
}

// ===============================================================
//                        模块视图
// 严格对齐 main_app.py 中的导航与 UI 结构
// ===============================================================

// ---- 主页（.exe 版：欢迎使用 + 手势图标） ----
function renderHome(c) {
  c.innerHTML = '<div class="module-card fade-in welcome-card">' +
    '<div class="welcome-content">' +
    '<div class="welcome-icon">👋</div>' +
    '<h2 class="welcome-title">欢迎使用</h2>' +
    '<p class="welcome-subtitle">请从左侧导航栏选择您需要处理的业务模块</p>' +
    '</div></div>';
}

// ---- 资金对账（.exe 版：3 步操作） ----
function renderFinanceAudit(c) {
  c.innerHTML = '<div class="module-card fade-in">' +
    '<div class="module-card-title"><span class="card-icon" style="color:#d32f2f;">🏦</span> 资金对账</div>' +
    '<div class="module-divider"></div>' +

    // 第1步：选择计提台账文件夹
    createFilePickRow('audit-accrual', '第1步：选择计提台账文件夹', '选择文件夹', true) +

    '<div class="step-spacer"></div>' +

    // 第2步：选择实际账单文件夹
    createFilePickRow('audit-actual', '第2步：选择实际账单文件夹', '选择文件夹', true) +

    '<div class="step-spacer"></div>' +

    '<div class="step-row"><div class="step-label">第3步：仓库映射表（可选）</div><div style="font-size:0.85rem;color:var(--color-text-secondary);">运行对账时，将自动从系统设置中读取已上传的映射表规则文件</div></div>' +

    resultHTML('audit') +
    '<div class="action-bar"><button class="btn btn-primary" id="audit-run-btn">▶ 开始自动对账</button></div></div>';

  var getAccrual = bindFilePickEvents('audit-accrual');
  var getActual = bindFilePickEvents('audit-actual');

  document.getElementById('audit-run-btn').addEventListener('click', function () {
    if (getAccrual().length === 0) { logger.warn('请先选择计提台账文件'); return; }
    if (getActual().length === 0) { logger.warn('请先选择实际账单文件'); return; }
    runModule({
      btnId: 'audit-run-btn', resultId: 'audit',
      getFileSets: function () {
        var sets = [
          { dir: '/workspace/accrual', files: getAccrual() },
          { dir: '/workspace/actual', files: getActual() }
        ];
        return sets;
      },
      engineCall: '_run_audit("/workspace/accrual", "/workspace/actual", "/workspace/output", _js_log)'
    });
  });
}

// ---- 快递计提（.exe 版：蓝色提示框 + 单文件） ----
function renderBillParser(c) {
  c.innerHTML = '<div class="module-card fade-in">' +
    '<div class="module-card-title"><span class="card-icon" style="color:#4caf50;">🧹</span> 快递计提</div>' +
    '<div class="module-divider"></div>' +

    // 蓝色提示框（与 .exe 版一致）
    '<div class="info-banner">' +
    '<span class="info-icon">ℹ️</span>' +
    '<span>根据菜鸟费用计提规则，将金掌柜后台费用分类并调整为快递计提台账上传模板格式</span>' +
    '</div>' +

    '<div class="step-spacer"></div>' +

    // 选择供应商原始账单
    createFilePickRow('parser-source', '选择供应商原始账单（Excel）', '选择文件', false) +

    resultHTML('parser') +
    '<div class="action-bar"><button class="btn btn-success" id="parser-run-btn">▶ 开始清洗账单</button></div></div>';

  var getFiles = bindFilePickEvents('parser-source');

  document.getElementById('parser-run-btn').addEventListener('click', function () {
    if (getFiles().length === 0) { logger.warn('请先选择源文件'); return; }
    runModule({
      btnId: 'parser-run-btn', resultId: 'parser',
      getFileSets: function () { return [{ dir: '/workspace/source', files: getFiles() }]; },
      engineCall: '_run_parser("/workspace/source", "/workspace/output", _js_log)'
    });
  });
}

// ---- 报价计算（.exe 版：加载规则+品牌搜索+目的地+重量/托板+结果卡片网格） ----
function renderQuoteCalculator(c) {
  c.innerHTML = '<div class="module-card fade-in">' +
    // 标题栏
    '<div class="module-card-title" style="justify-content:space-between;">' +
    '<div style="display:flex;align-items:center;gap:var(--space-sm);"><span class="card-icon" style="color:#1976d2;">🧮</span> 报价计算</div>' +
    '</div>' +

    // 规则文件状态
    '<div class="rules-status" id="calc-rules-status">' +
    '<span class="info-icon-sm">ℹ️</span> <span id="calc-rules-text">计算引擎运行时将自动从系统设置中读取已上传的规则文件</span>' +
    '</div>' +

    '<div class="module-divider"></div>' +

    // 输入区域 Row 1：品牌搜索 + 品牌选择 + 目的地
    '<div class="form-row-3">' +
    '<div class="form-group"><label class="form-label">搜索品牌</label><input type="text" class="form-input" id="calc-brand-search" placeholder="输入关键字" disabled /></div>' +
    '<div class="form-group"><label class="form-label">选择品牌</label><select class="form-select" id="calc-brand" disabled><option value="">规则加载中...</option></select></div>' +
    '<div class="form-group"><label class="form-label">目的地仓库</label><select class="form-select" id="calc-dest" disabled><option value="">规则加载中...</option></select></div>' +
    '<div class="form-row-action">' +
    '<div class="form-group" style="width:150px;"><label class="form-label">重量 (KG)</label><input type="number" class="form-input" id="calc-weight" placeholder="输入重量" min="0" step="0.1" /></div>' +
    '<div class="form-group" style="width:150px;"><label class="form-label">托板数</label><input type="number" class="form-input" id="calc-pallets" placeholder="输入托板数" min="0" step="1" /></div>' +
    '<div style="flex:1;"></div>' +
    '<button class="btn btn-primary-blue" id="calc-run-btn">🧮 计算报价</button>' +
    '</div>' +
    '</div>' +

    // 最优方案提示
    '<div class="best-tip" id="calc-best-tip" style="display:none;"><span>✅</span> <span id="calc-best-text"></span></div>' +

    // 结果卡片网格
    '<div class="result-grid" id="calc-result-grid"></div>' +

    resultHTML('calc') +
    '</div>';

  setTimeout(loadQuoteOptions, 10);

  async function loadQuoteOptions() {
    var brandSel = document.getElementById('calc-brand');
    var destSel = document.getElementById('calc-dest');
    var searchInp = document.getElementById('calc-brand-search');
    try {
      // 检查规则文件是否已上传
      await ensureRulesDBReady();
      var rules = await rulesDB.getAllFiles();
      var hasQuoteRule = rules.some(f => f.name.includes("空运报价费用规则"));
      if (!hasQuoteRule) {
        brandSel.innerHTML = '<option value="">请在"系统设置"上传报价规则文件</option>';
        destSel.innerHTML = '<option value="">请在"系统设置"上传报价规则文件</option>';
        document.getElementById('calc-rules-text').textContent = '⚠️ 缺失报价规则文件，请前往"系统设置"上传《空运报价费用规则.xlsx》';
        document.getElementById('calc-rules-text').style.color = '#ef4444';
        return;
      }

      await ensurePyodideReady('报价计算');

      // 同步规则文件到工作区
      await syncRulesToWorkspace();

      // 确保日志回调已注入 Pyodide 全局变量
      setLogCallback(function (msg) { logger.log(msg); });
      var resStr = await pyodide.runPythonAsync('_get_quote_options(_js_log)');
      var res = JSON.parse(resStr);
      if (res.success) {
        window._quoteBrands = res.brands;
        window._quoteDests = res.destinations;
        renderBrandsOptions(res.brands);
        destSel.innerHTML = res.destinations.map(d => `<option value="${d}">${d}</option>`).join('');
        destSel.disabled = false;
        searchInp.disabled = false;
        document.getElementById('calc-rules-text').textContent = '✅ 已成功加载规则，可开始计算';
        document.getElementById('calc-rules-text').style.color = '#10b981';
      } else {
        brandSel.innerHTML = `<option value="">加载失败: ${res.message}</option>`;
        destSel.innerHTML = `<option value="">加载失败: ${res.message}</option>`;
      }
    } catch (e) {
      brandSel.innerHTML = `<option value="">引擎错误: ${e.message}</option>`;
      destSel.innerHTML = `<option value="">引擎错误: ${e.message}</option>`;
    }
  }

  function renderBrandsOptions(brands) {
    var brandSel = document.getElementById('calc-brand');
    if (!brandSel) return; // DOM 已被替换，跳过
    var val = brandSel.value; // 保留原选择
    brandSel.innerHTML = brands.map(b => `<option value="${b}">${b}</option>`).join('');
    brandSel.disabled = false;
    if (brands.includes(val)) brandSel.value = val;
  }

  // 品牌搜索过滤
  document.getElementById('calc-brand-search').addEventListener('input', function (e) {
    if (!window._quoteBrands) return;
    var kw = e.target.value.trim().toLowerCase();
    if (!kw) {
      renderBrandsOptions(window._quoteBrands);
    } else {
      var filtered = window._quoteBrands.filter(b => b.toLowerCase().includes(kw));
      renderBrandsOptions(filtered);
    }
  });

  document.getElementById('calc-run-btn').addEventListener('click', async function () {
    var brand = document.getElementById('calc-brand').value;
    var dest = document.getElementById('calc-dest').value;
    var weight = document.getElementById('calc-weight').value;
    var pallets = document.getElementById('calc-pallets').value;

    if (!brand || !dest) { logger.warn('请完善品牌与目的地信息'); return; }
    if (!weight || !pallets) { logger.warn('请输入重量和托板数'); return; }

    var container = document.getElementById('calc-result');
    container.classList.remove('active');

    // 执行计算引擎
    logger.info('>>> 启动报价计算引擎');
    try {
      var pyCode = `
import json
res = _run_calculate("${brand}", "${dest}", ${weight}, ${pallets}, "/workspace/output", _js_log)
res
`;
      var resStr = await pyodide.runPythonAsync(pyCode);
      var res = JSON.parse(resStr);
      if (res.success) {
        logger.success(res.message);
        // 渲染网格结果
        var scenarios = res.results.scenarios;
        var min_scen = res.results.min_scenario.name;
        var max_scen = res.results.max_scenario.name;

        // 将 LTL/FTL 转换为中文
        function toCN(name) { return name.replace(/LTL/g, '零担').replace(/FTL/g, '整车'); }

        var gridHtml = '<div class="best-tip"><span>💡</span> 最优方案：' + toCN(min_scen) + '（¥' + res.results.min_scenario.total.toFixed(2) + '）</div>';
        gridHtml += '<div class="result-grid">';

        for (var i = 0; i < scenarios.length; i++) {
          var s = scenarios[i];
          var cssClass = 'result-card';
          if (s.name === min_scen) cssClass += ' best';
          if (s.name === max_scen) cssClass += ' worst';

          gridHtml += `
                <div class="${cssClass}">
                    <div class="result-card-title">${toCN(s.name)}</div>
                    <div class="result-card-price">¥${s.total.toFixed(2)}</div>
                    <div class="result-card-detail"><span class="label">起运国提货费</span><span class="value">¥${s.origin.toFixed(2)}</span></div>
                    <div class="result-card-detail"><span class="label">空运费</span><span class="value">¥${s.air.toFixed(2)}</span></div>
                    <div class="result-card-detail"><span class="label">目的港费用</span><span class="value">¥${s.dest_port.toFixed(2)}</span></div>
                    <div class="result-card-detail"><span class="label">港到仓费用</span><span class="value">¥${s.dest_wh.toFixed(2)}</span></div>
                </div>`;
        }
        gridHtml += '</div>';

        showResult('calc', true, gridHtml);


      } else {
        logger.error(res.message);
        showResult('calc', false, res.message);
      }
    } catch (e) {
      logger.error('严重错误: ' + e.message);
      showResult('calc', false, '严重错误: ' + e.message);
    }
  });
}

// ---- 数据看板（iframe 嵌入独立看板页面） ----
function renderDashboard(c) {
  c.innerHTML = '<div class="dashboard-frame fade-in">' +
    '<iframe src="../国际电商看板展示V10.html" ' +
    'style="width:100%;height:calc(100vh - 60px);border:none;border-radius:8px;">' +
    '</iframe></div>';
}

// ---- 系统设置 / 规则管理 ----
function renderRuleManagement(c) {
  c.innerHTML = '<div class="module-card fade-in">' +
    '<div class="module-card-title"><span class="card-icon" style="color:#607d8b;">⚙️</span> 规则管理</div>' +
    '<div class="module-divider"></div>' +
    '<div class="info-banner" style="margin-bottom:15px;">' +
    '<span class="info-icon">ℹ️</span>' +
    '<span>上传的规则文件（如 .xlsx, .csv）自动存入内部存储，各模块自动加载。<br/>重新上传同名文件将覆盖旧版本。</span>' +
    '</div>' +
    createUploadZone('rules-upload', '上传新的规则文件', '.xlsx,.xls,.csv', true) +
    '<div class="action-bar"><button class="btn btn-primary" id="save-rules-btn">💾 保存并应用规则</button></div>' +
    '<div class="module-divider"></div>' +
    '<h4>已保存的规则文件</h4>' +
    '<div id="rules-saved-list" class="file-list" style="margin-top:10px;">加载中...</div>' +
    '</div>';

  // 修复因为 innerHTML 导致的事件监听防重或丢失
  setTimeout(function () {
    var uploadCtrl = bindUploadEvents('rules-upload');

    async function refreshSavedRules() {
      var listEl = document.getElementById('rules-saved-list');
      if (!listEl) return;
      try {
        await ensureRulesDBReady();
        var files = await rulesDB.getAllFiles();
        if (files.length === 0) {
          listEl.innerHTML = '<div style="color:#999;font-size:0.9rem;">暂无保存的规则文件</div>';
          return;
        }
        listEl.innerHTML = files.map(function (f) {
          var d = new Date(f.date).toLocaleString('zh-CN');
          return '<div class="file-item"><span class="file-name">📄 ' + f.name + '</span>' +
            '<span class="file-size" style="margin-left:auto;margin-right:15px;color:#888;">' + d + '</span>' +
            '<span class="file-size" style="margin-right:15px;">' + formatSize(f.size) + '</span>' +
            '<span class="file-remove" onclick="deleteRule(\'' + f.name + '\')">✕</span></div>';
        }).join('');
      } catch (e) {
        listEl.innerHTML = '<div style="color:red;">读取规则失败</div>';
      }
    }

    // 将全局删除函数暴露到 window 以便 onclick 调用
    window.deleteRule = async function (name) {
      if (confirm("确定删除规则 " + name + " 吗？")) {
        await ensureRulesDBReady();
        await rulesDB.deleteFile(name);
        logger.info("已删除规则: " + name);
        refreshSavedRules();
      }
    };

    var saveBtn = document.getElementById('save-rules-btn');
    if (saveBtn) {
      // 防止重复绑定
      var newBtn = saveBtn.cloneNode(true);
      saveBtn.parentNode.replaceChild(newBtn, saveBtn);

      newBtn.addEventListener('click', async function () {
        var files = uploadCtrl.getFiles();
        if (files.length === 0) { logger.warn("请先选择要上传的规则文件"); return; }
        await ensureRulesDBReady();
        for (var i = 0; i < files.length; i++) {
          await rulesDB.saveFile(files[i]);
          logger.success("已保存规则: " + files[i].name);
        }
        uploadCtrl.clear();
        refreshSavedRules();
      });
    }

    refreshSavedRules();
  }, 0);
}

// 模块渲染映射
var MODULE_RENDERERS = {
  'home': renderHome,
  'finance-audit': renderFinanceAudit,
  'bill-parser': renderBillParser,
  'quote-calculator': renderQuoteCalculator,
  'dashboard': renderDashboard,
  'rule-management': renderRuleManagement
};

function buildNav() {
  var nav = document.getElementById('sidebar-nav');
  nav.innerHTML = NAV_ITEMS.map(function (it) {
    return '<div class="nav-item" data-module="' + it.id + '">' +
      '<span class="nav-icon">' + it.icon + '</span><span>' + it.label + '</span></div>';
  }).join('');

  nav.addEventListener('click', function (e) {
    var item = e.target.closest('.nav-item');
    if (item) switchModule(item.dataset.module);
  });
}

function switchModule(id) {
  if (currentModule === id) return;
  currentModule = id;

  var item = NAV_ITEMS.find(function (n) { return n.id === id; });
  if (!item) return;

  document.querySelectorAll('.nav-item').forEach(function (el) {
    el.classList.toggle('active', el.dataset.module === id);
  });

  // 更新顶栏当前模块名
  var moduleEl = document.getElementById('topbar-module');
  if (moduleEl) moduleEl.textContent = item.icon + ' ' + item.label;

  var container = document.getElementById('content-area');
  var renderer = MODULE_RENDERERS[id];
  if (renderer) renderer(container);
}

// ===================== Python 引擎适配器 =====================
// 注意：这些是简化版适配器，P1 阶段会替换为原始引擎代码

var ENGINE_AUDIT_CODE = `
import json, traceback, glob, os
from auto_audit import FinanceAuditEngine

def _run_audit(acc_dir, act_dir, out_dir, log_fn):
    try:
        os.makedirs(out_dir, exist_ok=True)
        engine = FinanceAuditEngine(workspace_root="/workspace")
        engine.set_custom_dirs(acc_dir, act_dir)
        engine.output_dir = out_dir
        success, msg = engine.run_audit(map_path=None, log_callback=log_fn)
        
        out_files = glob.glob(os.path.join(out_dir, "*.xlsx"))
        latest_out = max(out_files, key=os.path.getctime) if out_files else None
        
        return json.dumps({"success": success, "message": msg, "output_file": latest_out}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"success": False, "message": f"执行失败: {str(e)}\\n{traceback.format_exc()}"}, ensure_ascii=False)
`;

var ENGINE_PARSER_CODE = `
import json, traceback, glob, os
from bill_parser import BillParserEngine

def _run_parser(source_dir, output_dir, log_fn):
    try:
        os.makedirs(output_dir, exist_ok=True)
        engine = BillParserEngine(workspace_root="/workspace")
        engine.output_dir = output_dir
        files = glob.glob(os.path.join(source_dir, "*.xlsx")) + glob.glob(os.path.join(source_dir, "*.xls"))
        if not files:
            return json.dumps({"success": False, "message": "未找到可处理的Excel源文件"}, ensure_ascii=False)

        # UI 只允许单文件，若用户重复操作导致存在多个文件，默认取最新上传文件
        source_file = max(files, key=os.path.getmtime)
        success, msg = engine.run_parser(source_file, log_callback=log_fn)
        
        out_files = glob.glob(os.path.join(engine.output_dir, "*.xlsx"))
        latest_out = max(out_files, key=os.path.getctime) if out_files and success else None
        
        return json.dumps({"success": success, "message": msg, "output_file": latest_out}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"success": False, "message": f"执行失败: {str(e)}\\n{traceback.format_exc()}"}, ensure_ascii=False)
`;

var ENGINE_CALCULATOR_CODE = `
import json, traceback, os
from quote_calculator import QuoteCalculatorEngine

_quote_engine = None

def _get_quote_options(log_fn):
    global _quote_engine
    try:
        if not _quote_engine:
            _quote_engine = QuoteCalculatorEngine(workspace_root="/workspace")
            # 兼容规则文件名
            rf = os.path.join("/workspace/规则文件", "空运报价费用规则.xlsx")
            if not os.path.exists(rf):
                rf2 = os.path.join("/workspace/规则文件", "price_database.xlsx")
                if os.path.exists(rf2): _quote_engine.rules_file = rf2
            
            success, msg = _quote_engine.load_rules(log_callback=log_fn)
            if not success:
                return json.dumps({"success": False, "message": msg}, ensure_ascii=False)
                
        brands = _quote_engine.get_brands()
        dests = _quote_engine.get_destinations()
        return json.dumps({"success": True, "brands": brands, "destinations": dests}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"success": False, "message": str(e)}, ensure_ascii=False)

def _run_calculate(brand, dest, weight, pallets, output_dir, log_fn):
    global _quote_engine
    try:
        if not _quote_engine:
            return json.dumps({"success": False, "message": "引擎未初始化"}, ensure_ascii=False)
            
        success, msg, results = _quote_engine.calculate(brand, dest, float(weight), int(pallets), log_callback=log_fn)
        if not success:
            return json.dumps({"success": False, "message": msg}, ensure_ascii=False)
        
        _quote_engine.output_dir = output_dir
        export_success, export_path = _quote_engine.export_results(results, log_callback=log_fn)
        
        return json.dumps({
            "success": True, 
            "message": "计算完成",
            "results": results, 
            "output_file": export_path if export_success else None
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"success": False, "message": f"计算失败: {str(e)}\\n{traceback.format_exc()}"}, ensure_ascii=False)
`;

// ===================== 启动 =====================
document.addEventListener('DOMContentLoaded', async function () {
  logger = new Logger();

  var loadingBar = document.getElementById('loading-bar');
  var loadingStatus = document.getElementById('loading-status');
  var loadingScreen = document.getElementById('loading-screen');
  var mainApp = document.getElementById('main-app');

  try {
    loadingBar.style.width = '100%';
    loadingStatus.textContent = '界面资源加载完成，Python 引擎将在首次使用时初始化';

    loadingScreen.classList.add('hide');
    mainApp.classList.remove('hidden');
    setTimeout(function () { mainApp.classList.add('visible'); }, 50);

    buildNav();

    // 默认显示主页（与 .exe 一致）
    switchModule('home');

    initRulesDBBackground().then(function () {
      logger.info('本地规则存储已就绪');
    }).catch(function (err) {
      logger.warn('规则存储初始化失败，部分功能可能受限: ' + err.message);
    });

    logger.info('页面加载完成');
    logger.info('首页已优先渲染，规则存储在后台初始化');
    logger.info('已启用按需初始化，首次执行模块时再加载 Python 引擎');
    logger.success('系统就绪，可以开始使用');

  } catch (err) {
    loadingStatus.textContent = '初始化失败: ' + err.message;
    loadingBar.style.background = '#ef4444';
    console.error('Init error:', err);
  }
});
