/**
 * 模块视图管理器
 * 负责渲染各功能模块的 UI
 */

import { loadFileToFS, readFileFromFS, runPython, downloadFile, setLogCallback } from './engine.js';

/**
 * 工具函数：格式化文件大小
 */
function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

/**
 * 工具函数：创建文件上传区
 */
function createUploadZone(id, label, accept = '.xlsx,.xls,.csv', multiple = true) {
    return `
    <div class="form-group">
      <label class="form-label">${label}</label>
      <div class="file-upload-zone" id="${id}-zone">
        <div class="upload-icon">📁</div>
        <div class="upload-text">拖放文件到此处，或 <strong>点击选择</strong></div>
        <input type="file" id="${id}-input" accept="${accept}" ${multiple ? 'multiple' : ''} />
      </div>
      <div class="file-list" id="${id}-list"></div>
    </div>
  `;
}

/**
 * 工具函数：绑定文件上传事件
 * @returns {Function} getFiles - 获取已选文件列表
 */
function bindUploadEvents(id) {
    const zone = document.getElementById(`${id}-zone`);
    const input = document.getElementById(`${id}-input`);
    const listEl = document.getElementById(`${id}-list`);
    let selectedFiles = [];

    // 拖放
    zone.addEventListener('dragover', (e) => {
        e.preventDefault();
        zone.classList.add('dragover');
    });
    zone.addEventListener('dragleave', () => {
        zone.classList.remove('dragover');
    });
    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('dragover');
        addFiles(e.dataTransfer.files);
    });

    // 点击选择
    input.addEventListener('change', (e) => {
        addFiles(e.target.files);
        input.value = ''; // 允许重复选同一文件
    });

    function addFiles(fileList) {
        for (const f of fileList) {
            // 避免重复
            if (!selectedFiles.some(sf => sf.name === f.name && sf.size === f.size)) {
                selectedFiles.push(f);
            }
        }
        renderList();
    }

    function renderList() {
        listEl.innerHTML = selectedFiles.map((f, i) => `
      <div class="file-item">
        <span class="file-name">📄 ${f.name}</span>
        <span class="file-size">${formatSize(f.size)}</span>
        <span class="file-remove" data-idx="${i}">✕</span>
      </div>
    `).join('');

        listEl.querySelectorAll('.file-remove').forEach(btn => {
            btn.addEventListener('click', () => {
                selectedFiles.splice(parseInt(btn.dataset.idx), 1);
                renderList();
            });
        });
    }

    return () => selectedFiles;
}

/**
 * 创建结果区
 */
function createResultSection(id) {
    return `<div class="result-section" id="${id}-result">
    <div class="result-title" id="${id}-result-title"></div>
    <div class="result-message" id="${id}-result-msg"></div>
  </div>`;
}

/**
 * 显示结果
 */
function showResult(id, success, message) {
    const el = document.getElementById(`${id}-result`);
    const titleEl = document.getElementById(`${id}-result-title`);
    const msgEl = document.getElementById(`${id}-result-msg`);

    el.className = `result-section active ${success ? 'result-success' : 'result-error'}`;
    titleEl.textContent = success ? '✅ 处理成功' : '❌ 处理失败';
    msgEl.textContent = message;
}

/**
 * 隐藏结果
 */
function hideResult(id) {
    const el = document.getElementById(`${id}-result`);
    if (el) el.className = 'result-section';
}

// ===================== 各模块视图 =====================

/**
 * 资金对账模块
 */
export function renderFinanceAudit(container, logger) {
    container.innerHTML = `
    <div class="module-card fade-in">
      <div class="module-card-title"><span class="card-icon">📊</span> 资金对账引擎</div>
      <p style="color: var(--color-text-secondary); font-size: 0.88rem; margin-bottom: var(--space-lg);">
        核对"计提台账"与"实际账单"的金额与明细差异，自动生成对账报告。
      </p>
      ${createUploadZone('audit-accrual', '计提台账文件（Excel）')}
      ${createUploadZone('audit-actual', '实际账单文件（Excel）')}
      ${createResultSection('audit')}
      <div class="action-bar">
        <button class="btn btn-primary" id="audit-run-btn">
          🚀 开始对账
        </button>
      </div>
    </div>
  `;

    const getAccrualFiles = bindUploadEvents('audit-accrual');
    const getActualFiles = bindUploadEvents('audit-actual');

    document.getElementById('audit-run-btn').addEventListener('click', async () => {
        const accrualFiles = getAccrualFiles();
        const actualFiles = getActualFiles();

        if (accrualFiles.length === 0) {
            logger.warn('请先选择计提台账文件'); return;
        }
        if (actualFiles.length === 0) {
            logger.warn('请先选择实际账单文件'); return;
        }

        const btn = document.getElementById('audit-run-btn');
        btn.disabled = true;
        btn.innerHTML = '<div class="spinner"></div> 处理中...';
        hideResult('audit');

        try {
            // 设置日志回调
            setLogCallback((msg) => logger.info(msg));

            logger.info('正在加载文件到处理引擎...');

            // 创建临时目录
            await runPython(`
import os
os.makedirs('/tmp/accrual', exist_ok=True)
os.makedirs('/tmp/actual', exist_ok=True)
os.makedirs('/tmp/output', exist_ok=True)
      `);

            // 加载计提文件
            for (const f of accrualFiles) {
                await loadFileToFS(f, `/tmp/accrual/${f.name}`);
                logger.info(`已加载计提文件: ${f.name}`);
            }

            // 加载实际文件
            for (const f of actualFiles) {
                await loadFileToFS(f, `/tmp/actual/${f.name}`);
                logger.info(`已加载实际文件: ${f.name}`);
            }

            logger.info('启动资金对账引擎...');

            // 执行对账
            const resultJson = await runPython(`
from engine_audit import run_audit_web
import json
result = run_audit_web('/tmp/accrual', '/tmp/actual', '/tmp/output', _log)
json.dumps(result, ensure_ascii=False)
      `);

            const result = JSON.parse(resultJson);

            if (result.success) {
                logger.success(result.message);
                showResult('audit', true, result.message);

                // 下载结果文件
                if (result.output_file) {
                    const data = readFileFromFS(result.output_file);
                    const filename = result.output_file.split('/').pop();
                    downloadFile(data, filename);
                    logger.success(`已下载: ${filename}`);
                }
            } else {
                logger.error(result.message);
                showResult('audit', false, result.message);
            }
        } catch (err) {
            logger.error(`对账失败: ${err.message}`);
            showResult('audit', false, err.message);
        } finally {
            btn.disabled = false;
            btn.innerHTML = '🚀 开始对账';
        }
    });
}

/**
 * 账单清洗模块
 */
export function renderBillParser(container, logger) {
    container.innerHTML = `
    <div class="module-card fade-in">
      <div class="module-card-title"><span class="card-icon">🔧</span> ERP 账单清洗器</div>
      <p style="color: var(--color-text-secondary); font-size: 0.88rem; margin-bottom: var(--space-lg);">
        将供应商导出的复杂 ERP 格式账单，转换为标准中间格式。
      </p>
      ${createUploadZone('parser-source', '供应商原始账单（Excel）')}
      ${createResultSection('parser')}
      <div class="action-bar">
        <button class="btn btn-primary" id="parser-run-btn">
          🚀 开始清洗
        </button>
      </div>
    </div>
  `;

    const getSourceFiles = bindUploadEvents('parser-source');

    document.getElementById('parser-run-btn').addEventListener('click', async () => {
        const files = getSourceFiles();
        if (files.length === 0) {
            logger.warn('请先选择源文件'); return;
        }

        const btn = document.getElementById('parser-run-btn');
        btn.disabled = true;
        btn.innerHTML = '<div class="spinner"></div> 处理中...';
        hideResult('parser');

        try {
            setLogCallback((msg) => logger.info(msg));

            await runPython(`
import os
os.makedirs('/tmp/source', exist_ok=True)
os.makedirs('/tmp/output', exist_ok=True)
      `);

            for (const f of files) {
                await loadFileToFS(f, `/tmp/source/${f.name}`);
                logger.info(`已加载源文件: ${f.name}`);
            }

            logger.info('启动账单清洗引擎...');

            const resultJson = await runPython(`
from engine_parser import run_parser_web
import json
result = run_parser_web('/tmp/source', '/tmp/output', _log)
json.dumps(result, ensure_ascii=False)
      `);

            const result = JSON.parse(resultJson);

            if (result.success) {
                logger.success(result.message);
                showResult('parser', true, result.message);

                for (const outFile of (result.output_files || [])) {
                    const data = readFileFromFS(outFile);
                    const filename = outFile.split('/').pop();
                    downloadFile(data, filename);
                    logger.success(`已下载: ${filename}`);
                }
            } else {
                logger.error(result.message);
                showResult('parser', false, result.message);
            }
        } catch (err) {
            logger.error(`清洗失败: ${err.message}`);
            showResult('parser', false, err.message);
        } finally {
            btn.disabled = false;
            btn.innerHTML = '🚀 开始清洗';
        }
    });
}

/**
 * 上传模板生成模块
 */
export function renderUploadTemplate(container, logger) {
    container.innerHTML = `
    <div class="module-card fade-in">
      <div class="module-card-title"><span class="card-icon">📝</span> 上传模板生成器</div>
      <p style="color: var(--color-text-secondary); font-size: 0.88rem; margin-bottom: var(--space-lg);">
        将清洗后的账单数据，映射并填入 OMS 系统标准上传模板。
      </p>
      ${createUploadZone('template-source', '数据源文件（清洗后的明细表）')}
      ${createResultSection('template')}
      <div class="action-bar">
        <button class="btn btn-primary" id="template-run-btn">
          🚀 生成模板
        </button>
      </div>
    </div>
  `;

    const getSourceFiles = bindUploadEvents('template-source');

    document.getElementById('template-run-btn').addEventListener('click', async () => {
        const files = getSourceFiles();
        if (files.length === 0) {
            logger.warn('请先选择源文件'); return;
        }

        const btn = document.getElementById('template-run-btn');
        btn.disabled = true;
        btn.innerHTML = '<div class="spinner"></div> 处理中...';
        hideResult('template');

        try {
            setLogCallback((msg) => logger.info(msg));

            await runPython(`
import os
os.makedirs('/tmp/source', exist_ok=True)
os.makedirs('/tmp/output', exist_ok=True)
      `);

            for (const f of files) {
                await loadFileToFS(f, `/tmp/source/${f.name}`);
                logger.info(`已加载: ${f.name}`);
            }

            logger.info('启动模板生成引擎...');

            const resultJson = await runPython(`
from engine_template import run_template_web
import json
result = run_template_web('/tmp/source', '/tmp/output', _log)
json.dumps(result, ensure_ascii=False)
      `);

            const result = JSON.parse(resultJson);

            if (result.success) {
                logger.success(result.message);
                showResult('template', true, result.message);

                for (const outFile of (result.output_files || [])) {
                    const data = readFileFromFS(outFile);
                    const filename = outFile.split('/').pop();
                    downloadFile(data, filename);
                    logger.success(`已下载: ${filename}`);
                }
            } else {
                logger.error(result.message);
                showResult('template', false, result.message);
            }
        } catch (err) {
            logger.error(`生成失败: ${err.message}`);
            showResult('template', false, err.message);
        } finally {
            btn.disabled = false;
            btn.innerHTML = '🚀 生成模板';
        }
    });
}

/**
 * 报价稽核模块
 */
export function renderQuoteAudit(container, logger) {
    container.innerHTML = `
    <div class="module-card fade-in">
      <div class="module-card-title"><span class="card-icon">🔍</span> 报价符合性稽核</div>
      <p style="color: var(--color-text-secondary); font-size: 0.88rem; margin-bottom: var(--space-lg);">
        校验供应商账单中的单价是否符合合同报价。
      </p>
      ${createUploadZone('quote-bill', '账单明细文件（Excel）')}
      ${createResultSection('quote')}
      <div class="action-bar">
        <button class="btn btn-primary" id="quote-run-btn">
          🚀 开始稽核
        </button>
      </div>
    </div>
  `;

    const getBillFiles = bindUploadEvents('quote-bill');

    document.getElementById('quote-run-btn').addEventListener('click', async () => {
        const files = getBillFiles();
        if (files.length === 0) {
            logger.warn('请先选择账单文件'); return;
        }

        const btn = document.getElementById('quote-run-btn');
        btn.disabled = true;
        btn.innerHTML = '<div class="spinner"></div> 处理中...';
        hideResult('quote');

        try {
            setLogCallback((msg) => logger.info(msg));

            await runPython(`
import os
os.makedirs('/tmp/source', exist_ok=True)
os.makedirs('/tmp/output', exist_ok=True)
      `);

            for (const f of files) {
                await loadFileToFS(f, `/tmp/source/${f.name}`);
                logger.info(`已加载: ${f.name}`);
            }

            logger.info('启动报价稽核引擎...');

            const resultJson = await runPython(`
from engine_quote_audit import run_quote_audit_web
import json
result = run_quote_audit_web('/tmp/source', '/tmp/output', _log)
json.dumps(result, ensure_ascii=False)
      `);

            const result = JSON.parse(resultJson);

            if (result.success) {
                logger.success(result.message);
                showResult('quote', true, result.message);

                for (const outFile of (result.output_files || [])) {
                    const data = readFileFromFS(outFile);
                    const filename = outFile.split('/').pop();
                    downloadFile(data, filename);
                    logger.success(`已下载: ${filename}`);
                }
            } else {
                logger.error(result.message);
                showResult('quote', false, result.message);
            }
        } catch (err) {
            logger.error(`稽核失败: ${err.message}`);
            showResult('quote', false, err.message);
        } finally {
            btn.disabled = false;
            btn.innerHTML = '🚀 开始稽核';
        }
    });
}

/**
 * 报价计算模块
 */
export function renderQuoteCalculator(container, logger) {
    container.innerHTML = `
    <div class="module-card fade-in">
      <div class="module-card-title"><span class="card-icon">🧮</span> 物流报价计算</div>
      <p style="color: var(--color-text-secondary); font-size: 0.88rem; margin-bottom: var(--space-lg);">
        根据品牌、目的地、重量、托板数计算物流报价。
      </p>

      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: var(--space-md);">
        <div class="form-group">
          <label class="form-label">品牌</label>
          <select class="form-select" id="calc-brand">
            <option value="">请先加载规则文件</option>
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">目的地仓库</label>
          <select class="form-select" id="calc-dest">
            <option value="">请先加载规则文件</option>
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">重量 (kg)</label>
          <input type="number" class="form-input" id="calc-weight" placeholder="输入重量" min="0" step="0.1" />
        </div>
        <div class="form-group">
          <label class="form-label">托板数</label>
          <input type="number" class="form-input" id="calc-pallets" placeholder="输入托板数" min="0" step="1" />
        </div>
      </div>

      ${createResultSection('calc')}

      <div class="action-bar">
        <button class="btn btn-primary" id="calc-run-btn">
          🧮 计算报价
        </button>
        <button class="btn btn-secondary" id="calc-export-btn" style="display: none;">
          📥 导出结果
        </button>
      </div>
    </div>

    <div class="module-card fade-in" style="margin-top: var(--space-lg);">
      <div class="module-card-title"><span class="card-icon">📋</span> 计算结果</div>
      <div id="calc-result-table" style="color: var(--color-text-secondary); font-size: 0.88rem;">
        请输入参数并点击计算。
      </div>
    </div>
  `;

    // 报价计算是实时计算，不需要文件上传，但需要从服务端拉取规则文件
    document.getElementById('calc-run-btn').addEventListener('click', async () => {
        const brand = document.getElementById('calc-brand').value;
        const dest = document.getElementById('calc-dest').value;
        const weight = parseFloat(document.getElementById('calc-weight').value) || 0;
        const pallets = parseInt(document.getElementById('calc-pallets').value) || 0;

        if (!brand || !dest) {
            logger.warn('请选择品牌和目的地'); return;
        }
        if (weight <= 0) {
            logger.warn('请输入有效的重量'); return;
        }

        const btn = document.getElementById('calc-run-btn');
        btn.disabled = true;
        btn.innerHTML = '<div class="spinner"></div> 计算中...';
        hideResult('calc');

        try {
            setLogCallback((msg) => logger.info(msg));

            const resultJson = await runPython(`
from engine_calculator import run_calculate_web
import json
result = run_calculate_web("${brand}", "${dest}", ${weight}, ${pallets}, _log)
json.dumps(result, ensure_ascii=False)
      `);

            const result = JSON.parse(resultJson);

            if (result.success) {
                logger.success('报价计算完成');
                showResult('calc', true, result.message);

                // 渲染结果表格
                const tableEl = document.getElementById('calc-result-table');
                if (result.details) {
                    let html = '<table style="width: 100%; border-collapse: collapse; margin-top: 8px;">';
                    html += '<tr style="border-bottom: 2px solid var(--color-border);"><th style="text-align: left; padding: 8px;">费用项</th><th style="text-align: right; padding: 8px;">金额</th></tr>';
                    for (const [key, val] of Object.entries(result.details)) {
                        const isTotal = key.includes('合计') || key.includes('总');
                        html += `<tr style="border-bottom: 1px solid var(--color-border-light); ${isTotal ? 'font-weight: 700;' : ''}">
              <td style="padding: 8px;">${key}</td>
              <td style="text-align: right; padding: 8px;">${typeof val === 'number' ? val.toFixed(2) : val}</td>
            </tr>`;
                    }
                    html += '</table>';
                    tableEl.innerHTML = html;
                }

                document.getElementById('calc-export-btn').style.display = '';
            } else {
                logger.error(result.message);
                showResult('calc', false, result.message);
            }
        } catch (err) {
            logger.error(`计算失败: ${err.message}`);
            showResult('calc', false, err.message);
        } finally {
            btn.disabled = false;
            btn.innerHTML = '🧮 计算报价';
        }
    });
}

/**
 * 设置页面
 */
export function renderSettings(container, logger) {
    container.innerHTML = `
    <div class="module-card fade-in">
      <div class="module-card-title"><span class="card-icon">⚙️</span> 系统设置</div>

      <div class="settings-section">
        <div class="settings-section-title">引擎状态</div>
        <div style="display: flex; align-items: center; gap: var(--space-sm); padding: var(--space-md); background: var(--color-success-bg); border-radius: var(--radius-sm);">
          <span style="font-size: 1.2rem;">✅</span>
          <div>
            <div style="font-weight: 600;">Python 引擎运行中</div>
            <div style="font-size: 0.8rem; color: var(--color-text-secondary);">Pyodide + Pandas + Openpyxl 已加载</div>
          </div>
        </div>
      </div>

      <div class="settings-section">
        <div class="settings-section-title">关于</div>
        <p style="color: var(--color-text-secondary); font-size: 0.88rem; line-height: 1.6;">
          国际电商工具箱 Web 版 v1.0<br/>
          基于 Pyodide 技术在浏览器中运行 Python 引擎。<br/>
          所有数据处理均在本地完成，不会上传到服务器。
        </p>
      </div>
    </div>
  `;
}
