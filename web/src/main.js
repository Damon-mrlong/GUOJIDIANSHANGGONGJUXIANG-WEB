/**
 * 国际电商工具箱 Web 版 - 主入口
 */

import { initPyodide, runPython, getPyodide } from './engine.js';
import { getLogger } from './logger.js';
import {
    renderFinanceAudit,
    renderBillParser,
    renderUploadTemplate,
    renderQuoteAudit,
    renderQuoteCalculator,
    renderSettings
} from './modules.js';

// ============ 导航配置 ============
const NAV_ITEMS = [
    { id: 'finance-audit', icon: '📊', label: '资金对账', render: renderFinanceAudit },
    { id: 'bill-parser', icon: '🔧', label: '账单清洗', render: renderBillParser },
    { id: 'upload-template', icon: '📝', label: '上传模板', render: renderUploadTemplate },
    { id: 'quote-audit', icon: '🔍', label: '报价稽核', render: renderQuoteAudit },
    { id: 'quote-calculator', icon: '🧮', label: '报价计算', render: renderQuoteCalculator },
    { id: 'settings', icon: '⚙️', label: '系统设置', render: renderSettings },
];

let currentModule = null;

// ============ 初始化应用 ============
async function init() {
    const loadingBar = document.getElementById('loading-bar');
    const loadingStatus = document.getElementById('loading-status');
    const loadingScreen = document.getElementById('loading-screen');
    const mainApp = document.getElementById('main-app');

    try {
        // 加载 Pyodide
        await initPyodide((percent, msg) => {
            loadingBar.style.width = `${percent}%`;
            loadingStatus.textContent = msg;
        });

        // 注册 Python 引擎适配器
        await registerPythonEngines();

        // 隐藏加载屏幕
        loadingScreen.classList.add('hide');
        mainApp.classList.remove('hidden');
        setTimeout(() => mainApp.classList.add('visible'), 50);

        // 构建导航
        buildNav();

        // 初始化日志
        const logger = getLogger();
        logger.info('Python 引擎初始化完成');
        logger.info('Pandas + Openpyxl 已加载');
        logger.success('系统就绪，可以开始使用');

        // 默认显示第一个模块
        switchModule(NAV_ITEMS[0].id);

    } catch (err) {
        loadingStatus.textContent = `初始化失败: ${err.message}`;
        loadingBar.style.background = 'var(--color-error)';
        console.error('Init failed:', err);
    }
}

// ============ 构建侧边栏导航 ============
function buildNav() {
    const nav = document.getElementById('sidebar-nav');
    nav.innerHTML = NAV_ITEMS.map(item => `
    <div class="nav-item" data-module="${item.id}" id="nav-${item.id}">
      <span class="nav-icon">${item.icon}</span>
      <span>${item.label}</span>
    </div>
  `).join('');

    nav.addEventListener('click', (e) => {
        const navItem = e.target.closest('.nav-item');
        if (navItem) {
            switchModule(navItem.dataset.module);
        }
    });
}

// ============ 切换模块 ============
function switchModule(moduleId) {
    if (currentModule === moduleId) return;
    currentModule = moduleId;

    const item = NAV_ITEMS.find(n => n.id === moduleId);
    if (!item) return;

    // 更新导航高亮
    document.querySelectorAll('.nav-item').forEach(el => {
        el.classList.toggle('active', el.dataset.module === moduleId);
    });

    // 更新标题
    document.getElementById('page-title').textContent = item.label;

    // 渲染模块内容
    const container = document.getElementById('content-area');
    const logger = getLogger();
    item.render(container, logger);
}

// ============ 注册 Python 引擎适配器 ============
async function registerPythonEngines() {
    // 资金对账引擎适配器
    await runPython(`
import pandas as pd
import os
import glob
from datetime import datetime

def _engine_audit_run(accrual_dir, actual_dir, output_dir, log_fn):
    """资金对账 Web 适配器"""
    try:
        log_fn("开始资金对账处理...")

        # 查找计提文件
        accrual_files = glob.glob(os.path.join(accrual_dir, '*.xlsx')) + glob.glob(os.path.join(accrual_dir, '*.xls'))
        actual_files = glob.glob(os.path.join(actual_dir, '*.xlsx')) + glob.glob(os.path.join(actual_dir, '*.xls'))

        if not accrual_files:
            return {"success": False, "message": "未找到计提台账文件"}
        if not actual_files:
            return {"success": False, "message": "未找到实际账单文件"}

        log_fn(f"找到 {len(accrual_files)} 个计提文件, {len(actual_files)} 个实际文件")

        # 读取所有计提文件
        dfs_accrual = []
        for f in accrual_files:
            log_fn(f"读取计提: {os.path.basename(f)}")
            df = pd.read_excel(f)
            dfs_accrual.append(df)
        df_accrual = pd.concat(dfs_accrual, ignore_index=True) if dfs_accrual else pd.DataFrame()

        # 读取所有实际文件
        dfs_actual = []
        for f in actual_files:
            log_fn(f"读取实际: {os.path.basename(f)}")
            df = pd.read_excel(f)
            dfs_actual.append(df)
        df_actual = pd.concat(dfs_actual, ignore_index=True) if dfs_actual else pd.DataFrame()

        log_fn(f"计提数据: {len(df_accrual)} 行, 实际数据: {len(df_actual)} 行")

        # 简单对比：按列合并两个DataFrame的汇总信息
        summary_accrual = {}
        summary_actual = {}

        # 尝试找到金额相关列
        amount_cols_accrual = [c for c in df_accrual.columns if any(k in str(c) for k in ['金额', '费用', '合计', '总计', 'amount', 'total'])]
        amount_cols_actual = [c for c in df_actual.columns if any(k in str(c) for k in ['金额', '费用', '合计', '总计', 'amount', 'total'])]

        log_fn(f"计提金额列: {amount_cols_accrual}")
        log_fn(f"实际金额列: {amount_cols_actual}")

        # 生成对比报告
        report_rows = []
        report_rows.append({
            "项目": "文件数",
            "计提": len(accrual_files),
            "实际": len(actual_files),
            "差异": len(accrual_files) - len(actual_files)
        })
        report_rows.append({
            "项目": "数据行数",
            "计提": len(df_accrual),
            "实际": len(df_actual),
            "差异": len(df_accrual) - len(df_actual)
        })

        for col in amount_cols_accrual:
            val = pd.to_numeric(df_accrual[col], errors='coerce').sum()
            summary_accrual[col] = val

        for col in amount_cols_actual:
            val = pd.to_numeric(df_actual[col], errors='coerce').sum()
            summary_actual[col] = val

        # 合并金额对比
        all_cols = set(list(summary_accrual.keys()) + list(summary_actual.keys()))
        for col in all_cols:
            a_val = summary_accrual.get(col, 0)
            b_val = summary_actual.get(col, 0)
            report_rows.append({
                "项目": f"[金额] {col}",
                "计提": round(a_val, 2),
                "实际": round(b_val, 2),
                "差异": round(a_val - b_val, 2)
            })

        df_report = pd.DataFrame(report_rows)

        # 写入输出
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(output_dir, f"对账报告_{timestamp}.xlsx")

        with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
            df_report.to_excel(writer, sheet_name='对账汇总', index=False)
            df_accrual.to_excel(writer, sheet_name='计提明细', index=False)
            df_actual.to_excel(writer, sheet_name='实际明细', index=False)

        log_fn(f"对账报告已生成: {os.path.basename(out_path)}")
        diff_count = sum(1 for r in report_rows if r.get('差异', 0) != 0)

        return {
            "success": True,
            "message": f"对账完成，{diff_count} 项存在差异。报告已生成。",
            "output_file": out_path
        }
    except Exception as e:
        return {"success": False, "message": f"对账失败: {str(e)}"}

# 注册为全局模块
import types
engine_audit = types.ModuleType('engine_audit')
engine_audit.run_audit_web = _engine_audit_run
import sys
sys.modules['engine_audit'] = engine_audit
  `);

    // 账单清洗引擎适配器
    await runPython(`
import pandas as pd
import os
import glob
from datetime import datetime
import types, sys

def _engine_parser_run(source_dir, output_dir, log_fn):
    """账单清洗 Web 适配器"""
    try:
        log_fn("开始账单清洗处理...")

        source_files = glob.glob(os.path.join(source_dir, '*.xlsx')) + glob.glob(os.path.join(source_dir, '*.xls')) + glob.glob(os.path.join(source_dir, '*.csv'))

        if not source_files:
            return {"success": False, "message": "未找到源文件"}

        log_fn(f"找到 {len(source_files)} 个源文件")

        output_files = []
        for src in source_files:
            fname = os.path.basename(src)
            log_fn(f"处理: {fname}")

            try:
                if fname.endswith('.csv'):
                    df = pd.read_csv(src)
                else:
                    df = pd.read_excel(src)
            except Exception as e:
                log_fn(f"⚠️ 读取失败: {fname} - {e}")
                continue

            if df.empty:
                log_fn(f"⚠️ 空文件跳过: {fname}")
                continue

            log_fn(f"  数据行数: {len(df)}, 列数: {len(df.columns)}")

            # 数据清洗：去除空行、标准化列名
            df = df.dropna(how='all')
            df.columns = [str(c).strip() for c in df.columns]

            # 尝试识别并提取关键字段
            log_fn(f"  列名: {list(df.columns[:10])}{'...' if len(df.columns) > 10 else ''}")

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = os.path.splitext(fname)[0]
            out_path = os.path.join(output_dir, f"{base_name}_清洗_{timestamp}.xlsx")
            df.to_excel(out_path, index=False)
            output_files.append(out_path)
            log_fn(f"  ✅ 输出: {os.path.basename(out_path)}")

        if not output_files:
            return {"success": False, "message": "所有文件处理失败"}

        return {
            "success": True,
            "message": f"清洗完成，处理 {len(output_files)} 个文件。",
            "output_files": output_files
        }
    except Exception as e:
        return {"success": False, "message": f"清洗失败: {str(e)}"}

engine_parser = types.ModuleType('engine_parser')
engine_parser.run_parser_web = _engine_parser_run
sys.modules['engine_parser'] = engine_parser
  `);

    // 上传模板生成引擎适配器
    await runPython(`
import pandas as pd
import os
import glob
from datetime import datetime
import types, sys

def _engine_template_run(source_dir, output_dir, log_fn):
    """上传模板生成 Web 适配器"""
    try:
        log_fn("开始生成上传模板...")

        source_files = glob.glob(os.path.join(source_dir, '*.xlsx')) + glob.glob(os.path.join(source_dir, '*.xls'))

        if not source_files:
            return {"success": False, "message": "未找到源文件"}

        output_files = []
        for src in source_files:
            fname = os.path.basename(src)
            log_fn(f"处理: {fname}")

            try:
                df = pd.read_excel(src)
            except Exception as e:
                log_fn(f"⚠️ 读取失败: {fname} - {e}")
                continue

            if df.empty:
                log_fn(f"⚠️ 空文件跳过: {fname}")
                continue

            # 生成标准模板格式
            df.columns = [str(c).strip() for c in df.columns]

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = os.path.splitext(fname)[0]
            out_path = os.path.join(output_dir, f"上传模板_{base_name}_{timestamp}.xlsx")
            df.to_excel(out_path, index=False)
            output_files.append(out_path)
            log_fn(f"  ✅ 输出: {os.path.basename(out_path)}")

        if not output_files:
            return {"success": False, "message": "所有文件处理失败"}

        return {
            "success": True,
            "message": f"模板生成完成，处理 {len(output_files)} 个文件。",
            "output_files": output_files
        }
    except Exception as e:
        return {"success": False, "message": f"模板生成失败: {str(e)}"}

engine_template = types.ModuleType('engine_template')
engine_template.run_template_web = _engine_template_run
sys.modules['engine_template'] = engine_template
  `);

    // 报价稽核引擎适配器
    await runPython(`
import pandas as pd
import os
import glob
from datetime import datetime
import types, sys

def _engine_quote_audit_run(source_dir, output_dir, log_fn):
    """报价稽核 Web 适配器"""
    try:
        log_fn("开始报价稽核...")

        source_files = glob.glob(os.path.join(source_dir, '*.xlsx')) + glob.glob(os.path.join(source_dir, '*.xls'))

        if not source_files:
            return {"success": False, "message": "未找到账单文件"}

        output_files = []
        for src in source_files:
            fname = os.path.basename(src)
            log_fn(f"读取账单: {fname}")

            try:
                df = pd.read_excel(src)
            except Exception as e:
                log_fn(f"⚠️ 读取失败: {fname} - {e}")
                continue

            if df.empty:
                log_fn(f"⚠️ 空文件跳过: {fname}")
                continue

            # 识别列
            col_route = next((c for c in df.columns if '线路' in str(c)), None)
            col_item = next((c for c in df.columns if '费用' in str(c)), None)
            col_price = next((c for c in df.columns if '单价' in str(c)), None)
            col_qty = next((c for c in df.columns if '数量' in str(c)), None)

            log_fn(f"  识别列: 线路={col_route}, 费用={col_item}, 单价={col_price}, 数量={col_qty}")

            if not (col_route and col_item and col_price):
                log_fn(f"  ⚠️ 缺少必要列(线路/费用/单价)")
                continue

            # 结果标记（无报价库时仅输出数据摘要）
            results = []
            for idx, row in df.iterrows():
                results.append({
                    "行号": idx + 1,
                    "线路": str(row[col_route]).strip() if col_route else "",
                    "费用项": str(row[col_item]).strip() if col_item else "",
                    "单价": pd.to_numeric(row[col_price], errors='coerce') if col_price else 0,
                    "数量": pd.to_numeric(row[col_qty], errors='coerce') if col_qty else 0,
                    "备注": "待对比报价库"
                })

            if results:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                base_name = os.path.splitext(fname)[0]
                out_path = os.path.join(output_dir, f"稽核报告_{base_name}_{timestamp}.xlsx")
                pd.DataFrame(results).to_excel(out_path, index=False)
                output_files.append(out_path)
                log_fn(f"  ✅ 生成稽核报告: {os.path.basename(out_path)}")

        if not output_files:
            return {"success": False, "message": "所有文件处理失败或无有效数据"}

        return {
            "success": True,
            "message": f"稽核完成，生成 {len(output_files)} 份报告。",
            "output_files": output_files
        }
    except Exception as e:
        return {"success": False, "message": f"稽核失败: {str(e)}"}

engine_quote_audit = types.ModuleType('engine_quote_audit')
engine_quote_audit.run_quote_audit_web = _engine_quote_audit_run
sys.modules['engine_quote_audit'] = engine_quote_audit
  `);

    // 报价计算引擎（占位）
    await runPython(`
import types, sys

def _engine_calculate_run(brand, dest, weight, pallets, log_fn):
    """报价计算 Web 适配器（占位）"""
    log_fn(f"计算参数: 品牌={brand}, 目的地={dest}, 重量={weight}kg, 托板={pallets}")
    return {
        "success": False,
        "message": "报价计算功能需要加载报价规则文件，请先在设置中配置规则数据源。"
    }

engine_calculator = types.ModuleType('engine_calculator')
engine_calculator.run_calculate_web = _engine_calculate_run
sys.modules['engine_calculator'] = engine_calculator
  `);
}

// ============ 启动 ============
document.addEventListener('DOMContentLoaded', init);
