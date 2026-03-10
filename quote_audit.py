import pandas as pd
import os
from datetime import datetime

class QuoteAuditEngine:
    def __init__(self, workspace_root=None):
        self.workspace_root = workspace_root
        if workspace_root:
            self.config_dir = os.path.join(workspace_root, "规则文件")
            self.output_dir = os.path.join(workspace_root, "输出汇总文件夹")
            self.db_file = os.path.join(self.config_dir, "price_database.xlsx")

    def set_workspace(self, root):
        self.workspace_root = root
        self.config_dir = os.path.join(root, "规则文件")
        self.output_dir = os.path.join(root, "输出汇总文件夹")
        self.db_file = os.path.join(self.config_dir, "price_database.xlsx")

    def ensure_default_database(self):
        if not os.path.exists(self.db_file):
            try:
                # 创建默认报价库模版
                data = {
                    "线路": ["示例线路A", "示例线路B"],
                    "费用项": ["运费", "操作费"],
                    "标准单价": [10.5, 2.0],
                    "生效日期": ["2023-01-01", "2023-01-01"]
                }
                pd.DataFrame(data).to_excel(self.db_file, index=False)
            except: pass

    def run_audit(self, bill_path, log_callback=None):
        if not self.workspace_root:
            return False, "工作区未设置"

        def log(msg):
            if log_callback: log_callback(msg)
            else: print(msg)

        self.ensure_default_database()
        
        if not os.path.exists(self.db_file):
             return False, "找不到报价库文件: 00_Config/price_database.xlsx"
        if not os.path.exists(bill_path):
             return False, "找不到账单文件"

        log(f"加载报价库: {os.path.basename(self.db_file)}")
        try:
            df_db = pd.read_excel(self.db_file)
            # 建立查找字典: (线路, 费用项) -> 单价
            price_map = {}
            for _, row in df_db.iterrows():
                key = (str(row.get('线路', '')).strip(), str(row.get('费用项', '')).strip())
                val = pd.to_numeric(row.get('标准单价', 0), errors='coerce')
                price_map[key] = val
            log(f"已加载 {len(price_map)} 条单价记录")
        except Exception as e:
            return False, f"读取报价库失败: {e}"

        log(f"读取账单文件: {os.path.basename(bill_path)}")
        try:
            df_bill = pd.read_excel(bill_path)
        except Exception as e:
            return False, f"读取账单失败: {e}"

        if df_bill.empty:
            return False, "账单文件为空"

        log("正在稽核单价...")
        audit_results = []
        
        # 尝试识别关键列
        # 假设列名包含：线路, 费用项, 单价, 数量
        col_route = next((c for c in df_bill.columns if "线路" in str(c)), None)
        col_item = next((c for c in df_bill.columns if "费用" in str(c)), None)
        col_price = next((c for c in df_bill.columns if "单价" in str(c)), None)
        col_qty = next((c for c in df_bill.columns if "数量" in str(c)), None)
        
        if not (col_route and col_item and col_price):
             return False, f"账单缺少必要列(线路, 费用*, 单价*). 找到: {[col_route, col_item, col_price]}"

        for idx, row in df_bill.iterrows():
            r_val = str(row[col_route]).strip()
            i_val = str(row[col_item]).strip()
            p_val = pd.to_numeric(row[col_price], errors='coerce') or 0
            q_val = 0
            if col_qty:
                q_val = pd.to_numeric(row[col_qty], errors='coerce') or 0

            std_price = price_map.get((r_val, i_val))
            
            is_violation = "否"
            diff_amt = 0
            std_price_display = std_price
            
            if std_price is not None:
                if abs(p_val - std_price) > 0.001:
                    is_violation = "是"
                    diff_amt = (p_val - std_price) * q_val
            else:
                std_price_display = "未找到报价"
                # 是否视为违规取决于业务逻辑，暂定为是
                is_violation = "未知"

            if is_violation == "是" or is_violation == "未知":
                # 只记录有问题的，或者全部记录？需求暗示"稽核报告"，通常包含差异
                audit_results.append({
                    "行号": idx + 1,
                    "线路": r_val,
                    "费用项": i_val,
                    "账单单价": p_val,
                    "标准单价": std_price_display,
                    "数量": q_val,
                    "差异金额": diff_amt,
                    "是否违规": is_violation
                })

        if not audit_results:
            return True, "稽核通过，未发现价格差异。"

        # 输出
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
            
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_file = os.path.join(self.output_dir, f"报价稽核报告_{timestamp}.xlsx")
        
        pd.DataFrame(audit_results).to_excel(out_file, index=False)
        return True, f"发现 {len(audit_results)} 条异常/差异。报告: {os.path.basename(out_file)}"
