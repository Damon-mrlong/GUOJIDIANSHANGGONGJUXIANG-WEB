import pandas as pd
import os
import csv
from datetime import datetime

class UploadTemplateEngine:
    def __init__(self, workspace_root=None):
        self.workspace_root = workspace_root
        if workspace_root:
            self.config_dir = os.path.join(workspace_root, "规则文件")
            self.output_dir = os.path.join(workspace_root, "输出汇总文件夹")
            self.mapping_file = os.path.join(self.config_dir, "upload_mapping.csv")

    def set_workspace(self, root):
        self.workspace_root = root
        self.config_dir = os.path.join(root, "规则文件")
        self.output_dir = os.path.join(root, "输出汇总文件夹")
        self.mapping_file = os.path.join(self.config_dir, "upload_mapping.csv")

    def ensure_default_mapping(self):
        if not os.path.exists(self.mapping_file):
            try:
                # 创建默认映射文件模版
                header = ["Target_Column", "Source_Column", "Fixed_Value", "Description"]
                data = [
                    ["OMS单号", "计提账单编号", "", "对应系统中的唯一ID"],
                    ["费用类型", "", "物流杂费", "固定值"],
                    ["金额", "实际账单金额", "", ""],
                    ["供应商", "供应商名称", "", ""],
                    ["备注", "原因反馈", "", ""]
                ]
                with open(self.mapping_file, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(header)
                    writer.writerows(data)
            except: pass

    def load_mapping(self):
        self.ensure_default_mapping()
        mapping = []
        if os.path.exists(self.mapping_file):
            try:
                df = pd.read_csv(self.mapping_file)
                for _, row in df.iterrows():
                    mapping.append({
                        "target": str(row.get("Target_Column", "")).strip(),
                        "source": str(row.get("Source_Column", "")).strip(),
                        "fixed": str(row.get("Fixed_Value", "")).strip()
                    })
            except Exception as e:
                print(f"读取映射失败: {e}")
        return mapping

    def generate_template(self, source_file, log_callback=None):
        if not self.workspace_root:
            return False, "工作区未设置"

        def log(msg):
            if log_callback: log_callback(msg)
            else: print(msg)

        if not os.path.exists(source_file):
             return False, "源文件不存在"

        log("正在加载映射规则...")
        mapping = self.load_mapping()
        if not mapping:
             return False, "映射规则为空或读取失败"

        log(f"读取源文件: {os.path.basename(source_file)}")
        try:
            # 尝试读取第一张表
            df_source = pd.read_excel(source_file)
        except Exception as e:
            return False, f"读取 Excel 失败: {e}"

        if df_source.empty:
            return False, "源文件为空"

        log("正在转换数据...")
        # 准备目标 DataFrame
        target_data = {}
        
        # 将 source 列名转为字符串以防万一
        source_cols = [str(c).strip() for c in df_source.columns]
        df_source.columns = source_cols

        for rule in mapping:
            target_col = rule["target"]
            source_col = rule["source"]
            fixed_val = rule["fixed"]
            
            if not target_col or target_col == 'nan':
                continue

            if fixed_val and fixed_val != 'nan':
                # 使用固定值
                target_data[target_col] = [fixed_val] * len(df_source)
            elif source_col and source_col != 'nan':
                # 从源读取
                if source_col in df_source.columns:
                    target_data[target_col] = df_source[source_col]
                else:
                    log(f"⚠️ 警告: 源文件中找不到列 '{source_col}'，目标列 '{target_col}' 将为空。")
                    target_data[target_col] = [""] * len(df_source)
            else:
                 # 既无源也无固定值
                 target_data[target_col] = [""] * len(df_source)

        df_target = pd.DataFrame(target_data)
        
        # 输出
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
            
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_file = os.path.join(self.output_dir, f"OMS上传模板_{timestamp}.xlsx")
        
        df_target.to_excel(out_file, index=False)
        return True, f"生成成功！文件: {os.path.basename(out_file)}"
