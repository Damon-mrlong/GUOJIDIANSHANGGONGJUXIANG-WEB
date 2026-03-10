import pandas as pd
import os
import string
import csv
import glob
from datetime import datetime

class FinanceAuditEngine:
    def __init__(self, workspace_root=None):
        self.workspace_root = workspace_root
        # 默认路径配置 (An.md standard)
        self.config_dir = None
        self.input_acc_dir = None
        self.input_act_dir = None
        self.output_dir = None
        self.rules_file = None
        
        if workspace_root:
            self.set_workspace(workspace_root)

    def set_workspace(self, root):
        self.workspace_root = root
        self.config_dir = os.path.join(root, "规则文件")
        self.input_acc_dir = os.path.join(root, "计提实际账单核对", "计提台账文件夹")
        self.input_act_dir = os.path.join(root, "计提实际账单核对", "实际台账文件夹")
        self.output_dir = os.path.join(root, "输出汇总文件夹")
        self.rules_file = os.path.join(self.config_dir, "audit_rules.csv")
    
    def set_custom_dirs(self, accrual_dir, actual_dir):
        """手动设置计提和实际账单目录"""
        self.input_acc_dir = accrual_dir
        self.input_act_dir = actual_dir

    def col2num(self, col_str):
        num = 0
        for c in col_str:
            if c in string.ascii_letters:
                num = num * 26 + (ord(c.upper()) - ord('A')) + 1
        return num - 1

    def ensure_default_rules(self):
        """如果规则文件不存在，创建默认规则"""
        if not self.rules_file: return
        if os.path.exists(self.rules_file): return

        # 默认规则
        default_data = [
            ["Type", "Role_Index", "Acc_Vendor", "Acc_Amt", "Act_Vendor", "Act_Amt", "Acc_Id", "Acc_Person", "Act_Id", "Act_Self_Id"],
            ["快递物流", "1", "D", "T", "H", "W", "A", "I", "E", "B"],
            ["区间调拨", "1", "Y", "AJ|AK", "AN", "AO|AP", "B", "J", "E", "B"],
            ["区间调拨", "2", "Z", "AL", "AC", "AQ", "B", "J", "E", "B"],
            ["区间调拨", "3", "AA", "AP", "AD", "AT", "B", "J", "E", "B"],
            ["区间调拨", "4", "R", "AN|AO", "P", "AR|AS", "B", "J", "E", "B"],
            ["区间调拨", "5", "E", "AQ|AR", "L", "AU|AV", "B", "J", "E", "B"],
            ["一线入境", "1", "D", "BK", "H", "BQ", "A", "L", "E", "B"],
            ["一线入境", "2", "AL", "BG", "AS", "BI", "A", "L", "E", "B"],
            ["一线入境", "3", "AM", "BH", "AQ", "BJ", "A", "L", "E", "B"],
            ["一线入境", "4", "AN", "BI", "AW", "BK", "A", "L", "E", "B"],
            ["一线入境", "5", "AL", "AU", "AS", "AU", "A", "L", "E", "B"],
            ["一线入境", "6", "AN", "AW", "AW", "AY", "A", "L", "E", "B"]
        ]
        
        try:
            if not os.path.exists(self.config_dir):
                os.makedirs(self.config_dir)
            with open(self.rules_file, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerows(default_data)
        except Exception as e:
            print(f"创建默认规则失败: {e}")

    def load_rules(self):
        self.ensure_default_rules()
        rules = {}
        if not self.rules_file or not os.path.exists(self.rules_file):
            return rules

        try:
            df = pd.read_csv(self.rules_file)
            for _, row in df.iterrows():
                t = row['Type']
                if t not in rules:
                    rules[t] = {
                        "desc": t,
                        "acc_id": row['Acc_Id'],
                        "acc_person": row['Acc_Person'],
                        "act_id": row['Act_Id'],
                        "act_self_id": row['Act_Self_Id'],
                        "mapping": []
                    }
                
                acc_amts = str(row['Acc_Amt']).split('|')
                act_amts = str(row['Act_Amt']).split('|')
                rules[t]['mapping'].append([
                    row['Acc_Vendor'], acc_amts, row['Act_Vendor'], act_amts
                ])
        except Exception as e:
            print(f"读取规则失败: {e}")
            
        return rules
    
    def load_warehouse_map(self, custom_map_path=None):
        """读取仓库列表.csv或用户指定的映射文件"""
        wh_path = None
        
        if custom_map_path and os.path.exists(custom_map_path):
             wh_path = custom_map_path
        elif self.workspace_root:
             p = os.path.join(self.config_dir, "仓库列表.csv")
             if os.path.exists(p): wh_path = p
        
        if not wh_path:
            return {}

        try:
            try:
                df = pd.read_csv(wh_path, encoding='gbk')
            except:
                df = pd.read_csv(wh_path, encoding='utf-8')
            
            if df.shape[1] < 24: 
                return {}
                
            wh_map = dict(zip(df.iloc[:, 1].astype(str).str.strip(), df.iloc[:, 23].astype(str).str.strip()))
            clean_map = {k: v for k, v in wh_map.items() if k != 'nan' and v != 'nan'}
            return clean_map
        except Exception:
            return {}

    def auto_detect_files(self):
        """探测是否有文件，返回数量信息"""
        found_info = {
            "accrual_count": 0,
            "actual_count": 0,
            "map": None
        }
        
        if not self.workspace_root:
            return found_info

        if os.path.exists(self.input_acc_dir):
            files = glob.glob(os.path.join(self.input_acc_dir, "*.xlsx")) + glob.glob(os.path.join(self.input_acc_dir, "*.xls"))
            files = [f for f in files if not os.path.basename(f).startswith("~$")]
            found_info["accrual_count"] = len(files)

        if os.path.exists(self.input_act_dir):
            files = glob.glob(os.path.join(self.input_act_dir, "*.xlsx")) + glob.glob(os.path.join(self.input_act_dir, "*.xls"))
            files = [f for f in files if not os.path.basename(f).startswith("~$")]
            found_info["actual_count"] = len(files)
        
        map_path = os.path.join(self.config_dir, "仓库列表.csv")
        if os.path.exists(map_path):
            found_info["map"] = map_path
            
        return found_info



    def extract_data(self, df, id_col, person_col, mapping_list, is_actual=False, self_id_col=None):
        extracted = []
        id_idx = self.col2num(id_col)
        
        person_idx = -1
        if not is_actual and person_col and str(person_col) != 'nan':
            person_idx = self.col2num(person_col)
            
        self_id_idx = -1
        if is_actual and self_id_col and str(self_id_col) != 'nan':
            self_id_idx = self.col2num(self_id_col)

        for idx, row in df.iterrows():
            try:
                if id_idx >= len(row): continue
                raw_id = row.iloc[id_idx]
                if pd.isna(raw_id): continue
                row_id = str(raw_id).strip()
                if not row_id: continue

                if "单号" in row_id or "编号" in row_id or "计提" in row_id:
                    continue

                person_name = ""
                if person_idx >= 0 and person_idx < len(row):
                    try:
                        p_val = str(row.iloc[person_idx]).strip()
                        person_name = p_val if p_val != 'nan' else ""
                    except: pass
                    
                act_bill_no = ""
                if self_id_idx >= 0 and self_id_idx < len(row):
                    try:
                        b_val = str(row.iloc[self_id_idx]).strip()
                        act_bill_no = b_val if b_val != 'nan' else ""
                    except: pass

                for role_i, config in enumerate(mapping_list):
                    if is_actual:
                        vendor_col = config[2]
                        amt_cols = config[3]
                    else:
                        vendor_col = config[0]
                        amt_cols = config[1]
                    
                    v_idx = self.col2num(vendor_col)
                    vendor_name = ""
                    if v_idx < len(row):
                        try:
                            v_val = str(row.iloc[v_idx]).strip()
                            vendor_name = v_val if v_val != 'nan' else ""
                        except: pass
                    
                    if ("供应商" in vendor_name or "名称" in vendor_name or "仓库" in vendor_name) and len(vendor_name) < 10:
                         pass 

                    total_amt = 0.0
                    for a_col in amt_cols:
                        a_idx = self.col2num(a_col)
                        if a_idx < len(row):
                            try:
                                val = pd.to_numeric(row.iloc[a_idx], errors='coerce')
                                if pd.notna(val): total_amt += val
                            except: pass
                    
                    if total_amt != 0 or vendor_name:
                        item = {
                            'Global_ID': row_id,
                            'Person': person_name,
                            'Vendor': vendor_name,
                            'Amount': total_amt,
                            'Role_Index': role_i + 1
                        }
                        if is_actual:
                            item['Act_Bill_No'] = act_bill_no
                        extracted.append(item)
                        
            except Exception: continue
        
        return pd.DataFrame(extracted)

    def load_files_by_keyword(self, directory, keyword, log_callback=None):
        """
        在指定目录下查找包含关键字的文件并加载
        """
        if not os.path.exists(directory):
            return pd.DataFrame(), []
            
        files = glob.glob(os.path.join(directory, "*.xlsx")) + glob.glob(os.path.join(directory, "*.xls"))
        # 过滤: 排除临时文件，且文件名包含关键字
        target_files = []
        for f in files:
            fname = os.path.basename(f)
            if fname.startswith("~$"): continue
            if keyword in fname:
                target_files.append(f)
        
        if not target_files:
            return pd.DataFrame(), []
            
        dfs = []
        loaded_names = []
        for f in target_files:
            try:
                df = pd.read_excel(f, header=None)
                if not df.empty:
                    dfs.append(df)
                    loaded_names.append(os.path.basename(f))
            except Exception as e:
                if log_callback: 
                    log_callback(f"⚠️ 无法读取 {os.path.basename(f)}: {e}")
        
        if not dfs:
            return pd.DataFrame(), []
            
        combined = pd.concat(dfs, ignore_index=True)
        return combined, loaded_names

    def run_audit(self, map_path=None, log_callback=None):
        """
        批量对账入口
        不需传入具体文件路径，根据规则中的Type自动匹配 Inputs_Accrual 和 Inputs_Actual 下的文件
        """
        if not self.workspace_root:
            return False, "工作区未设置"
        
        def log(msg):
            if log_callback: log_callback(msg)
            else: print(msg)

        log("正在加载审计规则...")
        rules = self.load_rules()
        if not rules:
            return False, "未找到有效的审计规则"

        log(f"加载了 {len(rules)} 组规则。")
        
        wh_map = self.load_warehouse_map(map_path)
        if wh_map:
            log(f"已加载仓库映射表: {len(wh_map)} 条记录")
        else:
            log("未加载仓库映射表，将直接使用源名称。")
        
        all_results = []
        
        # 遍历每一种规则类型，分别去加载对应的文件
        for key, cfg in rules.items():
            log(f"────────────────────────────────────────")
            log(f"正在分析任务: [{key}]")
            
            # 1. 加载计提台账 (文件名包含 key)
            df_acc_raw, acc_files = self.load_files_by_keyword(self.input_acc_dir, key, log_callback)
            if df_acc_raw.empty:
                log(f"  ⚠️ 跳过: 在计提文件夹未找到包含 '{key}' 的文件")
                continue
            log(f"  已加载计提文件: {', '.join(acc_files)}")

            # 2. 加载实际台账 (文件名包含 key)
            df_act_raw, act_files = self.load_files_by_keyword(self.input_act_dir, key, log_callback)
            if df_act_raw.empty:
                log(f"  ⚠️ 跳过: 在实际文件夹未找到包含 '{key}' 的文件")
                continue
            log(f"  已加载实际文件: {', '.join(act_files)}")

            # 3. 开始执行匹配逻辑
            try:
                df_acc = self.extract_data(df_acc_raw, cfg['acc_id'], cfg['acc_person'], cfg['mapping'], False)
                df_act = self.extract_data(df_act_raw, cfg['act_id'], None, cfg['mapping'], True, cfg.get('act_self_id'))

                if df_acc.empty: 
                    log("  ❌ 提取数据失败: 计提数据为空")
                    continue

                acc_grouped = df_acc.groupby(['Global_ID', 'Role_Index']).agg({
                    'Amount': 'sum', 'Vendor': 'first', 'Person': 'first'
                }).reset_index().rename(columns={'Amount': 'acc_amt', 'Vendor': 'vendor', 'Person': 'person'})

                act_grouped = pd.DataFrame()
                if not df_act.empty:
                    def join_ids(x):
                        valid_ids = sorted(set([str(i) for i in x if str(i) and str(i) != 'nan']))
                        return ','.join(valid_ids)
                    act_grouped = df_act.groupby(['Global_ID', 'Role_Index']).agg({
                        'Amount': 'sum', 'Act_Bill_No': join_ids 
                    }).reset_index().rename(columns={'Amount': 'act_amt', 'Act_Bill_No': 'act_bill_nos'})
                
                if act_grouped.empty:
                    merged = acc_grouped
                    merged['act_amt'] = 0
                    merged['act_bill_nos'] = ""
                else:
                    merged = pd.merge(acc_grouped, act_grouped, on=['Global_ID', 'Role_Index'], how='outer')

                merged['acc_amt'] = merged['acc_amt'].fillna(0)
                merged['act_amt'] = merged['act_amt'].fillna(0)
                merged['diff'] = merged['acc_amt'] - merged['act_amt']
                merged['vendor'] = merged['vendor'].fillna("未知")
                merged['act_bill_nos'] = merged['act_bill_nos'].fillna("")
                merged['Type'] = key
                
                all_results.append(merged)
                log(f"  ✅ [{key}] 匹配完成: {len(merged)} 条记录")

            except Exception as e:
                log(f"  ❌ [{key}] 处理出错: {e}")

        if not all_results:
            return False, "未生成任何对账结果。请检查文件名是否包含规则类型名称(如'快递物流')。"

        log("正在生成最终报表...")
        final = pd.concat(all_results, ignore_index=True)
        
        # 应用映射
        if wh_map:
            mapped_vendor = final['vendor'].map(wh_map)
            final['供应商名称'] = mapped_vendor.fillna(final['vendor'])
        else:
            final['供应商名称'] = final['vendor']

        final['计提账单编号'] = final['Global_ID']
        final['实际账单编号'] = final['act_bill_nos']
        final['物流对接人'] = final['person']
        final['计提账单金额'] = final['acc_amt']
        final['实际账单金额'] = final['act_amt']
        
        def get_reason(row):
            vendor_str = str(row['供应商名称'])
            if "菜鸟" in vendor_str: return "正常"
            if abs(row['diff']) <= 0.05: return "正常"
            if row['acc_amt'] > 0 and row['act_amt'] == 0: return "实际漏录"
            if row['acc_amt'] == 0 and row['act_amt'] > 0: return "无头账(未计提)"
            return "金额差异"

        final['原因反馈'] = final.apply(get_reason, axis=1)

        cols = ['Type', '供应商名称', '计提账单编号', '实际账单编号', '物流对接人', 
                '计提账单金额', '实际账单金额', 'diff', '原因反馈']
        
        out_cols = [c for c in cols if c in final.columns]
        output_df = final[out_cols].sort_values(by=['原因反馈', 'Type', 'diff'], ascending=[False, True, False])
        output_df.rename(columns={'diff': '差异'}, inplace=True)

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
            
        out_file = os.path.join(self.output_dir, f"计提实际差异数据_{timestamp}.xlsx")
        output_df.to_excel(out_file, index=False)
        
        return True, f"对账成功！已生成报告: {os.path.basename(out_file)}"