import pandas as pd
import os
import re
from datetime import datetime
from pathlib import Path

# ================= 工具函数 =================

def normalize_text(text):
    """
    标准化文本：全角转半角、去除空格
    用于确保物流商品和费用项的准确匹配
    """
    if pd.isna(text):
        return ""
    text = str(text).strip()
    # 全角破折号、减号转半角
    text = text.replace('－', '-').replace('—', '-').replace('─', '-')
    return text

class BillParserEngine:
    def __init__(self, workspace_root=None):
        self.workspace_root = workspace_root
        if workspace_root:
            self.config_dir = os.path.join(workspace_root, "规则文件")
            self.output_dir = os.path.join(workspace_root, "输出汇总文件夹")
            self.source_dir = os.path.join(workspace_root, "快递计提台账处理")
            self.rule_file = os.path.join(self.config_dir, "菜鸟费用计提规则V2.xlsx")
            self.match_file = os.path.join(self.config_dir, "品牌仓库对接人匹配关系.xlsx")
            self.template_file = os.path.join(self.config_dir, "快递发货物流台账批量导入模板.xlsx")

    def set_workspace(self, root):
        self.workspace_root = root
        self.config_dir = os.path.join(root, "规则文件")
        self.output_dir = os.path.join(root, "输出汇总文件夹")
        self.source_dir = os.path.join(root, "快递计提台账处理")
        self.rule_file = os.path.join(self.config_dir, "菜鸟费用计提规则V2.xlsx")
        self.match_file = os.path.join(self.config_dir, "品牌仓库对接人匹配关系.xlsx")
        self.template_file = os.path.join(self.config_dir, "快递发货物流台账批量导入模板.xlsx")

    def load_rule_data(self):
        if not os.path.exists(self.rule_file):
            return set(), set()
            
        try:
            df_rule = pd.read_excel(self.rule_file)
            file_kuai_di_set = set()
            file_cang_chu_set = set()

            for _, row in df_rule.iterrows():
                fee_item = row['费用项']
                category = row['管理报表归属费用科目']

                if pd.notna(category):
                    if category == '快递费':
                        file_kuai_di_set.add(fee_item)
                    elif category == '仓内增值费':
                        file_cang_chu_set.add(fee_item)
            return file_kuai_di_set, file_cang_chu_set
        except Exception:
            return set(), set()

    def load_brand_match_data(self):
        if not os.path.exists(self.match_file):
            return {}, {}, {}
            
        try:
            df_match = pd.read_excel(self.match_file)
            cols = df_match.columns.tolist()
            # 假设第4列(索引3)是品牌，最后一列是是否中心仓
            # 如果列不够，可能会报错，需注意
            if len(cols) < 4: return {}, {}, {}
            
            brand_col = cols[3]
            is_center_col = cols[-1]

            bonded_match_dict = {}
            center_match_dict = {}
            overseas_match_dict = {}  # 新增海外仓字典

            for _, row in df_match.iterrows():
                brand = row[brand_col]
                warehouse_type = str(row[is_center_col]).strip() if pd.notna(row[is_center_col]) else ""

                match_record = {
                    '货主编码': row[cols[0]],
                    '发货仓库名称': row[cols[1]],
                    '业务月份': row[cols[2]],
                    '品牌名称': row[brand_col],
                    '店铺名称': row[cols[4]],
                    '物流对接人': row[cols[5]],
                    '快递供应商物流对接人': row[cols[6]]
                }

                # 根据"是否中心仓"列的值分配到不同字典
                if warehouse_type == '是':
                    center_match_dict[brand] = match_record
                elif warehouse_type == '海外仓':
                    overseas_match_dict[brand] = match_record
                else:
                    bonded_match_dict[brand] = match_record

            return bonded_match_dict, center_match_dict, overseas_match_dict
        except Exception:
            return {}, {}, {}


    def build_classification_sets(self, file_kuai_di_set, file_cang_chu_set):
        INDEPENDENT_ITEMS = {'货值赔付', '服务赔付'}
        MANUAL_KUAI_DI_SET = {'指定效期残出库费', '经济上门', '优质上门'}
        MANUAL_CANG_CHU_SET = {'防尘袋安装费', '质检拒收费', '拆预包费', '隐形眼镜品类操作费'}

        FINAL_KUAI_DI_SET = file_kuai_di_set | MANUAL_KUAI_DI_SET
        FINAL_CANG_CHU_SET = file_cang_chu_set | MANUAL_CANG_CHU_SET

        return INDEPENDENT_ITEMS, FINAL_KUAI_DI_SET, FINAL_CANG_CHU_SET

    def calculate_metrics(self, df, INDEPENDENT_ITEMS, FINAL_KUAI_DI_SET, FINAL_CANG_CHU_SET):
        report_dict = {}
        
        # 快递发货单量 - 扩展支持3种基础服务费类型
        base_service_types = ['基础服务费', '基础服务费-非OTC', '基础服务费-OTC']
        report_dict['快递发货单量'] = df[df['费用项'].isin(base_service_types)]['主单行数量'].sum()
        
        # 赔付
        report_dict['货值赔付'] = df[df['费用项'] == '货值赔付']['支付金额'].sum()
        report_dict['服务赔付'] = df[df['费用项'] == '服务赔付']['支付金额'].sum()

        # 快递费
        kuai_di_filter = (df['费用项'].isin(FINAL_KUAI_DI_SET) & ~df['费用项'].isin(INDEPENDENT_ITEMS))
        report_dict['快递费'] = df[kuai_di_filter]['支付金额'].sum()

        # 仓内增值费
        cang_chu_filter = (df['费用项'].isin(FINAL_CANG_CHU_SET) & ~df['费用项'].isin(INDEPENDENT_ITEMS))
        report_dict['仓内增值费'] = df[cang_chu_filter]['支付金额'].sum()


        # 其他费用：排除所有已分类的费用项
        # 注意：基础服务费已在FINAL_KUAI_DI_SET中（归属于快递费），无需再次排除
        processed_items = INDEPENDENT_ITEMS | FINAL_KUAI_DI_SET | FINAL_CANG_CHU_SET
        df_other = df[~df['费用项'].isin(processed_items)]
        other_summary = df_other.groupby('费用项')['支付金额'].sum()
        report_dict['其他待纳入统计的费用'] = other_summary[other_summary != 0].to_dict()

        return report_dict

    def process_brand_data(self, df_full, brand_name, INDEPENDENT_ITEMS, FINAL_KUAI_DI_SET, FINAL_CANG_CHU_SET, bonded_match_dict, center_match_dict, overseas_match_dict):
        """
        处理品牌数据，支持三仓拆分（保税仓+中心仓+海外仓）和金额校验
        """
        # 获取匹配信息
        bonded_match_info = bonded_match_dict.get(brand_name)
        center_match_info = center_match_dict.get(brand_name)
        overseas_match_info = overseas_match_dict.get(brand_name)
        match_status = 'matched' if bonded_match_info else 'not_matched'

        if '物流商品' not in df_full.columns:
            raise ValueError(f"缺少'物流商品'列")

        # 数据清洗
        df_full = df_full.copy()
        df_full['费用项'] = df_full['费用项'].str.strip().fillna('')
        df_full['物流商品'] = df_full['物流商品'].apply(normalize_text)

        for col in ['支付金额', '主单行数量']:
            df_full[col] = df_full[col].astype(str)
            df_full[col] = df_full[col].str.replace(r'CNY|¥|\\$|,', '', regex=True).str.strip()
            df_full[col] = pd.to_numeric(df_full[col], errors='coerce').fillna(0)

        # 计算源数据总金额
        source_total = df_full['支付金额'].sum()

        # 判断中心仓拆分
        center_g_value = normalize_text("商家-保税中心仓")
        center_rows = df_full[df_full['物流商品'] == center_g_value]
        has_center = len(center_rows) > 0 and (center_rows['费用项'] == '基础服务费').any()

        # 判断海外仓拆分
        overseas_g_value = normalize_text("菜鸟海外仓配服务-商家")
        overseas_rows = df_full[df_full['物流商品'] == overseas_g_value]
        has_overseas = len(overseas_rows) > 0 and (
            overseas_rows['费用项'].isin(['基础服务费-非OTC', '基础服务费-OTC'])
        ).any()

        # 数据拆分
        warnings = []  # 警告信息列表
        
        if has_center or has_overseas:
            # 排除中心仓和海外仓，剩余为保税仓
            df_bonded = df_full[
                (df_full['物流商品'] != center_g_value) &
                (df_full['物流商品'] != overseas_g_value)
            ]
            df_center = center_rows if has_center else pd.DataFrame()
            df_overseas = overseas_rows if has_overseas else pd.DataFrame()
        else:
            df_bonded = df_full
            df_center = pd.DataFrame()
            df_overseas = pd.DataFrame()

        # 配置缺失检查
        if has_center and not center_match_info:
            warnings.append("缺少中心仓配置")
            # 合并到保税仓
            df_bonded = pd.concat([df_bonded, df_center], ignore_index=True)
            df_center = pd.DataFrame()
            has_center = False

        if has_overseas and not overseas_match_info:
            warnings.append("缺少海外仓配置")
            # 合并到保税仓
            df_bonded = pd.concat([df_bonded, df_overseas], ignore_index=True)
            df_overseas = pd.DataFrame()
            has_overseas = False

        # 计算各仓指标
        report_bonded = self.calculate_metrics(df_bonded, INDEPENDENT_ITEMS, FINAL_KUAI_DI_SET, FINAL_CANG_CHU_SET)
        report_center = self.calculate_metrics(df_center, INDEPENDENT_ITEMS, FINAL_KUAI_DI_SET, FINAL_CANG_CHU_SET) if not df_center.empty else None
        report_overseas = self.calculate_metrics(df_overseas, INDEPENDENT_ITEMS, FINAL_KUAI_DI_SET, FINAL_CANG_CHU_SET) if not df_overseas.empty else None

        # 提取其他费用
        bonded_other_fees = report_bonded.get('其他待纳入统计的费用', {})
        center_other_fees = report_center.get('其他待纳入统计的费用', {}) if report_center else {}
        overseas_other_fees = report_overseas.get('其他待纳入统计的费用', {}) if report_overseas else {}

        # 金额校验 - 计算输出金额总和（与row_total逻辑保持一致）
        def calc_output_amount(report, other_fees):
            if not report:
                return 0
            return (
                report.get('快递费', 0) +                    # 已包含基础服务费
                abs(report.get('仓内增值费', 0)) +          # 使用abs
                sum(other_fees.values()) -                   # 其他费用
                abs(report.get('服务赔付', 0)) -            # 减去abs
                abs(report.get('货值赔付', 0))              # 减去abs
            )

        output_total = (
            calc_output_amount(report_bonded, bonded_other_fees) +
            calc_output_amount(report_center, center_other_fees) +
            calc_output_amount(report_overseas, overseas_other_fees)
        )

        amount_diff = abs(source_total - output_total)
        amount_valid = amount_diff <= 0.01

        return (
            source_total, 
            report_bonded, report_center, report_overseas,
            bonded_match_info, center_match_info, overseas_match_info,
            has_center, has_overseas,
            output_total, amount_valid, amount_diff,
            bonded_other_fees, center_other_fees, overseas_other_fees,
            warnings,
            match_status
        )

    def run_parser(self, source_file, log_callback=None):
        if not self.workspace_root:
            return False, "工作区未设置"

        def log(msg):
            if log_callback: log_callback(msg)
            else: print(msg)

        if not os.path.exists(self.rule_file):
            return False, f"找不到规则文件: 规则文件/菜鸟费用计提规则V2.xlsx"
        if not os.path.exists(self.match_file):
            return False, f"找不到匹配文件: 规则文件/品牌仓库对接人匹配关系.xlsx"
        if not os.path.exists(self.template_file):
            return False, f"找不到模板文件: 规则文件/快递发货物流台账批量导入模板.xlsx"

        log("正在加载清洗规则...")
        file_kuai_di_set, file_cang_chu_set = self.load_rule_data()
        INDEPENDENT_ITEMS, FINAL_KUAI_DI_SET, FINAL_CANG_CHU_SET = self.build_classification_sets(file_kuai_di_set, file_cang_chu_set)
        
        log("正在加载品牌匹配关系...")
        bonded_match_dict, center_match_dict, overseas_match_dict = self.load_brand_match_data()

        log(f"正在读取源文件: {os.path.basename(source_file)}")
        try:
            xl_file = pd.ExcelFile(source_file)
        except Exception as e:
            return False, f"无法读取 Excel: {e}"

        excel_rows = []
        processed_count = 0
        error_log = []
        warning_list = []  # 收集配置缺失警告

        for sheet_name in xl_file.sheet_names:
            log(f"处理品牌: {sheet_name}")
            try:
                df = pd.read_excel(xl_file, sheet_name=sheet_name)
                if df.empty: 
                    log(f"  - 跳过空Sheet")
                    continue
                
                # 检查必需列
                if '物流商品' not in df.columns:
                    log(f"  - ⚠️ 缺少'物流商品'列，跳过")
                    continue

                # 调用process_brand_data（新增overseas_match_dict参数）
                (source_total, report_bonded, report_center, report_overseas,
                 bonded_match_info, center_match_info, overseas_match_info,
                 has_center, has_overseas,
                 output_total, amount_valid, amount_diff,
                 bonded_other_fees, center_other_fees, overseas_other_fees,
                 warnings, match_status) = self.process_brand_data(
                    df, sheet_name, INDEPENDENT_ITEMS,
                    FINAL_KUAI_DI_SET, FINAL_CANG_CHU_SET,
                    bonded_match_dict, center_match_dict, overseas_match_dict
                )

                # 金额校验提示
                if not amount_valid:
                    msg = f"  ⚠️ 金额校验失败：源数据¥{source_total:,.2f}，输出¥{output_total:,.2f}，差额¥{amount_diff:,.2f}"
                    log(msg)
                    error_log.append(msg)
                else:
                    # 校验成功也提示
                    log(f"  ✓ 金额校验通过：源数据¥{source_total:,.2f}，输出¥{output_total:,.2f}")

                # 配置缺失警告
                if warnings:
                    for warning in warnings:
                        warning_msg = f"{sheet_name}：{warning}"
                        warning_list.append(warning_msg)
                        log(f"  ⚠️ {sheet_name} {warning}，数据已合并到保税仓")

                # 未匹配处理
                if match_status == 'not_matched':
                    msg = f"  - ⚠️ {sheet_name} 未在匹配表中找到对应记录"
                    log(msg)
                    error_log.append(msg)
                    warning_row = [''] * 16
                    warning_row[4] = sheet_name
                    warning_row[15] = '未找到匹配关系'
                    excel_rows.append(warning_row)
                    continue

                # 生成行数据逻辑（新增row_total参数）
                def make_row(info, report, other_fees):
                    r = []
                    r.append(info.get('货主编码', ''))
                    r.append(info.get('发货仓库名称', ''))
                    r.append(info.get('业务月份', ''))
                    r.append('计提账单')
                    r.append(info.get('品牌名称', sheet_name))
                    r.append(info.get('店铺名称', ''))
                    r.append(info.get('物流对接人', ''))
                    r.append('菜鸟')
                    r.append('菜鸟')
                    r.append(report.get('快递发货单量', 0))
                    r.append(report.get('快递费', 0))
                    r.append(abs(report.get('服务赔付', 0)))
                    r.append(abs(report.get('货值赔付', 0)))
                    r.append(abs(report.get('仓内增值费', 0)))
                    r.append(0)
                    
                    # 计算本行总计（与源数据对应）
                    # = 快递费（含基础服务费）+ 仓储作业费 + 其他费用 - 服务赔付 - 货值赔付
                    # r[9]=单量 r[10]=快递费 r[11]=服务赔付abs r[12]=货值赔付abs r[13]=仓内增值费abs r[14]=综合税
                    row_total = (
                        r[10] +  # 快递费（已包含基础服务费）
                        r[13] +  # 仓内增值费
                        sum(other_fees.values()) -  # 其他费用
                        r[11] -  # 服务赔付（减去）
                        r[12]    # 货值赔付（减去）
                    )
                    
                    # 备注栏：其他费用 + 本行总计
                    if other_fees:
                        fees_str = ','.join([f'{k}{v:.2f}' for k, v in other_fees.items()])
                        remark = f'菜鸟金掌柜 | 其他费用:{fees_str} | 本行总计:¥{row_total:,.2f}'
                    else:
                        remark = f'菜鸟金掌柜 | 本行总计:¥{row_total:,.2f}'
                    r.append(remark)
                    return r

                # 输出保税仓
                if report_bonded and bonded_match_info:
                    excel_rows.append(make_row(bonded_match_info, report_bonded, bonded_other_fees))
                    processed_count += 1
                
                # 输出中心仓（仅当有配置时）
                if has_center and report_center and center_match_info:
                    excel_rows.append(make_row(center_match_info, report_center, center_other_fees))
                    processed_count += 1
                
                # 输出海外仓（仅当有配置时）
                if has_overseas and report_overseas and overseas_match_info:
                    excel_rows.append(make_row(overseas_match_info, report_overseas, overseas_other_fees))
                    processed_count += 1
                    
            except Exception as e:
                log(f"  - ❌ 出错: {e}")
                error_log.append(f"{sheet_name}: {e}")

        # 汇总警告信息
        if warning_list:
            log("")
            log("⚠️ 检测到以下品牌存在未配置的仓库数据，已合并到保税仓：")
            for warning in warning_list:
                log(f"  - {warning}")


        # 输出
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
            
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_file = os.path.join(self.output_dir, f"快递物流台账批量导入模板_{timestamp}.xlsx")
        
        # 读取模板获取表头
        try:
             df_template = pd.read_excel(self.template_file, header=1)
             columns = df_template.columns.tolist()
        except Exception as e:
             return False, f"读取模板表头失败: {e}"
        
        try:
            with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
                pd.DataFrame(excel_rows, columns=columns).to_excel(writer, index=False, startrow=1)
        except Exception as e:
            return False, f"生成 Excel 失败: {e}"
        
        return True, f"清洗完成！处理了 {processed_count} 个条目。输出: {os.path.basename(out_file)}"
