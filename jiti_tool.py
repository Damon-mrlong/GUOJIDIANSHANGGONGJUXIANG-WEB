#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
本地计提工具 V2
处理品牌费用数据，生成快递发货物流台账Excel报告
"""

import pandas as pd
import os
from datetime import datetime
from pathlib import Path


def load_rule_data(rule_file):
    """
    加载规则文件，提取快递费和仓内增值费的费用项

    Args:
        rule_file: 规则文件路径

    Returns:
        tuple: (file_kuai_di_set, file_cang_chu_set)
    """
    df_rule = pd.read_excel(rule_file)

    # 根据规则文件提取快递费和仓内增值费的费用项
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


def load_brand_match_data(match_file):
    """
    加载品牌仓库对接人匹配关系表

    Args:
        match_file: 匹配关系文件路径

    Returns:
        tuple: (bonded_match_dict, center_match_dict)
            - bonded_match_dict: 保税仓匹配字典 {品牌名: 匹配记录}
            - center_match_dict: 中心仓匹配字典 {品牌名: 匹配记录}
    """
    df_match = pd.read_excel(match_file)

    # 获取列名（中文）
    cols = df_match.columns.tolist()
    brand_col = cols[3]  # 品牌列
    is_center_col = cols[-1]  # 是否中心仓列

    bonded_match_dict = {}
    center_match_dict = {}

    for _, row in df_match.iterrows():
        brand = row[brand_col]
        is_center = row[is_center_col]

        match_record = {
            '货主编码': row[cols[0]],
            '发货仓库名称': row[cols[1]],
            '业务月份': row[cols[2]],
            '品牌名称': row[brand_col],
            '店铺名称': row[cols[4]],
            '物流对接人': row[cols[5]],
            '快递供应商物流对接人': row[cols[6]]
        }

        if is_center == '是':
            center_match_dict[brand] = match_record
        else:
            bonded_match_dict[brand] = match_record

    return bonded_match_dict, center_match_dict


def build_classification_sets(file_kuai_di_set, file_cang_chu_set):
    """
    构建最终的分类集合

    Args:
        file_kuai_di_set: 文件规则中的快递费集合
        file_cang_chu_set: 文件规则中的仓内增值费集合

    Returns:
        tuple: (INDEPENDENT_ITEMS, FINAL_KUAI_DI_SET, FINAL_CANG_CHU_SET)
    """
    # 独立指标项（最高优先级）
    INDEPENDENT_ITEMS = {'货值赔付', '服务赔付'}

    # 手动规则（硬编码，用于覆盖或补充文件规则）
    MANUAL_KUAI_DI_SET = {'指定效期残出库费', '经济上门', '优质上门'}
    MANUAL_CANG_CHU_SET = {
        '防尘袋安装费', '质检拒收费', '拆预包费', '隐形眼镜品类操作费'
    }

    # 最终分类集合
    FINAL_KUAI_DI_SET = file_kuai_di_set | MANUAL_KUAI_DI_SET
    FINAL_CANG_CHU_SET = file_cang_chu_set | MANUAL_CANG_CHU_SET

    return INDEPENDENT_ITEMS, FINAL_KUAI_DI_SET, FINAL_CANG_CHU_SET


def calculate_metrics(df, INDEPENDENT_ITEMS, FINAL_KUAI_DI_SET, FINAL_CANG_CHU_SET):
    """
    计算各项指标

    Args:
        df: 品牌DataFrame
        INDEPENDENT_ITEMS: 独立指标项集合
        FINAL_KUAI_DI_SET: 快递费集合
        FINAL_CANG_CHU_SET: 仓内增值费集合

    Returns:
        dict: 包含6个指标的字典
    """
    report_dict = {}

    # 1. 快递发货单量（基础服务费的主单行数量总和）
    report_dict['快递发货单量'] = df[
        df['费用项'] == '基础服务费'
    ]['主单行数量'].sum()

    # 2. 货值赔付（独立计算）
    report_dict['货值赔付'] = df[
        df['费用项'] == '货值赔付'
    ]['支付金额'].sum()

    # 3. 服务赔付（独立计算）
    report_dict['服务赔付'] = df[
        df['费用项'] == '服务赔付'
    ]['支付金额'].sum()

    # 4. 快递费（排除独立指标项）
    kuai_di_filter = (
        df['费用项'].isin(FINAL_KUAI_DI_SET) &
        ~df['费用项'].isin(INDEPENDENT_ITEMS)
    )
    report_dict['快递费'] = df[kuai_di_filter]['支付金额'].sum()

    # 5. 仓内增值费（排除独立指标项）
    cang_chu_filter = (
        df['费用项'].isin(FINAL_CANG_CHU_SET) &
        ~df['费用项'].isin(INDEPENDENT_ITEMS)
    )
    report_dict['仓内增值费'] = df[cang_chu_filter]['支付金额'].sum()

    # 6. 其他待纳入统计的费用
    processed_items = INDEPENDENT_ITEMS | FINAL_KUAI_DI_SET | FINAL_CANG_CHU_SET
    df_other = df[~df['费用项'].isin(processed_items)]
    other_summary = df_other.groupby('费用项')['支付金额'].sum()
    report_dict['其他待纳入统计的费用'] = other_summary[other_summary != 0].to_dict()

    # 7. 金额汇总（排除其他费用）
    # 计算除"其他待纳入统计的费用"外的所有金额汇总
    main_fees_sum = (
        report_dict['货值赔付'] +
        report_dict['服务赔付'] +
        report_dict['快递费'] +
        report_dict['仓内增值费']
    )
    report_dict['金额汇总(不含其他费用)'] = main_fees_sum

    return report_dict


def process_brand_data(df_full, brand_name, INDEPENDENT_ITEMS,
                       FINAL_KUAI_DI_SET, FINAL_CANG_CHU_SET,
                       bonded_match_dict, center_match_dict):
    """
    处理单个品牌数据

    Args:
        df_full: 完整品牌DataFrame
        brand_name: 品牌名称（Sheet名称）
        INDEPENDENT_ITEMS: 独立指标项集合
        FINAL_KUAI_DI_SET: 快递费集合
        FINAL_CANG_CHU_SET: 仓内增值费集合
        bonded_match_dict: 保税仓匹配字典
        center_match_dict: 中心仓匹配字典

    Returns:
        tuple: (total_payment, bonded_report, center_report, do_split,
                bonded_match_info, center_match_info, match_status,
                bonded_other_fees, center_other_fees)
    """
    # 获取匹配信息
    bonded_match_info = bonded_match_dict.get(brand_name)
    center_match_info = center_match_dict.get(brand_name)

    match_status = 'matched' if bonded_match_info else 'not_matched'
    # 检查G列是否存在
    if '物流商品' not in df_full.columns:
        raise ValueError(f"缺少'物流商品'列")

    # 数据清洗
    df_full = df_full.copy()
    df_full['费用项'] = df_full['费用项'].str.strip().fillna('')
    df_full['物流商品'] = df_full['物流商品'].str.strip().fillna('')

    # 清理数值列（去除CNY符号、货币符号、逗号等）
    for col in ['支付金额', '主单行数量']:
        # 先转换为字符串
        df_full[col] = df_full[col].astype(str)
        # 去除CNY符号、货币符号、逗号、前后空格
        df_full[col] = df_full[col].str.replace(r'CNY|¥|\$|,', '', regex=True).str.strip()
        # 转换为数值
        df_full[col] = pd.to_numeric(df_full[col], errors='coerce').fillna(0)

    # 汇总金额
    total_payment = df_full['支付金额'].sum()

    # G列前置条件检查
    center_g_value = "商家-保税中心仓"
    center_rows = df_full[df_full['物流商品'] == center_g_value]
    do_split = '基础服务费' in center_rows['费用项'].values

    # 数据拆分与计算
    if do_split:
        # 拆分：保税仓 + 中心仓
        df_bonded = df_full[df_full['物流商品'] != center_g_value]
        df_center = center_rows

        report_bonded = calculate_metrics(
            df_bonded, INDEPENDENT_ITEMS, FINAL_KUAI_DI_SET, FINAL_CANG_CHU_SET
        )
        report_center = calculate_metrics(
            df_center, INDEPENDENT_ITEMS, FINAL_KUAI_DI_SET, FINAL_CANG_CHU_SET
        )
    else:
        # 不拆分：只有保税仓
        df_bonded = df_full
        report_bonded = calculate_metrics(
            df_bonded, INDEPENDENT_ITEMS, FINAL_KUAI_DI_SET, FINAL_CANG_CHU_SET
        )
        report_center = None

    # 获取其他费用信息
    bonded_other_fees = report_bonded.get('其他待纳入统计的费用', {}) if report_bonded else {}
    center_other_fees = report_center.get('其他待纳入统计的费用', {}) if report_center else {}

    return total_payment, report_bonded, report_center, do_split, \
           bonded_match_info, center_match_info, match_status, \
           bonded_other_fees, center_other_fees


def generate_excel_output(excel_data_list, output_file):
    """
    生成Excel输出文件

    Args:
        excel_data_list: Excel数据行列表
        output_file: 输出文件路径
    """
    # 读取模板以获取正确的列名
    df_template = pd.read_excel('规则/快递发货物流台账批量导入模板.xlsx', header=1)
    columns = df_template.columns.tolist()

    # 创建DataFrame
    df_output = pd.DataFrame(excel_data_list, columns=columns)

    # 写入Excel文件
    df_output.to_excel(output_file, index=False, header=True)
    print(f"   Excel报告已生成: {output_file}")


def format_report_output(brand_name, total_payment, report_bonded,
                         report_center, do_split):
    """
    格式化报告输出

    Args:
        brand_name: 品牌名称（Sheet名称）
        total_payment: 总金额
        report_bonded: 保税仓报告
        report_center: 中心仓报告
        do_split: 是否拆分

    Returns:
        str: 格式化的报告文本
    """
    lines = []

    # 处理品牌名称（去掉Sheet名称中的"-中心仓"或"-拆分"后缀）
    if brand_name.endswith('-中心仓'):
        base_brand_name = brand_name.removesuffix('-中心仓')  # 安全去掉后缀
        bonded_name = f"{base_brand_name}-保税仓"
        center_name = f"{base_brand_name}-中心仓"
    elif brand_name.endswith('-拆分'):
        base_brand_name = brand_name.removesuffix('-拆分')
        bonded_name = f"{base_brand_name}-保税仓"
        center_name = f"{base_brand_name}-中心仓"
    else:
        base_brand_name = brand_name
        bonded_name = f"{brand_name}-保税仓"
        center_name = f"{brand_name}-中心仓"

    # 格式化数值（2位小数）
    def fmt(value):
        if isinstance(value, (int, float)):
            return round(float(value), 2)
        return value

    # 总金额行
    lines.append(f"{base_brand_name} 汇总金额: {fmt(total_payment)}")
    lines.append("")

    # 保税仓报告
    lines.append(bonded_name)
    for key, value in report_bonded.items():
        if key == '其他待纳入统计的费用':
            lines.append(f"  {key}:")
            for item, amount in value.items():
                lines.append(f"    {item}: {fmt(amount)}")
        else:
            lines.append(f"  {key}: {fmt(value)}")
    lines.append("")

    # 中心仓报告（如果拆分）
    if do_split and report_center:
        lines.append(center_name)
        for key, value in report_center.items():
            if key == '其他待纳入统计的费用':
                lines.append(f"  {key}:")
                for item, amount in value.items():
                    lines.append(f"    {item}: {fmt(amount)}")
            else:
                lines.append(f"  {key}: {fmt(value)}")
        lines.append("")

    return "\n".join(lines)


def main():
    """
    主函数
    """
    print("=" * 80)
    print("本地计提工具 V2 启动【开发者：小龙】")
    print("生成快递发货物流台账Excel报告")
    print("=" * 80)

    # 文件路径
    data_file = "源文件/计提用.xlsx"
    rule_file = "规则/菜鸟费用计提规则V2.xlsx"
    match_file = "规则/品牌仓库对接人匹配关系.xlsx"
    output_dir = Path("输出")
    output_dir.mkdir(exist_ok=True)

    # 错误日志
    error_log = []

    try:
        # 加载规则数据
        print("\n1. 加载规则数据...")
        file_kuai_di_set, file_cang_chu_set = load_rule_data(rule_file)
        INDEPENDENT_ITEMS, FINAL_KUAI_DI_SET, FINAL_CANG_CHU_SET = \
            build_classification_sets(file_kuai_di_set, file_cang_chu_set)

        print(f"   独立指标项: {INDEPENDENT_ITEMS}")
        print(f"   快递费费用项数量: {len(FINAL_KUAI_DI_SET)}")
        print(f"   仓内增值费费用项数量: {len(FINAL_CANG_CHU_SET)}")

        # 加载品牌匹配关系
        print("\n2. 加载品牌匹配关系...")
        bonded_match_dict, center_match_dict = load_brand_match_data(match_file)
        print(f"   保税仓品牌数量: {len(bonded_match_dict)}")
        print(f"   中心仓品牌数量: {len(center_match_dict)}")

        # 处理所有品牌Sheet
        print("\n3. 处理品牌数据...")
        xl_file = pd.ExcelFile(data_file)

        excel_rows = []
        processed_count = 0

        for sheet_name in xl_file.sheet_names:
            print(f"\n   处理品牌: {sheet_name}")
            df = pd.read_excel(xl_file, sheet_name=sheet_name)

            # 检查是否有数据
            if df.empty or len(df) == 0:
                print(f"     警告: {sheet_name} 无数据，跳过")
                continue

            try:
                # 处理品牌数据
                total_payment, report_bonded, report_center, do_split, \
                bonded_match_info, center_match_info, match_status, \
                bonded_other_fees, center_other_fees = \
                    process_brand_data(
                        df, sheet_name, INDEPENDENT_ITEMS,
                        FINAL_KUAI_DI_SET, FINAL_CANG_CHU_SET,
                        bonded_match_dict, center_match_dict
                    )

                # 如果没有匹配信息，生成警告行
                if match_status == 'not_matched':
                    warning_row = [''] * 16
                    warning_row[4] = sheet_name  # E列 品牌名称
                    warning_row[15] = '未找到匹配关系'  # P列 备注
                    excel_rows.append(warning_row)
                    error_msg = f"{sheet_name} 未在匹配表中找到对应记录"
                    error_log.append(error_msg)
                    print(f"     警告: {error_msg}")
                    processed_count += 1
                    continue

                # 生成保税仓数据行
                if report_bonded and bonded_match_info:
                    row = []
                    # A 货主编码
                    row.append(bonded_match_info['货主编码'] if pd.notna(bonded_match_info['货主编码']) else '')
                    # B 发货仓库名称
                    row.append(bonded_match_info['发货仓库名称'] if pd.notna(bonded_match_info['发货仓库名称']) else '')
                    # C 业务月份
                    row.append(bonded_match_info['业务月份'] if pd.notna(bonded_match_info['业务月份']) else '')
                    # D 账单类型 - 固定"计提账单"
                    row.append('计提账单')
                    # E 品牌名称
                    row.append(bonded_match_info['品牌名称'] if pd.notna(bonded_match_info['品牌名称']) else sheet_name)
                    # F 店铺名称
                    row.append(bonded_match_info['店铺名称'] if pd.notna(bonded_match_info['店铺名称']) else '')
                    # G 物流对接人
                    row.append(bonded_match_info['物流对接人'] if pd.notna(bonded_match_info['物流对接人']) else '')
                    # H 快递供应商 - 固定"菜鸟"
                    row.append('菜鸟')
                    # I 税金供应商 - 固定"菜鸟"
                    row.append('菜鸟')
                    # J 发货单量
                    row.append(report_bonded.get('快递发货单量', 0))
                    # K 快递费
                    row.append(report_bonded.get('快递费', 0))
                    # L 仓储服务赔付费 - 取绝对值
                    row.append(abs(report_bonded.get('服务赔付', 0)))
                    # M 仓储货值赔付费 - 取绝对值
                    row.append(abs(report_bonded.get('货值赔付', 0)))
                    # N 仓储作业费（仓内增值费）
                    row.append(report_bonded.get('仓内增值费', 0))
                    # O 跨境电商综合税 - 输出0
                    row.append(0)
                    # P 备注 - 动态构建
                    if bonded_other_fees:
                        # 构建其他费用信息
                        other_fees_str = ','.join([f'{item}{amount:.2f}' for item, amount in bonded_other_fees.items()])
                        remark = f'菜鸟金掌柜 | 其他费用:{other_fees_str}'
                    else:
                        remark = '菜鸟金掌柜 | 无其他待纳入统计费用'
                    row.append(remark)

                    excel_rows.append(row)
                    processed_count += 1

                # 生成中心仓数据行（如果需要拆分）
                if do_split and report_center and center_match_info:
                    row = []
                    # A 货主编码
                    row.append(center_match_info['货主编码'] if pd.notna(center_match_info['货主编码']) else '')
                    # B 发货仓库名称
                    row.append(center_match_info['发货仓库名称'] if pd.notna(center_match_info['发货仓库名称']) else '')
                    # C 业务月份
                    row.append(center_match_info['业务月份'] if pd.notna(center_match_info['业务月份']) else '')
                    # D 账单类型 - 固定"计提账单"
                    row.append('计提账单')
                    # E 品牌名称
                    row.append(center_match_info['品牌名称'] if pd.notna(center_match_info['品牌名称']) else sheet_name)
                    # F 店铺名称
                    row.append(center_match_info['店铺名称'] if pd.notna(center_match_info['店铺名称']) else '')
                    # G 物流对接人
                    row.append(center_match_info['物流对接人'] if pd.notna(center_match_info['物流对接人']) else '')
                    # H 快递供应商 - 固定"菜鸟"
                    row.append('菜鸟')
                    # I 税金供应商 - 固定"菜鸟"
                    row.append('菜鸟')
                    # J 发货单量
                    row.append(report_center.get('快递发货单量', 0))
                    # K 快递费
                    row.append(report_center.get('快递费', 0))
                    # L 仓储服务赔付费 - 取绝对值
                    row.append(abs(report_center.get('服务赔付', 0)))
                    # M 仓储货值赔付费 - 取绝对值
                    row.append(abs(report_center.get('货值赔付', 0)))
                    # N 仓储作业费（仓内增值费）
                    row.append(report_center.get('仓内增值费', 0))
                    # O 跨境电商综合税 - 输出0
                    row.append(0)
                    # P 备注 - 动态构建
                    if center_other_fees:
                        # 构建其他费用信息
                        other_fees_str = ','.join([f'{item}{amount:.2f}' for item, amount in center_other_fees.items()])
                        remark = f'菜鸟金掌柜 | 其他费用:{other_fees_str}'
                    else:
                        remark = '菜鸟金掌柜 | 无其他待纳入统计费用'
                    row.append(remark)

                    excel_rows.append(row)
                    processed_count += 1

                print(f"     完成 - 汇总金额: {total_payment}")
                print(f"     拆分状态: {'是' if do_split else '否'}")

            except ValueError as e:
                error_msg = f"{sheet_name} 错误：{str(e)}"
                error_log.append(error_msg)
                print(f"     错误: {error_msg}")

        # 生成Excel文件
        print("\n4. 生成Excel文件...")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_output_file = output_dir / f"快递发货物流台账批量导入模板_{timestamp}.xlsx"

        generate_excel_output(excel_rows, excel_output_file)

        # 生成错误日志
        if error_log:
            error_log_file = output_dir / f"错误日志_{timestamp}.txt"
            with open(error_log_file, 'w', encoding='utf-8') as f:
                f.write("错误日志\n")
                f.write("=" * 80 + "\n")
                f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("=" * 80 + "\n\n")
                for error in error_log:
                    f.write(error + "\n")
            print(f"   错误日志已生成: {error_log_file}")
        else:
            print("   无错误日志")

        print("\n" + "=" * 80)
        print("处理完成！")
        print(f"共处理 {processed_count} 个品牌Sheet")
        print(f"共生成 {len(excel_rows)} 行Excel数据")
        print(f"错误数量: {len(error_log)}")
        print(f"输出文件: {excel_output_file}")
        print("=" * 80)

    except Exception as e:
        print(f"\n严重错误: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
