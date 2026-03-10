import pandas as pd
import os
from datetime import datetime
from typing import Tuple, Dict, List, Optional

# ================= 列索引常量定义 =================
# 关键字段
COL_BRAND = 4                    # 品牌 (列E)
COL_DESTINATION = 6              # 目的仓 (列G)

# 起运地费用
COL_ORIGIN_LTL_UNIT = 9          # 起运地-拼车单价 (列J)
COL_ORIGIN_FTL_3T = 10           # 起运地-3T整车 (列K)
COL_ORIGIN_FTL_5T = 11           # 起运地-5T整车 (列L)
COL_ORIGIN_FTL_8T = 12           # 起运地-8T整车 (列M)
COL_ORIGIN_FTL_10T = 13          # 起运地-10T整车 (列N)
COL_ORIGIN_FTL_20T = 14          # 起运地-20T整车 (列O)
COL_ORIGIN_LTL_MIN = 15          # 起运地-拼车最低收费 (列P)

# 空运段费用
COL_AIR_FIXED = 16               # 按票收费all in (列Q)
COL_AIR_UNIT = 17                # 按KG收费all in (列R)
COL_AIR_MIN = 18                 # 空运最低收费 (列S)
COL_AIR_DOC = 19                 # 抽单费 (列T)
COL_AIR_DEPOT_UNIT = 20          # 货站提货费单价 (列U)
COL_AIR_DEPOT_MIN = 21           # 提货最低收费 (列V)

# 目的地费用
COL_DEST_LTL_UNIT = 22           # 目的地-拼车单价 (列W)
COL_DEST_FTL_3T = 23             # 目的地-3T整车 (列X)
COL_DEST_FTL_5T = 24             # 目的地-5T整车 (列Y)
COL_DEST_FTL_8T = 25             # 目的地-8T整车 (列Z)
COL_DEST_FTL_10T = 26            # 目的地-10T整车 (列AA)
# COL_DEST_FTL_20GP = 27         # 跳过！业务规则要求忽略20GP
COL_DEST_FTL_40GP = 28           # 目的地-40GP (列AC) - 用于17-20托板
COL_DEST_FTL_45GP = 29           # 目的地-45GP (列AD) - 用于21-23托板
COL_DEST_FTL_58GP = 30           # 目的地-58GP (列AE) - 用于24-32托板
COL_DEST_LTL_MIN = 31            # 目的地-最低收费 (列AF)
COL_DEST_DELIVERY = 32           # 目的地-拼车提送货费 (列AG)


class QuoteCalculatorEngine:
    """物流报价计算引擎"""
    
    def __init__(self, workspace_root=None):
        """初始化计算引擎
        
        Args:
            workspace_root: 工作区根目录
        """
        self.workspace_root = workspace_root
        self.df_rules = None  # 规则数据DataFrame
        
        if workspace_root:
            self.rules_dir = os.path.join(workspace_root, "报价计算文件夹")
            self.rules_file = os.path.join(self.rules_dir, "空运报价费用规则.xlsx")
            self.output_dir = os.path.join(workspace_root, "输出汇总文件夹")
        else:
            self.rules_dir = None
            self.rules_file = None
            self.output_dir = None
    
    def set_workspace(self, root):
        """设置工作区"""
        self.workspace_root = root
        
        # 优先在"规则文件"查找，其次在"报价计算文件夹"
        possible_paths = [
            os.path.join(root, "规则文件", "空运报价费用规则.xlsx"),
            os.path.join(root, "报价计算文件夹", "空运报价费用规则.xlsx")
        ]
        
        self.rules_file = None
        for p in possible_paths:
            if os.path.exists(p):
                self.rules_file = p
                break
        
        # 如果都没找到，默认由load_rules处理（会报错）
        if not self.rules_file:
            self.rules_file = possible_paths[0]
            
        self.output_dir = os.path.join(root, "输出汇总文件夹")
    
    def load_rules(self, log_callback=None) -> Tuple[bool, str]:
        """加载费用规则文件"""
        def log(msg):
            if log_callback:
                log_callback(msg)
        
        if not self.rules_file or not os.path.exists(self.rules_file):
            # 尝试重新寻找
            if self.workspace_root:
                self.set_workspace(self.workspace_root)
            
            if not self.rules_file or not os.path.exists(self.rules_file):
                return False, f"规则文件不存在: {self.rules_file}"
        
        try:
            log(f"正在加载规则文件: {os.path.basename(self.rules_file)}")
            self.df_rules = pd.read_excel(self.rules_file, header=None, skiprows=2)
            self.df_rules = self.df_rules.fillna(0)
            log(f"成功加载 {len(self.df_rules)} 条规则记录")
            return True, "规则加载成功"
        except Exception as e:
            return False, f"加载规则文件失败: {str(e)}"

    # ... get_brands, get_destinations, find_row, calc_origin_ltl, calc_origin_ftl (保持不变) ...

    def get_brands(self) -> List[str]:
        """获取所有品牌列表"""
        if self.df_rules is None:
            return []
        brands = self.df_rules.iloc[:, COL_BRAND].astype(str).str.strip()
        brands = brands[brands != '0']
        brands = brands[brands != '']
        return sorted(brands.unique().tolist())
    
    def get_destinations(self) -> List[str]:
        """获取所有目的地仓库列表"""
        if self.df_rules is None:
            return []
        destinations = self.df_rules.iloc[:, COL_DESTINATION].astype(str).str.strip()
        destinations = destinations[destinations != '0']
        destinations = destinations[destinations != '']
        return sorted(destinations.unique().tolist())
    
    def find_row(self, brand: str, destination: str) -> Optional[pd.Series]:
        """查找匹配的规则行"""
        if self.df_rules is None:
            return None
        brand_match = self.df_rules.iloc[:, COL_BRAND].astype(str).str.strip() == brand
        dest_match = self.df_rules.iloc[:, COL_DESTINATION].astype(str).str.strip() == destination
        matched_rows = self.df_rules[brand_match & dest_match]
        if len(matched_rows) == 0:
            return None
        return matched_rows.iloc[0]
    
    def calc_origin_ltl(self, row: pd.Series, weight: float) -> float:
        """计算启用国提货费（LTL）"""
        unit_price = float(row.iloc[COL_ORIGIN_LTL_UNIT])
        min_fee = float(row.iloc[COL_ORIGIN_LTL_MIN])
        calculated = weight * unit_price
        return max(calculated, min_fee)
    
    def calc_origin_ftl(self, row: pd.Series, pallets: int) -> float:
        """计算启用国提货费（FTL）"""
        if pallets <= 0: return 0.0
        if 1 <= pallets <= 3: return float(row.iloc[COL_ORIGIN_FTL_3T])
        elif 4 <= pallets <= 8: return float(row.iloc[COL_ORIGIN_FTL_5T])
        elif 9 <= pallets <= 13: return float(row.iloc[COL_ORIGIN_FTL_8T])
        elif 14 <= pallets <= 16: return float(row.iloc[COL_ORIGIN_FTL_10T])
        elif 17 <= pallets <= 33: return float(row.iloc[COL_ORIGIN_FTL_20T])
        else:
            cost_33 = float(row.iloc[COL_ORIGIN_FTL_20T])
            remaining = pallets - 33
            return cost_33 + self.calc_origin_ftl(row, remaining)

    def calc_air(self, row: pd.Series, weight: float) -> float:
        """计算空运费
        公式: Max(重量 * R列, S列) + Q列
        """
        unit = float(row.iloc[COL_AIR_UNIT])    # R列
        min_fee = float(row.iloc[COL_AIR_MIN])  # S列
        fixed = float(row.iloc[COL_AIR_FIXED])  # Q列
        
        return max(weight * unit, min_fee) + fixed

    def calc_dest_port(self, row: pd.Series, weight: float) -> float:
        """计算目的港费用
        公式: Max(重量 * U列, V列) + T列
        """
        unit = float(row.iloc[COL_AIR_DEPOT_UNIT]) # U列
        min_fee = float(row.iloc[COL_AIR_DEPOT_MIN]) # V列
        fixed = float(row.iloc[COL_AIR_DOC])       # T列 (原Doc Fee)
        
        return max(weight * unit, min_fee) + fixed

    def calc_dest_ltl(self, row: pd.Series, weight: float) -> float:
        """计算港到仓费用（LTL）
        公式: Max(重量 * W列 + AG列, AF列)
        """
        unit_price = float(row.iloc[COL_DEST_LTL_UNIT]) # W列
        delivery_addon = float(row.iloc[COL_DEST_DELIVERY]) # AG列
        min_fee = float(row.iloc[COL_DEST_LTL_MIN]) # AF列
        
        calculated = (weight * unit_price) + delivery_addon
        return max(calculated, min_fee)
    
    def calc_dest_ftl(self, row: pd.Series, pallets: int) -> float:
        """计算港到仓费用（FTL）"""
        if pallets <= 0: return 0.0
        if 1 <= pallets <= 3: return float(row.iloc[COL_DEST_FTL_3T])
        elif 4 <= pallets <= 8: return float(row.iloc[COL_DEST_FTL_5T])
        elif 9 <= pallets <= 13: return float(row.iloc[COL_DEST_FTL_8T])
        elif 14 <= pallets <= 16: return float(row.iloc[COL_DEST_FTL_10T])
        elif 17 <= pallets <= 20: return float(row.iloc[COL_DEST_FTL_40GP])
        elif 21 <= pallets <= 23: return float(row.iloc[COL_DEST_FTL_45GP])
        elif 24 <= pallets <= 32: return float(row.iloc[COL_DEST_FTL_58GP])
        else:
            cost_32 = float(row.iloc[COL_DEST_FTL_58GP])
            remaining = pallets - 32
            return cost_32 + self.calc_dest_ftl(row, remaining)
    
    def calculate(self, brand: str, destination: str, weight: float, pallets: int, 
                  log_callback=None) -> Tuple[bool, str, Optional[Dict]]:
        """执行报价计算"""
        def log(msg):
            if log_callback: log_callback(msg)
        
        if self.df_rules is None:
            success, msg = self.load_rules(log_callback)
            if not success: return False, msg, None
        
        if weight <= 0 or pallets <= 0:
            return False, "重量和托板数必须大于0", None
        
        row = self.find_row(brand, destination)
        if row is None:
            return False, f"未找到匹配的规则记录（品牌: {brand}, 目的地: {destination}）", None
        
        log(f"开始计算，重量={weight}KG, 托板数={pallets}")
        
        # 1. 启用国提货费
        origin_ltl = self.calc_origin_ltl(row, weight)
        origin_ftl = self.calc_origin_ftl(row, pallets)
        
        # 2. 空运费
        air_fee = self.calc_air(row, weight)
        
        # 3. 目的港费用
        dest_port_fee = self.calc_dest_port(row, weight)
        
        # 4. 港到仓费用
        dest_wh_ltl = self.calc_dest_ltl(row, weight)
        dest_wh_ftl = self.calc_dest_ftl(row, pallets)
        
        scenarios = [
            # 场景1: 零担/零担
            {
                'name': 'LTL/LTL',
                'origin': origin_ltl,
                'air': air_fee,
                'dest_port': dest_port_fee,
                'dest_wh': dest_wh_ltl,
                'total': origin_ltl + air_fee + dest_port_fee + dest_wh_ltl
            },
            # 场景2: 整车/整车
            {
                'name': 'FTL/FTL',
                'origin': origin_ftl,
                'air': air_fee,
                'dest_port': dest_port_fee,
                'dest_wh': dest_wh_ftl,
                'total': origin_ftl + air_fee + dest_port_fee + dest_wh_ftl
            },
            # 场景3: 零担/整车 (起运地零担 + 目的地整车)
            {
                'name': 'LTL/FTL',
                'origin': origin_ltl,
                'air': air_fee,
                'dest_port': dest_port_fee,
                'dest_wh': dest_wh_ftl,
                'total': origin_ltl + air_fee + dest_port_fee + dest_wh_ftl
            },
            # 场景4: 整车/零担 (起运地整车 + 目的地零担)
            {
                'name': 'FTL/LTL',
                'origin': origin_ftl,
                'air': air_fee,
                'dest_port': dest_port_fee,
                'dest_wh': dest_wh_ltl,
                'total': origin_ftl + air_fee + dest_port_fee + dest_wh_ltl
            }
        ]
        
        min_scenario = min(scenarios, key=lambda x: x['total'])
        max_scenario = max(scenarios, key=lambda x: x['total'])
        
        log(f"计算完成！最优方案: {min_scenario['name']} (¥{min_scenario['total']:.2f})")
        
        results = {
            'scenarios': scenarios,
            'min_scenario': min_scenario,
            'max_scenario': max_scenario,
            'brand': brand,
            'destination': destination,
            'weight': weight,
            'pallets': pallets
        }
        
        return True, "计算成功", results
    
    def export_results(self, results: Dict, log_callback=None) -> Tuple[bool, str]:
        """导出计算结果到Excel"""
        def log(msg):
            if log_callback: log_callback(msg)
        
        if not self.output_dir: return False, "输出目录未设置"
        if not os.path.exists(self.output_dir): os.makedirs(self.output_dir)
        
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_file = os.path.join(self.output_dir, f"报价计算结果_{timestamp}.xlsx")
            
            output_data = []
            for scenario in results['scenarios']:
                output_data.append({
                    '场景': scenario['name'].replace('LTL', '零担').replace('FTL', '整车'),
                    '启用国提货费': f"¥{scenario['origin']:.2f}",
                    '空运费': f"¥{scenario['air']:.2f}",
                    '目的港费用': f"¥{scenario['dest_port']:.2f}",
                    '港到仓费用': f"¥{scenario['dest_wh']:.2f}",
                    '总费用': f"¥{scenario['total']:.2f}",
                    '是否最优': '✓' if scenario == results['min_scenario'] else ''
                })
            
            df_output = pd.DataFrame(output_data)
            
            with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
                df_output.to_excel(writer, sheet_name='场景对比', index=False)
            
            log(f"结果已导出: {os.path.basename(out_file)}")
            return True, out_file
        
        except Exception as e:
            return False, f"导出失败: {str(e)}"
