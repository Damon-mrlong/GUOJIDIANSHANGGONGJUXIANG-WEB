import flet as ft
import os
import json
import threading
import asyncio
from datetime import datetime
from auto_audit import FinanceAuditEngine
from bill_parser import BillParserEngine
from upload_template import UploadTemplateEngine
from quote_calculator import QuoteCalculatorEngine

# ================= 0. 核心配置 =================
APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(APP_DIR, "app_settings.json")

# 工作区结构定义
WORKSPACE_STRUCTURE = {
    "规则文件": ["audit_rules.csv", "菜鸟费用计提规则V2.xlsx", "品牌仓库对接人匹配关系.xlsx", "快递发货物流台账批量导入模板.xlsx", "upload_mapping.csv", "price_database.xlsx", "空运报价费用规则.xlsx"],
    "计提实际账单核对": ["计提台账文件夹", "实际台账文件夹"],
    "快递计提台账处理": [],  # 文件直接放在此文件夹下，不需要子文件夹
    "OMS上传台账模板生成": [],
    "报价计算文件夹": [],
    "输出汇总文件夹": []
}

# ================= 1. 后端逻辑：工作区管理 =================
def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except: pass
    return {}

def save_config(workspace_root):
    cfg = {"workspace_root": workspace_root}
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=4)

def migrate_folders(root_path):
    """迁移旧文件夹名称到中文名称"""
    if not root_path or not os.path.exists(root_path): return
    
    # 1. 重命名顶级目录
    top_map = {
        "00_Config": "规则文件",
        "01_Finance_Audit": "计提实际账单核对",
        "02_Bill_Parser": "快递计提台账处理",
        "03_Upload_Template": "OMS上传台账模板生成",
        "04_Quote_Audit": "报价计算文件夹",
        "报价稽核文件夹": "报价计算文件夹",  # 迁移旧名称
        "计提实际账单核对文件夹": "计提实际账单核对",
        "快递计提台账处理文件夹": "快递计提台账处理",
        "上传模板生成文件夹": "OMS上传台账模板生成"
    }
    
    for old, new in top_map.items():
        old_p = os.path.join(root_path, old)
        new_p = os.path.join(root_path, new)
        if os.path.exists(old_p) and not os.path.exists(new_p):
            try:
                os.rename(old_p, new_p)
                print(f"Migrated: {old} -> {new}")
            except Exception as e:
                print(f"Failed to migrate {old}: {e}")

    # 2. 重命名子目录 (此时顶级目录已是中文)
    sub_map = {
        "计提实际账单核对": {
            "Inputs_Accrual": "计提台账文件夹",
            "Inputs_Actual": "实际台账文件夹"
        },
        "快递计提台账处理": {
            # 不再需要子文件夹，文件直接放在此目录下
        },
        "OMS上传台账模板生成": {},
        "报价计算文件夹": {}
    }
    
    for parent, subs in sub_map.items():
        parent_path = os.path.join(root_path, parent)
        if not os.path.exists(parent_path): continue
        
        for old_sub, new_sub in subs.items():
            old_sub_p = os.path.join(parent_path, old_sub)
            new_sub_p = os.path.join(parent_path, new_sub)
            if os.path.exists(old_sub_p) and not os.path.exists(new_sub_p):
                try:
                    os.rename(old_sub_p, new_sub_p)
                    print(f"Migrated sub: {old_sub} -> {new_sub}")
                except: pass

def init_workspace(root_path):
    if not root_path:
        return False, "路径为空"
    
    try:
        if not os.path.exists(root_path):
            os.makedirs(root_path)
        
        # 执行自动迁移
        migrate_folders(root_path)

        created_log = []
        for folder, children in WORKSPACE_STRUCTURE.items():
            folder_path = os.path.join(root_path, folder)
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
                created_log.append(f"创建目录: {folder}")
            
            if isinstance(children, list):
                for child in children:
                    if "." in child: 
                         pass 
                    else:
                        sub_path = os.path.join(folder_path, child)
                        if not os.path.exists(sub_path):
                            os.makedirs(sub_path)
                            created_log.append(f"创建子目录: {folder}/{child}")
                            
        return True, f"初始化完成。{len(created_log)} 个新目录 created."
    except Exception as e:
        return False, f"初始化失败: {str(e)}"

# ================= 2. UI组件 =================

# 品牌色定义
COLOR_PRIMARY = "#d32f2f"     # 品牌红
COLOR_BG_MAIN = "#f5f5f5"     # 浅灰底色
COLOR_BG_SIDEBAR = "#ffffff"  # 侧边栏白底
COLOR_TEXT_PRIMARY = "#333333"
COLOR_TEXT_SECONDARY = "#757575"

class LogControl(ft.Container):
    def __init__(self):
        super().__init__()
        self.log_view = ft.ListView(expand=True, spacing=2, auto_scroll=True)
        # 头部：标题 + 状态灯
        self.status_indicator = ft.Container(width=10, height=10, border_radius=5, bgcolor="green")
        
        self.content = ft.Column([
            ft.Container(
                content=ft.Row([
                    ft.Text("系统运行日志", size=12, weight="bold", color="#555555"),
                    ft.Container(expand=True),
                    self.status_indicator,
                    ft.Text("运行正常", size=10, color="green")
                ]),
                padding=ft.padding.only(bottom=5)
            ),
            ft.Divider(height=1, color="#eeeeee"),
            self.log_view
        ], spacing=0)
        
        self.height = 160
        self.bgcolor = "white"
        self.padding = 10
        self.border_radius = 0 # 底部贴边，不需要圆角
        self.border = ft.border.only(top=ft.border.BorderSide(1, "#e0e0e0"))

    def log(self, msg, color="#333333"):
        timestamp = datetime.now().strftime('%H:%M:%S')
        # 如果 msg 包含 "失败" or "错误", 变红
        text_color = color
        if "失败" in msg or "错误" in msg or "Error" in msg:
            text_color = "red"
        elif "成功" in msg or "完成" in msg:
            text_color = "green"
            
        self.log_view.controls.append(
            ft.Text(f"[{timestamp}] {msg}", color=text_color, size=11, font_family="Microsoft YaHei")
        )
        self.update()

def open_output_folder(workspace_root):
    """自动打开输出汇总文件夹"""
    if not workspace_root: return
    out_dir = os.path.join(workspace_root, "输出汇总文件夹")
    if os.path.exists(out_dir):
        os.startfile(out_dir)

# ================= 3. 模块视图 =================

class BaseModuleView(ft.Container):
    """所有功能模块的基类，提供白色卡片样式"""
    def __init__(self):
        super().__init__()
        self.padding = 30
        self.bgcolor = "white"
        self.border_radius = 8
        self.shadow = ft.BoxShadow(blur_radius=10, color="#08000000", offset=ft.Offset(0, 2))
        self.expand = True 

class FinanceAuditView(BaseModuleView):
    def __init__(self, page, logger, get_workspace_fn):
        super().__init__()
        self.page_ref = page
        self.logger = logger
        self.get_workspace_fn = get_workspace_fn
        self.accrual_path_text = ft.Text("未选择文件", size=12, color=COLOR_TEXT_SECONDARY, overflow=ft.TextOverflow.ELLIPSIS)
        self.actual_path_text = ft.Text("未选择文件", size=12, color=COLOR_TEXT_SECONDARY, overflow=ft.TextOverflow.ELLIPSIS)
        self.map_path_text = ft.Text("可选", size=12, color=COLOR_TEXT_SECONDARY, overflow=ft.TextOverflow.ELLIPSIS)
        self.progress_bar = ft.ProgressBar(color=COLOR_PRIMARY, bgcolor="#ffcdd2", visible=False)
        self.run_btn = ft.ElevatedButton("开始自动对账", icon="play_arrow", on_click=self.run_audit, bgcolor=COLOR_PRIMARY, color="white", height=40)

        self.pick_accrual = ft.FilePicker(on_result=self.accrual_folder_picked)
        self.pick_actual = ft.FilePicker(on_result=self.actual_folder_picked)
        self.pick_map = ft.FilePicker(on_result=lambda e: self.file_picked(e, self.map_path_text))
        self.page_ref.overlay.extend([self.pick_accrual, self.pick_actual, self.pick_map])

        self.content = ft.Column(
            [
                ft.Row([
                    ft.Icon("account_balance", color=COLOR_PRIMARY, size=28),
                    ft.Text("资金对账", size=22, weight="bold", color=COLOR_TEXT_PRIMARY),
                ]),
                ft.Divider(height=10, color="transparent"),
                
                ft.Text("第1步：选择计提台账文件夹", weight="bold", color=COLOR_TEXT_PRIMARY),
                ft.Row([
                    ft.ElevatedButton("选择文件夹", icon="folder_open", on_click=lambda _: self.pick_accrual.get_directory_path(), bgcolor="white", color=COLOR_TEXT_PRIMARY, elevation=0, style=ft.ButtonStyle(side=ft.BorderSide(1, "#e0e0e0"))),
                    ft.Container(content=self.accrual_path_text, padding=10, bgcolor="#f9f9f9", border_radius=4, expand=True)
                ]),
                ft.Container(height=5),
                
                ft.Text("第2步：选择实际账单文件夹", weight="bold", color=COLOR_TEXT_PRIMARY),
                ft.Row([
                    ft.ElevatedButton("选择文件夹", icon="folder_open", on_click=lambda _: self.pick_actual.get_directory_path(), bgcolor="white", color=COLOR_TEXT_PRIMARY, elevation=0, style=ft.ButtonStyle(side=ft.BorderSide(1, "#e0e0e0"))),
                    ft.Container(content=self.actual_path_text, padding=10, bgcolor="#f9f9f9", border_radius=4, expand=True)
                ]),
                ft.Container(height=5),

                ft.Text("第3步：选择仓库映射表（可选）", weight="bold", color=COLOR_TEXT_PRIMARY),
                ft.Row([
                    ft.ElevatedButton("选择文件", icon="folder_open", on_click=lambda _: self.pick_map.pick_files(allowed_extensions=["xlsx", "xls", "csv"]), bgcolor="white", color=COLOR_TEXT_PRIMARY, elevation=0, style=ft.ButtonStyle(side=ft.BorderSide(1, "#e0e0e0"))),
                    ft.Container(content=self.map_path_text, padding=10, bgcolor="#f9f9f9", border_radius=4, expand=True)
                ]),
                
                ft.Container(height=10),
                self.progress_bar,
                ft.Row([self.run_btn], alignment=ft.MainAxisAlignment.END)
            ]
        )

    def accrual_folder_picked(self, e: ft.FilePickerResultEvent):
        if e.path:
            import glob
            files = glob.glob(os.path.join(e.path, "*.xlsx")) + glob.glob(os.path.join(e.path, "*.xls"))
            files = [f for f in files if not os.path.basename(f).startswith("~$")]
            if files:
                self.accrual_path_text.value = f"已选择: {len(files)} 个文件"
                self.accrual_path_text.color = "green"
                self.accrual_path_text.data = e.path  # 存储路径
            else:
                self.accrual_path_text.value = "文件夹为空"
                self.accrual_path_text.color = "orange"
            self.accrual_path_text.update()
    
    def actual_folder_picked(self, e: ft.FilePickerResultEvent):
        if e.path:
            import glob
            files = glob.glob(os.path.join(e.path, "*.xlsx")) + glob.glob(os.path.join(e.path, "*.xls"))
            files = [f for f in files if not os.path.basename(f).startswith("~$")]
            if files:
                self.actual_path_text.value = f"已选择: {len(files)} 个文件"
                self.actual_path_text.color = "green"
                self.actual_path_text.data = e.path
            else:
                self.actual_path_text.value = "文件夹为空"
                self.actual_path_text.color = "orange"
            self.actual_path_text.update()
    
    def file_picked(self, e: ft.FilePickerResultEvent, text_control):
        if e.files and len(e.files) > 0:
            text_control.value = e.files[0].path
            text_control.color = "black"
            text_control.update()

    def try_auto_detect(self):
        workspace = self.get_workspace_fn()
        if not workspace: return
        try:
            engine = FinanceAuditEngine(workspace)
            found = engine.auto_detect_files()
            acc_count = found.get("accrual_count", 0)
            act_count = found.get("actual_count", 0)
            
            self.accrual_path_text.value = f"自动探测: {acc_count} 个计提文件"
            self.accrual_path_text.color = "green" if acc_count > 0 else "orange"

            self.actual_path_text.value = f"自动探测: {act_count} 个实际文件"
            self.actual_path_text.color = "green" if act_count > 0 else "orange"

            if found.get("map"):
                self.map_path_text.value = found["map"]
                self.map_path_text.color = "black"
            self.update()
        except: pass

    def run_audit(self, e):
        workspace = self.get_workspace_fn()
        if not workspace:
            self.page_ref.show_snack_bar(ft.SnackBar(ft.Text("请先配置工作区！"), bgcolor="red"))
            return
        
        # 检查是自动识别还是手动选择
        accrual_dir = None
        actual_dir = None
        
        if "自动探测" in self.accrual_path_text.value:
            # 使用自动识别的路径
            accrual_dir = os.path.join(workspace, "计提实际账单核对", "计提台账文件夹")
            actual_dir = os.path.join(workspace, "计提实际账单核对", "实际台账文件夹")
        elif hasattr(self.accrual_path_text, 'data') and self.accrual_path_text.data:
            # 使用手动选择的文件夹
            accrual_dir = self.accrual_path_text.data
            actual_dir = self.actual_path_text.data if hasattr(self.actual_path_text, 'data') else None
        
        if not accrual_dir or not actual_dir:
            self.page_ref.open(ft.SnackBar(ft.Text("请先选择计提和实际账单文件夹！"), bgcolor="red"))
            return

        map_path = self.map_path_text.value
        if "可选" in map_path: map_path = None
        
        self.run_btn.disabled = True
        self.run_btn.text = "计算中..."
        self.progress_bar.visible = True
        self.update()
        
        self.logger.log(">>> 启动自动对账任务", "blue")

        self.logger.log(">>> 启动自动对账任务", "blue")

        # async_log 不需要定义在外面，直接在 wrappers 里用
        
        def task():
            engine = FinanceAuditEngine(workspace)
            
            # 如果是手动选择的文件夹，设置自定义目录
            if accrual_dir != os.path.join(workspace, "计提实际账单核对", "计提台账文件夹"):
                engine.set_custom_dirs(accrual_dir, actual_dir)
            
            def sync_log_wrapper(m):
                async def _log_task():
                    self.logger.log(m)
                self.page_ref.run_task(_log_task)
            
            # 注意：engine.run_audit 是同步阻塞的，所以在这里调用没问题
            success, msg = engine.run_audit(
                map_path=map_path,
                log_callback=sync_log_wrapper
            )
            
            if success:
                async def _success_task():
                    self.logger.log(f"任务成功: {msg}")
                    self.page_ref.open(ft.SnackBar(ft.Text("对账完成！"), bgcolor="green"))
                    open_output_folder(workspace)
                self.page_ref.run_task(_success_task)
            else:
                async def _fail_task():
                    self.logger.log(f"任务失败: {msg}")
                    self.page_ref.open(ft.SnackBar(ft.Text(f"失败: {msg}"), bgcolor="red"))
                self.page_ref.run_task(_fail_task)

            async def reset_ui():
                self.run_btn.disabled = False
                self.run_btn.text = "开始自动对账"
                self.progress_bar.visible = False
                self.update()
            
            self.page_ref.run_task(reset_ui)

        threading.Thread(target=task, daemon=True).start()

class BillParserView(BaseModuleView):
    def __init__(self, page, logger, get_workspace_fn):
        super().__init__()
        self.page_ref = page
        self.logger = logger
        self.get_workspace_fn = get_workspace_fn
        self.source_path_text = ft.Text("未选择文件", size=12, color=COLOR_TEXT_SECONDARY)
        self.progress_bar = ft.ProgressBar(color="#4caf50", bgcolor="#c8e6c9", visible=False)
        self.run_btn = ft.ElevatedButton("开始清洗账单", icon="play_arrow", on_click=self.run_parser, bgcolor="#4caf50", color="white", height=40)

        self.pick_source = ft.FilePicker(on_result=lambda e: self.file_picked(e, self.source_path_text))
        self.page_ref.overlay.append(self.pick_source)

        self.content = ft.Column(
            [
                ft.Row([
                    ft.Icon("cleaning_services", color="#4caf50", size=28),
                    ft.Text("快递计提", size=22, weight="bold", color=COLOR_TEXT_PRIMARY),
                ]),
                ft.Divider(height=30, color="transparent"),
                
                ft.Container(
                    content=ft.Row([
                        ft.Icon("info_outline", color="#1565c0", size=16),
                        ft.Text("根据菜鸟费用计提规则，将金掌柜后台费用分类并调整为快递计提台账上传模板格式", size=13, color="#1565c0")
                    ]),
                    bgcolor="#e3f2fd", padding=10, border_radius=4
                ),
                ft.Container(height=15),

                ft.Text("选择供应商原始账单（Excel）", weight="bold", color=COLOR_TEXT_PRIMARY),
                ft.Row([
                    ft.ElevatedButton("选择文件", icon="folder_open", on_click=lambda _: self.pick_source.pick_files(allowed_extensions=["xlsx", "xls"]), bgcolor="white", color=COLOR_TEXT_PRIMARY, elevation=0, style=ft.ButtonStyle(side=ft.BorderSide(1, "#e0e0e0"))),
                    ft.Container(content=self.source_path_text, padding=10, bgcolor="#f9f9f9", border_radius=4, expand=True)
                ]),
                
                ft.Container(height=30),
                self.progress_bar,
                ft.Row([self.run_btn], alignment=ft.MainAxisAlignment.END)
            ]
        )
        
        # 尝试自动加载源文件
        self.try_auto_detect()

    def file_picked(self, e: ft.FilePickerResultEvent, text_control):
        if e.files and len(e.files) > 0:
            text_control.value = e.files[0].path
            text_control.color = "black"
            text_control.update()
    
    def try_auto_detect(self):
        """尝试自动识别快递计提文件夹中的源文件"""
        workspace = self.get_workspace_fn()
        if not workspace:
            return
        
        try:
            # 修正路径：文件直接在"快递计提台账处理"文件夹下，没有"源文件"子文件夹
            source_dir = os.path.join(workspace, "快递计提台账处理")
            if not os.path.exists(source_dir):
                return
            
            import glob
            files = glob.glob(os.path.join(source_dir, "*.xlsx")) + glob.glob(os.path.join(source_dir, "*.xls"))
            files = [f for f in files if not os.path.basename(f).startswith("~$")]
            
            if not files:
                return
            
            # 方案A：选择最新文件
            latest_file = max(files, key=os.path.getmtime)
            self.source_path_text.value = latest_file
            self.source_path_text.color = "green"
            # 移除 self.update() 以避免初始化时的异常
            
            # 使用 safe_log 替代直接 logger.log
            try:
                self.logger.log(f"✅ 自动加载: {os.path.basename(latest_file)}", "green")
            except:
                print(f"[快递计提] ✅ 自动加载: {os.path.basename(latest_file)}")
        except Exception as e:
            # 打印错误以便调试
            print(f"[快递计提] 自动加载失败: {e}")

    def run_parser(self, e):
        workspace = self.get_workspace_fn()
        if not workspace:
            self.page_ref.show_snack_bar(ft.SnackBar(ft.Text("请先配置工作区！"), bgcolor="red"))
            return
        
        src_path = self.source_path_text.value
        if "未选择" in src_path:
            self.page_ref.open(ft.SnackBar(ft.Text("源文件路径无效"), bgcolor="red"))
            return

        self.run_btn.disabled = True
        self.run_btn.text = "处理中..."
        self.progress_bar.visible = True
        self.update()
        
        self.logger.log(">>> 启动账单清洗", "blue")

        self.logger.log(">>> 启动账单清洗", "blue")

        def task():
            engine = BillParserEngine(workspace)
            
            def sync_log_wrapper(m):
                async def _log_task():
                    self.logger.log(m)
                self.page_ref.run_task(_log_task)

            success, msg = engine.run_parser(src_path, log_callback=sync_log_wrapper)
            
            if success:
                async def _success_task():
                    self.logger.log(msg)
                    self.page_ref.open(ft.SnackBar(ft.Text("清洗完成！"), bgcolor="green"))
                    open_output_folder(workspace)
                self.page_ref.run_task(_success_task)
            else:
                async def _fail_task():
                    self.logger.log(f"失败: {msg}")
                    self.page_ref.show_snack_bar(ft.SnackBar(ft.Text(f"失败: {msg}"), bgcolor="red"))
                self.page_ref.run_task(_fail_task)

            async def reset_ui():
                self.run_btn.disabled = False
                self.run_btn.text = "开始清洗账单"
                self.progress_bar.visible = False
                self.update()
            self.page_ref.run_task(reset_ui)

        threading.Thread(target=task, daemon=True).start()

class UploadTemplateView(BaseModuleView):
    """OMS上传模板生成视图 - 占位，待台账上传功能完善后开发"""
    def __init__(self, page, logger, get_workspace_fn):
        super().__init__()
        self.page_ref = page
        self.logger = logger
        self.get_workspace_fn = get_workspace_fn

        self.content = ft.Column(
            [
                ft.Row([
                    ft.Icon("cloud_upload", color="#ff9800", size=28),
                    ft.Text("上传模版生成", size=22, weight="bold", color=COLOR_TEXT_PRIMARY),
                ]),
                ft.Divider(height=30, color="transparent"),
                
                # 占位提示信息
                ft.Container(
                    content=ft.Column([
                        ft.Icon("construction", size=80, color="#ffa726"),
                        ft.Container(height=20),
                        ft.Text(
                            "功能开发中", 
                            size=24, 
                            weight="bold", 
                            color=COLOR_TEXT_PRIMARY,
                            text_align=ft.TextAlign.CENTER
                        ),
                        ft.Container(height=10),
                        ft.Text(
                            "待台账上传功能完善后开发",
                            size=16,
                            color=COLOR_TEXT_SECONDARY,
                            text_align=ft.TextAlign.CENTER
                        ),
                        ft.Container(height=20),
                        ft.Container(
                            content=ft.Text(
                                "此模块将用于生成OMS系统所需的上传台账模板",
                                size=14,
                                color="#757575",
                                text_align=ft.TextAlign.CENTER
                            ),
                            padding=20,
                            bgcolor="#f5f5f5",
                            border_radius=8
                        )
                    ], 
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    alignment=ft.MainAxisAlignment.CENTER),
                    alignment=ft.alignment.center,
                    expand=True
                )
            ],
            expand=True
        )

class QuoteCalculatorView(BaseModuleView):
    def __init__(self, page, logger, get_workspace_fn):
        super().__init__()
        self.page_ref = page
        self.logger = logger
        self.get_workspace_fn = get_workspace_fn
        self.engine = None
        self.all_brands = [] # 缓存所有品牌用于搜索

        # 规则文件路径显示
        self.rules_file_text = ft.Text("未加载规则文件", size=12, color=COLOR_TEXT_SECONDARY, overflow=ft.TextOverflow.ELLIPSIS)
        
        # 文件选择器
        self.pick_rules = ft.FilePicker(on_result=self.rules_file_picked)
        self.page_ref.overlay.append(self.pick_rules)
        
        # --- UI 组件定义 ---
        
        # 1. 品牌搜索与选择
        # 使用 Textfield 过滤 Dropdown
        self.brand_search = ft.TextField(
            label="搜索品牌",
            hint_text="输入关键字",
            on_change=self.filter_brands,
            expand=1,
            height=40,
            text_size=13,
            content_padding=10,
            bgcolor="white",
            disabled=True
        )
        
        self.brand_dropdown = ft.Dropdown(
            label="选择品牌",
            hint_text="请先加载规则",
            expand=2,
            text_size=13,
            disabled=True,
            bgcolor="white"
        )
        
        self.destination_dropdown = ft.Dropdown(
            label="目的地仓库",
            hint_text="请先加载规则",
            expand=2,
            text_size=13,
            disabled=True,
            bgcolor="white"
        )
        
        # 2. 数值输入
        self.weight_input = ft.TextField(
            label="重量 (KG)",
            width=150,
            height=40,
            text_size=13,
            content_padding=10,
            keyboard_type=ft.KeyboardType.NUMBER,
            bgcolor="white"
        )
        
        self.pallets_input = ft.TextField(
            label="托板数",
            width=150,
            height=40,
            text_size=13,
            content_padding=10,
            keyboard_type=ft.KeyboardType.NUMBER,
            bgcolor="white"
        )
        
        self.run_btn = ft.ElevatedButton(
            "计算报价", 
            icon="calculate", 
            on_click=self.run_calculation, 
            bgcolor="#1976d2", 
            color="white", 
            height=40,
            style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=4))
        )
        
        # 3. 结果展示 - 自适应布局
        self.result_grid = ft.GridView(
            runs_count=4,
            max_extent=300,
            child_aspect_ratio=1.3,
            spacing=10,
            run_spacing=10,
        )
        
        self.result_container = ft.Container(
            content=self.result_grid,
            padding=ft.padding.only(top=10),
            visible=False
        )
        
        self.best_tip = ft.Container(
            content=ft.Row([
                ft.Icon("check_circle", color="green", size=16),
                ft.Text("", size=14, color="green", weight="bold")
            ]),
            visible=False,
            padding=5,
            bgcolor="#e8f5e9",
            border_radius=4
        )

        self.progress_bar = ft.ProgressBar(color="#1976d2", bgcolor="#bbdefb", visible=False, height=2)

        # --- 页面布局 ---
        self.content = ft.Column(
            [
                # 标题栏
                ft.Row([
                    ft.Icon("calculate", color="#1976d2", size=24),
                    ft.Text("报价计算", size=20, weight="bold", color=COLOR_TEXT_PRIMARY),
                    ft.Container(expand=True),
                    ft.OutlinedButton(
                        "加载规则", 
                        icon="folder_open", 
                        on_click=lambda _: self.pick_rules.pick_files(allowed_extensions=["xlsx", "xls"]),
                        height=30,
                        style=ft.ButtonStyle(padding=10)
                    )
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                
                # 状态与提示
                ft.Row([
                    ft.Container(
                        content=ft.Row([
                            ft.Icon("info_outline", color="#666", size=14),
                            self.rules_file_text
                        ]),
                    )
                ]),
                
                ft.Divider(height=10, color="transparent"),
                
                # 输入区域 (Row 1)
                ft.Row([
                    self.brand_search,
                    self.brand_dropdown,
                    self.destination_dropdown
                ], spacing=10),
                
                ft.Divider(height=5, color="transparent"),
                
                # 输入区域 (Row 2)
                ft.Row([
                    self.weight_input,
                    self.pallets_input,
                    ft.Container(expand=True),
                    self.run_btn
                ], spacing=10),
                
                self.progress_bar,
                
                # 结果区域
                self.best_tip,
                self.result_container,
            ],
            scroll=ft.ScrollMode.AUTO,
            spacing=5
        )
        
        # 尝试自动加载规则 (初始化时先不调用，以免工作区未配置时报错)
        # self.auto_load_rules() 
        pass

    def filter_brands(self, e):
        """根据输入过滤品牌列表"""
        if not self.all_brands: return
        
        keyword = e.control.value.lower().strip()
        if not keyword:
            filtered = self.all_brands
        else:
            filtered = [b for b in self.all_brands if keyword in b.lower()]
            
        self.brand_dropdown.options = [ft.dropdown.Option(b) for b in filtered]
        # 如果过滤后只有一个选项，不需要自动选中，因为用户可能还在输入
        if len(filtered) == 0:
            self.brand_dropdown.options = [ft.dropdown.Option("无匹配品牌")]
        else:
            # 保持之前的选择，如果仍在列表中
            if self.brand_dropdown.value not in filtered:
                self.brand_dropdown.value = None
                
        self.brand_dropdown.update()
    
    def rules_file_picked(self, e: ft.FilePickerResultEvent):
        """用户选择规则文件后的回调"""
        if e.files and len(e.files) > 0:
            rules_path = e.files[0].path
            self.load_rules_from_path(rules_path)
    
    def load_rules_from_path(self, rules_path: str):
        """从指定路径加载规则文件"""
        try:
            self.safe_log(f"正在加载规则文件: {os.path.basename(rules_path)}", "blue")
            
            # 创建临时引擎并加载规则
            temp_engine = QuoteCalculatorEngine()
            temp_engine.rules_file = rules_path
            
            success, msg = temp_engine.load_rules(log_callback=lambda m: self.safe_log(m))
            
            if success:
                # 加载成功，更新引擎
                self.engine = temp_engine
                workspace = self.get_workspace_fn()
                if workspace:
                    self.engine.set_workspace(workspace)
                
                # 更新文件路径显示
                self.rules_file_text.value = rules_path
                self.rules_file_text.color = "green"
                
                # 填充品牌下拉选项
                self.all_brands = self.engine.get_brands() # 更新缓存
                self.brand_dropdown.options = [ft.dropdown.Option(b) for b in self.all_brands]
                self.brand_dropdown.disabled = False
                self.brand_dropdown.hint_text = f"选择品牌 ({len(self.all_brands)}个)"
                self.brand_search.disabled = False # 启用搜索输入框
                
                # 填充目的地下拉选项
                destinations = self.engine.get_destinations()
                self.destination_dropdown.options = [ft.dropdown.Option(d) for d in destinations]
                self.destination_dropdown.disabled = False
                self.destination_dropdown.hint_text = f"选择目的地 ({len(destinations)}个)"
                
                self.safe_log(f"✅ 规则加载成功: {len(self.all_brands)}个品牌, {len(destinations)}个目的地", "green")
                
                # 安全地更新UI
                try:
                    self.update()
                except (AssertionError, AttributeError):
                    # 控件还未添加到page，跳过更新
                    pass
                
                # 安全地显示提示
                try:
                    self.page_ref.open(ft.SnackBar(ft.Text(f"加载成功！{len(brands)}个品牌, {len(destinations)}个目的地"), bgcolor="green"))
                except:
                    pass
            else:
                self.safe_log(f"❌ 规则加载失败: {msg}", "red")
                self.rules_file_text.value = f"加载失败: {msg}"
                self.rules_file_text.color = "red"
                
                # 安全地更新UI
                try:
                    self.update()
                except (AssertionError, AttributeError):
                    pass
                
                try:
                    self.page_ref.open(ft.SnackBar(ft.Text(f"加载失败: {msg}"), bgcolor="red"))
                except:
                    pass
        except Exception as e:
            error_msg = f"加载规则时出错: {str(e)}"
            self.safe_log(f"❌ {error_msg}", "red")
            self.rules_file_text.value = error_msg
            self.rules_file_text.color = "red"
            
            # 安全地更新UI
            try:
                self.update()
            except (AssertionError, AttributeError):
                pass
            
            try:
                self.page_ref.open(ft.SnackBar(ft.Text(error_msg), bgcolor="red"))
            except:
                pass
    
    def safe_log(self, msg, color="black"):
        """安全地记录日志，避免在初始化时出错"""
        try:
            self.logger.log(msg, color)
        except (AssertionError, AttributeError):
            # 如果logger还未添加到page，使用print
            print(f"[{color}] {msg}")
    
    def auto_load_rules(self):
        """尝试自动加载默认规则文件"""
        workspace = self.get_workspace_fn()
        if not workspace:
            # 仅仅静默返回，不要记录红色错误，以免启动时惊扰用户
            # self.safe_log("⚠️ 工作区未配置，请手动选择规则文件", "orange")
            return
            
        # 优先查找 "规则文件" 目录
        default_path = os.path.join(workspace, "规则文件", "空运报价费用规则.xlsx")
        
        # 备选路径
        backup_path = os.path.join(workspace, "报价计算文件夹", "空运报价费用规则.xlsx")

        target_path = None
        if os.path.exists(default_path):
            target_path = default_path
        elif os.path.exists(backup_path):
            target_path = backup_path
            
        if target_path and os.path.exists(target_path):
            self.load_rules_from_path(target_path)
        else:
            self.safe_log(f"⚠️ 未找到默认规则文件: {default_path}", "orange")
            self.safe_log("请点击'选择规则文件'按钮手动选择", "orange")
    
    def run_calculation(self, e):
        """执行报价计算"""
        workspace = self.get_workspace_fn()
        if not workspace:
            self.page_ref.open(ft.SnackBar(ft.Text("请先配置工作区！"), bgcolor="red"))
            return
        
        # 验证输入
        if not self.brand_dropdown.value:
            self.page_ref.open(ft.SnackBar(ft.Text("请选择品牌！"), bgcolor="orange"))
            return
        
        if not self.destination_dropdown.value:
            self.page_ref.open(ft.SnackBar(ft.Text("请选择目的地！"), bgcolor="orange"))
            return
        
        try:
            weight = float(self.weight_input.value or 0)
            pallets = int(self.pallets_input.value or 0)
            
            if weight <= 0:
                self.page_ref.open(ft.SnackBar(ft.Text("重量必须大于0！"), bgcolor="orange"))
                return
            
            if pallets <= 0:
                self.page_ref.open(ft.SnackBar(ft.Text("托板数必须大于0！"), bgcolor="orange"))
                return
        except ValueError:
            self.page_ref.open(ft.SnackBar(ft.Text("请输入有效的数字！"), bgcolor="red"))
            return
        
        # 禁用按钮，显示进度条
        self.run_btn.disabled = True
        self.run_btn.text = "计算中..."
        self.progress_bar.visible = True
        self.result_container.visible = False
        self.best_tip.visible = False
        self.update()
        
        self.logger.log(">>> 启动报价计算", "blue")
        
        # 在新线程中执行计算
        def task():
            if not self.engine:
                self.engine = QuoteCalculatorEngine(workspace)
                self.engine.load_rules()
            
            def sync_log_wrapper(m):
                async def _log_task():
                    self.logger.log(m)
                self.page_ref.run_task(_log_task)
            
            # 执行计算
            success, msg, results = self.engine.calculate(
                brand=self.brand_dropdown.value,
                destination=self.destination_dropdown.value,
                weight=weight,
                pallets=pallets,
                log_callback=sync_log_wrapper
            )
            
            if success:
                async def _success_task():
                    self.logger.log(f"计算成功: {msg}")
                    self.display_results(results)
                    self.page_ref.open(ft.SnackBar(ft.Text("计算完成！"), bgcolor="green"))
                    # 移除自动导出Excel和打开文件夹，提升体验
                self.page_ref.run_task(_success_task)
            else:
                async def _fail_task():
                    self.logger.log(f"计算失败: {msg}", "red")
                    self.page_ref.open(ft.SnackBar(ft.Text(f"失败: {msg}"), bgcolor="red"))
                self.page_ref.run_task(_fail_task)
            
            async def reset_ui():
                self.run_btn.disabled = False
                self.run_btn.text = "开始计算报价"
                self.progress_bar.visible = False
                self.update()
            
            self.page_ref.run_task(reset_ui)
        
        threading.Thread(target=task, daemon=True).start()
    
    def display_results(self, results):
        """显示计算结果"""
        # 清空旧卡片
        self.result_grid.controls.clear()
        
        # 术语映射
        name_map = {
            'LTL/LTL': '零担/零担',
            'FTL/FTL': '整车/整车',
            'LTL/FTL': '零担/整车',
            'FTL/LTL': '整车/零担'
        }
        
        # 创建场景卡片
        for scenario in results['scenarios']:
            is_best = scenario == results['min_scenario']
            is_worst = scenario == results['max_scenario']
            
            # 设置卡片颜色
            if is_best:
                card_bgcolor = "#e8f5e9"  # 浅绿色
                border_color = "#4caf50"  # 绿色边框
                title_color = "#2e7d32"
            elif is_worst:
                card_bgcolor = "#ffebee"  # 浅红色
                border_color = "#e57373"  # 红色边框
                title_color = "#c62828"
            else:
                card_bgcolor = "#ffffff"  # 白色
                border_color = "#e0e0e0"  # 灰色边框
                title_color = "#333333"
            
            # 场景名称显示
            scenario_name = name_map.get(scenario['name'], scenario['name'])
            if is_best:
                scenario_name += " ✓ (推荐)"
            
            # 创建卡片
            card = ft.Container(
                content=ft.Column([
                    # 场景名称（独立行）
                    ft.Text(
                        scenario_name,
                        size=14,
                        weight="bold",
                        color=title_color,
                        overflow=ft.TextOverflow.ELLIPSIS
                    ),
                    # 总价（独立行，可换行）
                    ft.Text(
                        f"¥{scenario['total']:,.2f}",
                        size=16,
                        weight="bold",
                        color=title_color,
                        overflow=ft.TextOverflow.VISIBLE
                    ),
                    ft.Divider(height=1, color=border_color),
                    # 费用明细 - 紧凑布局
                    ft.Column([
                        ft.Row([
                            ft.Text("启运国提货", size=10, color="#757575", expand=1),
                            ft.Text(f"¥{scenario['origin']:,.0f}", size=10, weight="bold", overflow=ft.TextOverflow.ELLIPSIS)
                        ], spacing=4),
                        ft.Row([
                            ft.Text("空运费", size=10, color="#757575", expand=1),
                            ft.Text(f"¥{scenario['air']:,.0f}", size=10, weight="bold", overflow=ft.TextOverflow.ELLIPSIS)
                        ], spacing=4),
                        ft.Row([
                            ft.Text("目的港费用", size=10, color="#757575", expand=1),
                            ft.Text(f"¥{scenario.get('dest_port', 0):,.0f}", size=10, weight="bold", overflow=ft.TextOverflow.ELLIPSIS)
                        ], spacing=4),
                        ft.Row([
                            ft.Text("港到仓费用", size=10, color="#757575", expand=1),
                            ft.Text(f"¥{scenario.get('dest_wh', 0):,.0f}", size=10, weight="bold", overflow=ft.TextOverflow.ELLIPSIS)
                        ], spacing=4),
                    ], spacing=3),
                ], spacing=6),
                padding=12,
                bgcolor=card_bgcolor,
                border=ft.border.all(2 if is_best or is_worst else 1, border_color),
                border_radius=8,
            )
            
            self.result_grid.controls.append(card)
        
        # 更新最优方案提示
        best = results['min_scenario']
        best_name = name_map.get(best['name'], best['name'])
        self.best_tip.content.controls[1].value = f"最优方案: {best_name} - 总费用 ¥{best['total']:,.2f}"
        self.best_tip.visible = True
        
        # 显示结果容器
        self.result_container.visible = True
        self.update()


# ================= 4. 主程序 =================
def main(page: ft.Page):
    # --- 页面基础设置 ---
    page.title = "国际电商工具箱"
    page.window_width = 1200
    page.window_height = 800
    page.bgcolor = COLOR_BG_MAIN # 全局浅灰背景
    page.padding = 0
    page.theme_mode = "light" # 强制亮色模式
    
    # 设置资源目录
    page.assets_dir = "assets"

    # --- 状态管理 ---
    current_workspace = None
    selected_nav_index = 0 # 0: Home, 1: Audit, ...
    
    # 组件引用
    logger = LogControl()
    
    def get_current_workspace():
        return current_workspace

    # --- 视图初始化 ---
    view_audit = FinanceAuditView(page, logger, get_current_workspace)
    view_parser = BillParserView(page, logger, get_current_workspace)
    view_upload = UploadTemplateView(page, logger, get_current_workspace)
    view_calculator = QuoteCalculatorView(page, logger, get_current_workspace)

    def create_placeholder_view():
        return ft.Container(
            content=ft.Column([
                ft.Icon(name="waving_hand", size=60, color="#ffb74d"),
                ft.Text("欢迎使用", size=28, weight="bold", color=COLOR_TEXT_PRIMARY),
                ft.Text("请从左侧导航栏选择您需要处理的业务模块", size=14, color=COLOR_TEXT_SECONDARY),
            ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            alignment=ft.alignment.center,
            bgcolor="white",
            border_radius=8,
            shadow=ft.BoxShadow(blur_radius=10, color="#08000000"),
            expand=True
        )

    view_home = create_placeholder_view()

    # 视图字典
    views = {
        0: view_home,
        1: view_audit,
        2: view_parser,
        3: view_calculator,
        4: view_upload,
    }

    # 内容区域容器
    content_area = ft.Container(
        content=views[0], 
        expand=True, 
        padding=20,
    )

    # --- 导航栏逻辑 ---
    
    # 侧边栏按钮组件
    class NavButton(ft.Container):
        def __init__(self, icon_name, text, index, on_click):
            super().__init__()
            self.index = index
            self.on_click_callback = on_click
            self.icon_name = icon_name
            self.text = text
            self.padding = 12
            self.border_radius = 8
            self.ink = True
            self.on_click = self.clicked
            
            # 内部状态
            self.icon_ctl = ft.Icon(icon_name, size=20, color=COLOR_TEXT_SECONDARY)
            self.text_ctl = ft.Text(text, size=14, color=COLOR_TEXT_SECONDARY)
            
            self.content = ft.Row([
                self.icon_ctl,
                ft.Container(width=10),
                self.text_ctl
            ])
            
        def clicked(self, e):
            self.on_click_callback(self.index)
            
        def set_active(self, active):
            if active:
                self.bgcolor = "#ffebee" # 浅红背景
                self.icon_ctl.color = COLOR_PRIMARY
                self.text_ctl.color = COLOR_PRIMARY
                self.text_ctl.weight = "bold"
            else:
                self.bgcolor = None
                self.icon_ctl.color = COLOR_TEXT_SECONDARY
                self.text_ctl.color = COLOR_TEXT_SECONDARY
                self.text_ctl.weight = "normal"
            if self.page:
                self.update()

    nav_btns = []
    
    def switch_tab(index):
        nonlocal selected_nav_index
        selected_nav_index = index
        
        # 更新按钮状态
        for btn in nav_btns:
            btn.set_active(btn.index == index)
            
        # 切换内容
        content_area.content = views.get(index, view_home)
        content_area.update()

    # 创建所有导航按钮
    nav_data = [
        ("home", "主页", 0),
        ("account_balance", "资金对账", 1),
        ("cleaning_services", "快递计提", 2),
        ("calculate", "报价计算", 3),
        ("cloud_upload", "上传模版", 4)
    ]
    
    for icon, txt, idx in nav_data:
        btn = NavButton(icon, txt, idx, switch_tab)
        nav_btns.append(btn)

    # 默认选中第一个
    nav_btns[0].set_active(True)

    # --- 顶部Header逻辑 ---
    
    # 工作区状态显示
    status_pill = ft.Container(
        content=ft.Row([
            ft.Icon(name="circle", size=12, color="orange"),
            ft.Text("未监控工作区", size=12, color=COLOR_TEXT_SECONDARY)
        ], alignment=ft.MainAxisAlignment.CENTER),
        padding=ft.padding.symmetric(horizontal=12, vertical=6),
        bgcolor="#eeeeee",
        border_radius=15,
        on_click=lambda _: folder_picker.get_directory_path() # 点击也可以触发选择
    )

    def update_workspace_status(path):
        nonlocal current_workspace
        if path and os.path.exists(path):
            current_workspace = path
            
            # 更新Pill样式
            status_pill.content.controls[0].color = "green"
            status_pill.content.controls[1].value = f"工作区: {os.path.basename(path)}"
            status_pill.bgcolor = "#e8f5e9"
            status_pill.update()
            
            logger.log(f"已加载工作区: {path}", "green")
            
            success, msg = init_workspace(path)
            if success:
                logger.log(msg, "blue")
                view_audit.try_auto_detect()
                view_parser.try_auto_detect()  # 触发快递计提自动加载
                # 触发报价计算器加载规则
                try:
                    view_calculator.auto_load_rules()
                except Exception as e:
                    logger.log(f"报价计算器规则加载失败: {str(e)}", "orange")
            else:
                logger.log(msg, "red")
        else:
            current_workspace = None
            status_pill.content.controls[0].color = "orange"
            status_pill.content.controls[1].value = "未设置工作区"
            status_pill.bgcolor = "#eeeeee"
            status_pill.update()

    def pick_folder_result(e: ft.FilePickerResultEvent):
        if e.path:
            save_config(e.path)
            update_workspace_status(e.path)

    folder_picker = ft.FilePicker(on_result=pick_folder_result)
    page.overlay.append(folder_picker)

    # Header 组件
    header = ft.Container(
        content=ft.Row([
            ft.Text("国际电商工具箱", size=18, weight="bold", color="#333333"),
            ft.Container(expand=True),
            status_pill,
            ft.IconButton(icon="folder_open", tooltip="切换工作区", on_click=lambda _: folder_picker.get_directory_path()),
        ]),
        height=60,
        bgcolor="white",
        padding=ft.padding.symmetric(horizontal=20),
        border=ft.border.only(bottom=ft.border.BorderSide(1, "#e0e0e0"))
    )

    # --- 侧边栏布局 ---
    sidebar = ft.Container(
        width=220,
        bgcolor="white",
        content=ft.Column([
            # Nav Area
            ft.Container(height=20),
            ft.Column(nav_btns, spacing=5, expand=True), # 按钮列表
            
            # --- 版本信息Footer ---
            ft.Container(
                content=ft.Column([
                    ft.Divider(height=1, color="#eeeeee"),
                    ft.Text("版本信息V5.0.3，作者：小龙", size=10, color=COLOR_TEXT_SECONDARY, text_align=ft.TextAlign.CENTER)
                ], alignment=ft.MainAxisAlignment.CENTER),
                padding=20,
                alignment=ft.alignment.center
            )
        ]),
        border=ft.border.only(right=ft.border.BorderSide(1, "#e0e0e0"))
    )

    # --- 整体布局组装 ---
    
    # 使用 Row 分割 侧边栏 和 右侧区域
    body = ft.Row(
        [
            sidebar,
            ft.Column([ # 右侧: Header + Content + Log
                header,
                content_area,
                logger
            ], expand=True, spacing=0)
        ],
        expand=True,
        spacing=0
    )

    page.add(body)

    # 初始化加载配置
    cfg = load_config()
    if cfg and "workspace_root" in cfg:
        update_workspace_status(cfg["workspace_root"])
    else:
        # 初次使用的欢迎弹窗
        page.dialog = ft.AlertDialog(
            title=ft.Text("欢迎使用"),
            content=ft.Text("请先选择一个文件夹作为【工作区】，系统将自动为您创建标准目录结构。"),
            actions=[
                ft.TextButton("去选择", on_click=lambda _: [setattr(page.dialog, 'open', False), page.update(), folder_picker.get_directory_path()])
            ],
        )
        page.dialog.open = True
        page.update()

ft.app(
    target=main, 
    view=ft.WEB_BROWSER,  # 指定为浏览器模式
    port=8550,              # 使用 80 端口（网页默认端口），如果冲突改 8080
    host="0.0.0.0"        # 关键！0.0.0.0 代表允许局域网内任何电脑访问
)