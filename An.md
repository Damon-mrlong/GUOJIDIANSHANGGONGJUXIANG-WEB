1. 项目背景与目标
本项目旨在开发一款基于 Python Flet 的桌面端应用“供应链效能中台”，用于解决供应链物流部门在资金对账、账单清洗、系统上传及报价稽核中的重复性手工劳动。 核心目标是将分散的 Python 脚本（auto_audit.py, jiti_tool.py）及新需求集成到一个统一的图形化终端（GUI）中，并通过标准化的本地文件系统（Workspace）管理数据流。

2. 技术栈要求
UI 框架: Flet (Python)

数据处理: Pandas, Openpyxl

打包目标: Windows .exe (PyInstaller)

存储方式: 本地文件存储 (Local File System)，无数据库。

配置文件: JSON (应用配置), CSV/Excel (业务规则)。

3. 系统架构与文件管理 (核心约束)
3.1 工作区 (Workspace) 机制
系统不硬编码任何数据路径。采用“工作区”模式：

首次启动: 检查配置文件 app_settings.json。若未配置工作区，弹出 FilePicker 让用户选择一个根目录（例如 D:\SCM_Workspace）。

初始化: 系统自动在该根目录下创建标准子文件夹结构（见 3.2）。

规则迁移: 如果是首次创建，系统需自动将内置的默认规则模板写入 Config 文件夹。

3.2 目录结构规范
Plaintext

[User_Selected_Workspace_Root]/
├── 00_Config/               # 存放所有模块的规则表 (CSV/XLSX)
├── 01_Finance_Audit/        # 模块1：资金对账
│   ├── Inputs_Accrual/      # 放入：计提台账
│   ├── Inputs_Actual/       # 放入：实际账单
│   └── Outputs/             # 生成：核对结果
├── 02_Bill_Parser/          # 模块2：ERP账单清洗 (原 jiti_tool)
│   ├── Source_Files/        # 放入：供应商原始文件
│   └── Parsed_Results/      # 生成：清洗后的中间数据
├── 03_Upload_Template/      # 模块3：上传模版转换
│   └── Generated_Templates/ # 生成：可直接上传系统的Excel
└── 04_Quote_Audit/          # 模块4：报价稽核
    └── Audit_Reports/       # 生成：报价差异报告
4. 功能模块详情
模块一：资金对账引擎 (原 auto_audit.py 重构)
功能描述: 核对“计提台账”与“实际账单”的金额与明细差异。

输入:

文件 A: 计提台账 (Excel) - 用户通过 GUI 选择或放入指定文件夹。

文件 B: 实际账单 (Excel) - 用户通过 GUI 选择。

规则: 读取 00_Config/audit_rules.csv (定义列名映射、匹配逻辑)。

逻辑:

加载规则配置。

执行 Pandas merge/compare 操作。

异常处理: 若列名不匹配，UI 需弹窗提示具体缺少的列名。

输出: 在 01_Finance_Audit/Outputs 生成带有差异高亮的 Excel 报告。

模块二：ERP 账单清洗器 (原 jiti_tool.py 重构)
功能描述: 将供应商导出的复杂 ERP 格式账单，转换为人类可读或系统可处理的标准中间格式。

输入:

源文件: 供应商原始账单 (Excel/CSV)。

规则: 读取 00_Config/parser_rules.json (定义不同供应商的清洗逻辑)。

逻辑:

识别供应商类型（根据文件名或表头特征）。

提取有效字段（单号、金额、重量、费用项）。

输出: 清洗后的明细表存入 02_Bill_Parser/Parsed_Results。

模块三：实际账单上传模版生成 (新开发)
功能描述: 将清洗后的账单数据，映射并填入公司 OMS 系统的标准上传模版中。

输入:

数据源: 模块二生成的清洗后数据，或用户手动上传的明细表。

映射规则: 00_Config/upload_mapping.csv (定义源列 -> 目标列的映射关系)。

逻辑:

读取映射规则。

创建符合 OMS 要求的 Excel 结构（特定表头、特定Sheet名）。

填充数据。

输出: 系统可识别的 Excel 文件存入 03_Upload_Template/Generated_Templates。

模块四：报价符合性稽核 (新开发)
功能描述: 校验供应商账单中的单价是否符合合同报价。

输入:

账单文件: 实际账单明细。

报价库: 必须是标准化的 Excel 数据库格式，存放于 00_Config/price_database.xlsx (字段: 线路, 费用项, 单价, 生效日期)。

逻辑:

根据“线路+费用项+时间”三个维度，在报价库中查找标准单价。

比对：账单单价 vs 标准单价。

计算：差异金额 = (账单单价 - 标准单价) * 数量。

输出: 包含“是否违规”、“应付金额”、“实付金额”、“差额”字段的稽核报告。

5. UI/UX 交互要求
侧边栏导航: 左侧为功能菜单（对账、清洗、生成、稽核、设置）。

日志控制台: 界面下方包含一个只读的文本区域，实时显示 Python 的处理日志（替代 print），需支持自动滚动。

防卡死设计: 所有耗时的 Pandas 数据处理任务，必须在子线程 (Files) 中运行，避免 Flet 主界面无响应。

配置管理: 在“设置”页面提供按钮：“打开工作区目录”、“重置规则文件”。