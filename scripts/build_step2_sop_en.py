#!/usr/bin/env python3
"""Fill English into the Step 2 cleaning SOP. Chinese stays. No comments added."""
import copy
import re
import zipfile
from io import BytesIO
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "Updated Cleaning and disinfection SOP for step 2-draft.docx"

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
HAN = re.compile(r"[\u4e00-\u9fff]")
EN_FONT = "Times New Roman"
TABLE_SZ = "16"
ALREADY = re.compile(r"[A-Za-z]{4}[^/\n]{0,160}/")

# Whole-paragraph replacements for English that does not match the Chinese.
FIXES = [
    (
        "replace the Tri-clamp gaskets at both ends of the valve/将拆解的阀门依次完成复位组装，更换阀门两端连接的法兰垫片",
        "replace the flange gaskets at both ends of the valve/将拆解的阀门依次完成复位组装，更换阀门两端连接的法兰垫片",
    ),
    (
        "In-situ or disassemble cleaning/原位清洁 or 拆卸清洁",
        "In-situ cleaning or disassembly cleaning/原位清洁或拆卸清洁",
    ),
    ("PTFE lined衬PTFE", "PTFE-lined/衬PTFE"),
    ("TPTFE lined衬PTFE", "PTFE-lined/衬PTFE"),
    ("Glass 搪玻璃", "Glass-lined/搪玻璃"),
    ("Glass 玻璃", "Glass/玻璃"),
    ("Glass玻璃", "Glass/玻璃"),
    ("Glass 视镜", "Sight glass/视镜"),
    ("Glass视镜", "Sight glass/视镜"),
    ("Hastelloy/ / 哈氏合金", "Hastelloy/哈氏合金"),
    ("Hastelloy/ /哈氏合金", "Hastelloy/哈氏合金"),
    ("Valve number/ / 阀门编号", "Valve number/阀门编号"),
    ("Type/ / 阀门类型", "Valve type/阀门类型"),
    ("Connection location/ / 连接位置", "Connection location/连接位置"),
    ("Equipment material/ / 设备材质", "Equipment material/设备材质"),
    ("大清洁清洁方法", "大清洁方法"),
    ("Q验收", "QA验收"),
    ("复位8装", "复位组装"),
    ("禁锢", "紧固"),
    ("316316L不锈钢", "316L不锈钢"),
    ("哈式合金", "哈氏合金"),
    ("Tri-clamp ball valve快开球阀", "Tri-clamp ball valve/快开球阀"),
    ("Tri-clamp ball valve法兰球阀", "Flanged ball valve/法兰球阀"),
    ("Tri-Clamp ball valve 快开球阀", "Tri-clamp ball valve/快开球阀"),
    ("Tri-Clamp ball valve 快装球阀", "Tri-clamp ball valve/快装球阀"),
    ("Flanged Ball valve法兰球阀", "Flanged ball valve/法兰球阀"),
    ("Flanged Ball valve/法兰球阀", "Flanged ball valve/法兰球阀"),
    ("Flanged Ball Valve 法兰球阀", "Flanged ball valve/法兰球阀"),
    ("Flanged Ball valve 法兰球阀", "Flanged ball valve/法兰球阀"),
    ("each Step 1 equipment are shown", "each item of Step 1 equipment is shown"),
    ("each individual Step 1 equipment", "each individual item of Step 1 equipment"),
    ("RATJ42024002", "RA-TJ42024002"),
    (
        "〈Risk Assessment for Equipment Sampling Point Locations 〉(RA-TJ42024002)",
        "Risk Assessment for Equipment Sampling Point Locations (RA-TJ42024002)",
    ),
]

SHORT = {
    "质量控制部，生产部，质量保证部": "Quality Control, Production, and Quality Assurance",
    "目录": "Contents",
    "厂区": "Site",
    "部门": "Department",
    "程序": "Procedure",
    "序号": "No.",
    "切割": "Cleavage",
    "沉淀": "Precipitation",
    "材质": "Material of construction",
    "衬哈": "Halar-lined",
    "/衬哈": "Halar-lined",
    "玻璃": "Glass",
    "弯头": "Elbow",
    "管道": "Pipe",
    "名称": "Name",
    "三通": "Tee",
    "变径": "Reducer",
    "软管": "Hose",
    "直管": "Straight pipe",
    "手套": "Glove",
    "乳胶": "Latex",
    "编号": "No.",
    "配液釜": "Solution-preparation reactor",
    "搪玻璃": "Glass-lined",
    "切割釜": "Cleavage reactor",
    "压滤罐": "Pressure filter",
    "沉淀釜": "Precipitation reactor",
    "搅拌罐": "Mixing tank",
    "异径管": "Reducer",
    "隔膜阀": "Diaphragm valve",
    "衬四氟": "PTFE-lined",
    "聚丙烯": "Polypropylene",
    "版本号": "Version",
    "适用对象": "Applicable to",
    "术语定义": "Definitions",
    "设备名称": "Equipment name",
    "设备编号": "Equipment No.",
    "设备用途": "Equipment purpose",
    "设备材质": "Material of construction",
    "树脂过滤": "Resin filtration",
    "哈氏合金": "Hastelloy",
    "粗品过滤": "Crude-product filtration",
    "清洁策略": "Cleaning strategy",
    "氮气吹扫": "Nitrogen purge",
    "阀门编号": "Valve No.",
    "阀门类型": "Valve type",
    "连接位置": "Connection location",
    "清洁方法": "Cleaning method",
    "法兰球阀": "Flanged ball valve",
    "回流管道": "Reflux line",
    "氮气管道": "Nitrogen line",
    "原位清洁": "In-situ cleaning",
    "溶剂管道": "Solvent line",
    "拆卸清洁": "Disassembly cleaning",
    "溶剂管路": "Solvent line",
    "管线编号": "Line No.",
    "管路编号": "Line No.",
    "尾气管道": "Vent line",
    "蒸馏管道": "Distillation line",
    "拆卸清洗": "Disassembly cleaning",
    "氮气管路": "Nitrogen line",
    "快开球阀": "Tri-clamp ball valve",
    "尾气管路": "Vent line",
    "耗材类别": "Consumable",
    "使用位置": "Where used",
    "取样方式": "Sampling method",
    "修订历史": "Revision history",
    "生效日期": "Effective date",
    "目标/范围": "Objective/Scope",
    "配置切割液": "Preparation of cleavage solution",
    "过滤干燥机": "Filter dryer",
    "批次间清洁": "Between-batch cleaning",
    "溶剂桶管道": "Solvent-drum line",
    "衬PTFE": "PTFE-lined",
    "安装后检查": "Post-installation check",
    "淋洗清洁：": "Rinse cleaning:",
    "甲醇淋洗：": "Methanol rinse:",
    "阀门·编号": "Valve No.",
    "流体过滤器": "Liquid filter",
    "三芯过滤器": "Three-element filter",
    "袋式过滤器": "Bag filter",
    "气动隔膜泵": "Pneumatic diaphragm pump",
    "批次间更换": "Replace between batches",
    "回流取样；": "Reflux sampling;",
    "淋洗取样；": "Rinse sampling;",
    "加料插底管": "Charging dip pipe",
    "氮气主管路": "Main nitrogen line",
    "设备概况描述": "Equipment overview",
    "MTBE预冷": "MTBE pre-cooling",
    "每批原位清洁": "In-situ cleaning every batch",
    "馏分转移管道": "Distillate transfer line",
    "反应釜釜下口": "Reactor bottom outlet",
    "新鲜溶剂管道": "Fresh-solvent line",
    "304不锈钢": "304 stainless steel",
    "压滤罐进料口": "Pressure-filter inlet",
    "压滤罐下管路": "Pressure-filter bottom line",
    "压滤罐温度计": "Pressure-filter thermometer",
    "压滤罐淋洗口": "Pressure-filter rinse port",
    "原位清洁方法": "In-situ cleaning method",
    "纯化水淋洗：": "Purified-water rinse:",
    "储罐出口管路": "Tank outlet line",
    "新鲜溶剂管线": "Fresh-solvent line",
    "设备消毒流程": "Equipment disinfection procedure",
    "接收罐液位计": "Receiver level gauge",
    "桶装溶剂管道": "Drummed-solvent line",
    "其他溶剂管道": "Other solvent line",
    "预冷MTBE": "Pre-cooled MTBE",
    "设备链的消毒": "Disinfection of the equipment train",
    "表1，设备概况": "Table 1. Equipment overview",
    "设备清洁和消毒": "Equipment cleaning and disinfection",
    "清洁策略的出处": "Source of the cleaning strategy",
    "每4批拆卸清洁": "Disassembly cleaning every 4 batches",
    "大清洁方法": "Major cleaning method",
    "压力表下口阀门": "Valve below the pressure gauge",
    "非专用溶剂管道": "Non-dedicated solvent line",
    "接收罐氮气管道": "Receiver nitrogen line",
    "接收罐尾气管道": "Receiver vent line",
    "接收罐废液管道": "Receiver waste line",
    "反应釜下口阀门": "Reactor bottom-outlet valve",
    "反应釜观察视镜": "Reactor sight glass",
    "316L不锈钢": "316L stainless steel",
    "反应釜清洗管路": "Reactor cleaning line",
    "完整性检查内容": "Integrity-check item",
    "所有配件齐全。": "All accessories are present.",
    "压滤罐预留管道": "Pressure-filter spare nozzle",
    "压滤罐至压力表": "Pressure filter to pressure gauge",
    "压滤罐观察视镜": "Pressure-filter sight glass",
    "罐体打压试密：": "Vessel pressure leak test:",
    "反应釜尾气管路": "Reactor vent line",
    "反应釜预留管路": "Reactor spare nozzle",
    "手套箱预留管路": "Glove-box spare nozzle",
    "手套箱尾气管路": "Glove-box vent line",
    "清洁中是否更换": "Replaced during cleaning",
    "袋式过滤器滤袋": "Bag-filter bag",
    "氮气过滤器滤芯": "Nitrogen-filter cartridge",
    "尾气过滤器滤芯": "Vent-filter cartridge",
    "流体过滤器滤芯": "Liquid-filter cartridge",
    "取样方式和位置": "Sampling method and location",
    "生产结束后清洁": "End-of-production cleaning",
    "反应釜下口管路": "Reactor bottom-outlet line",
    "304L不锈钢": "304L stainless steel",
    "接收罐液位计阀门": "Receiver level-gauge valve",
    "釜下氮气吹扫管道": "Nitrogen-purge line below the reactor",
    "反应釜蒸馏管路口": "Reactor distillation nozzle",
    "接收罐预留口阀门": "Receiver spare-nozzle valve",
    "反应釜釜下口阀门": "Reactor bottom-outlet valve",
    "新鲜溶剂进釜管路": "Fresh-solvent line into the reactor",
    "316L 不锈钢": "316L stainless steel",
    "反应釜固体投料口阀门": "Reactor solids-charging valve",
    "过滤干燥机清洗：": "Filter-dryer cleaning:",
    "3、手套箱清洗：": "3. Glove-box cleaning:",
    "过滤干燥机预留口": "Filter-dryer spare nozzle",
    "接受罐液位计管路": "Receiver level-gauge line",
    "过滤干燥机捕集器": "Filter-dryer trap",
    "储罐到废液间管路": "Line from the tank to the waste room",
    "目标物，分析方法": "Target, analytical method",
    "接收罐溢出口阀门": "Receiver overflow valve",
    "取样方式和位置*": "Sampling method and location*",
    "生产结束后拆卸清洁": "Disassembly cleaning at the end of production",
    "项目结束后拆卸清洁": "Disassembly cleaning at the end of the project",
    "反应釜淋洗预留管路": "Reactor rinse spare line",
    "接受罐预留管道阀门": "Receiver spare-line valve",
    "反应釜上口尾气管路": "Reactor head vent line",
    "取样箱的手套已安装": "The sampling-box gloves are installed",
    "淋洗取样和擦拭取样": "Rinse sampling and swab sampling",
    "设备及设备链的消毒": "Disinfection of the equipment and the equipment train",
    "反应釜液体加料管路": "Reactor liquid-charging line",
    "表4，切割釜主体清洗": "Table 4. Cleaning of the cleavage-reactor body",
    "接收罐压力表下口阀门": "Valve below the receiver pressure gauge",
    "表8，配液釜主体清洗": "Table 8. Cleaning of the solution-preparation reactor body",
    "甲醇淋洗（原位清洁）": "Methanol rinse (in-situ cleaning)",
    "氮气吹干（原位清洁）": "Nitrogen blow-down (in-situ cleaning)",
    "反应釜压力表下口阀门": "Valve below the reactor pressure gauge",
    "接收罐废液管道预留口": "Spare nozzle on the receiver waste line",
    "预冷MTBE转移管路": "Pre-cooled MTBE transfer line",
    "2、过滤干燥机清洗：": "2. Filter-dryer cleaning:",
    "非专用溶剂或桶装溶剂": "Non-dedicated solvent or drummed solvent",
    "擦拭取样和淋洗取样；": "Swab sampling and rinse sampling;",
    "使用75%乙醇淋洗：": "Rinse with 75% ethanol:",
    "反应釜下氮气吹扫管路": "Nitrogen-purge line below the reactor",
    "将气动隔膜泵回装复位": "Reinstall the pneumatic diaphragm pump",
    "AT06-反应釜管路": "AT06-to-reactor line",
    "表12，压滤罐主体清洗": "Table 12. Cleaning of the pressure-filter body",
    "表16，沉淀釜主体清洗": "Table 16. Cleaning of the precipitation-reactor body",
    "反应釜上口液体加料管路": "Reactor-head liquid-charging line",
    "过滤干燥机到压力变送器": "Filter dryer to pressure transmitter",
    "表30，切割釜主体清洗": "Table 30. Cleaning of the cleavage-reactor body",
    "表35，配液釜主体清洗": "Table 35. Cleaning of the solution-preparation reactor body",
    "其他溶剂加料装置及管路": "Other-solvent charging assembly and line",
    "表40，压滤罐主体清洗": "Table 40. Cleaning of the pressure-filter body",
    "表44，沉淀釜主体清洗": "Table 44. Cleaning of the precipitation-reactor body",
    "表54，搅拌罐主体清洗": "Table 54. Cleaning of the mixing-tank body",
    "设备链消毒后的处理方式": "Handling after disinfection of the equipment train",
    "甲醇淋洗（原位清洁）：": "Methanol rinse (in-situ cleaning):",
    "纯化水冲洗（原位清洁）": "Purified-water flush (in-situ cleaning)",
    "溶剂管道，接移动隔膜泵": "Solvent line, connected to a mobile diaphragm pump",
    "接收罐液位计预留口阀门": "Spare valve on the receiver level gauge",
    "纯化水淋洗（原位清洁）：": "Purified-water rinse (in-situ cleaning):",
    "接收罐尾气管道预留口阀门": "Spare valve on the receiver vent line",
    "接收罐废液管道预留口阀门": "Spare valve on the receiver waste line",
    "接收罐液位计预留管道阀门": "Spare valve on the receiver level-gauge line",
    "表7，切割釜设备安装确认": "Table 7. Installation check of the cleavage reactor",
    "参照工艺辅助材料使用计划": "Refer to the process auxiliary-material use plan",
    "表30，设备及设备链消毒": "Table 30. Disinfection of the equipment and equipment train",
    "3.1.1 切割釜主体清洗": "3.1.1 Cleaning of the cleavage-reactor body",
    "溶剂选择，清洁温度等的出处": "Source of solvent selection, cleaning temperature, and related conditions",
    "表14，压滤罐自身配件清洗": "Table 14. Cleaning of pressure-filter accessories",
    "表15，压滤罐设备安装确认": "Table 15. Installation check of the pressure filter",
    "外表面清洁，无灰尘和污渍。": "External surfaces are clean, with no dust or stains.",
    "表18，沉淀釜自身管线清洗": "Table 18. Cleaning of precipitation-reactor lines",
    "表19，沉淀釜设备安装确认": "Table 19. Installation check of the precipitation reactor",
    "表20，过滤干燥机主体清洗": "Table 20. Cleaning of the filter-dryer body",
    "表29，设备清洁后取样信息": "Table 29. Sampling after equipment cleaning",
    "内表面、物料进口、物料出口": "Inner surface, material inlet, material outlet",
    "表34，切割釜设备安装确认": "Table 34. Installation check of the cleavage reactor",
    "5.2.1 配液釜主体清洗": "5.2.1 Cleaning of the solution-preparation reactor body",
    "表39，配液釜设备安装确认": "Table 39. Installation check of the solution-preparation reactor",
    "5.2.5 压滤罐主体清洗": "5.2.5 Cleaning of the pressure-filter body",
    "表42，压滤罐自身配件清洗": "Table 42. Cleaning of pressure-filter accessories",
    "表43，压滤罐设备安装确认": "Table 43. Installation check of the pressure filter",
    "5.2.9 沉淀釜主体清洗": "5.2.9 Cleaning of the precipitation-reactor body",
    "表47，沉淀釜自身管线清洗": "Table 47. Cleaning of precipitation-reactor lines",
    "表48，沉淀釜设备安装确认": "Table 48. Installation check of the precipitation reactor",
    "表49，过滤干燥机主体清洗": "Table 49. Cleaning of the filter-dryer body",
    "表56，搅拌罐主体配件清洁": "Table 56. Cleaning of mixing-tank accessories",
    "表57，搅拌罐自身管路清洁": "Table 57. Cleaning of mixing-tank lines",
    "表58，搅拌罐设备安装确认": "Table 58. Installation check of the mixing tank",
    "表2，单体设备P&ID图信息": "Table 2. P&ID information for individual equipment",
    "该管路未接触物料，不涉及清洁": "This line has no product contact; cleaning is not applicable",
    "表9，配液釜阀门及部件的清洗": "Table 9. Cleaning of valves and components on the solution-preparation reactor",
    "表13，压滤罐设备完整性检查": "Table 13. Integrity check of the pressure filter",
    "表17，沉淀釜阀门及配件清洗": "Table 17. Cleaning of valves and fittings on the precipitation reactor",
    "5.1.1 切割釜主体的清洁": "5.1.1 Cleaning of the cleavage-reactor body",
    "表38，配液釜自身管件的清洗": "Table 38. Cleaning of fittings on the solution-preparation reactor",
    "表46，沉淀釜阀门及配件清洗": "Table 46. Cleaning of valves and fittings on the precipitation reactor",
    "5.2.16 搅拌罐主体清洗": "5.2.16 Cleaning of the mixing-tank body",
    "将拆解的αβ蝶阀依次完成复位组装": "Reassemble the split butterfly valve in the reverse order",
    "表10，配液釜自身配备的管路清洗": "Table 10. Cleaning of lines on the solution-preparation reactor",
    "该阀门未直接接触物料，不涉及清洗": "This valve has no direct product contact; cleaning is not applicable",
    "该管路不直接接触物料，不涉及清洗": "This line has no direct product contact; cleaning is not applicable",
    "将原来滤袋拆除废弃，更换新的滤袋": "Remove and discard the used filter bag, and install a new one",
    "仅涉及到溶剂，不涉及化学残留取样": "Solvent only; chemical-residue sampling is not applicable",
    "表33，切割釜自身配备的管路清洗": "Table 33. Cleaning of lines on the cleavage reactor",
    "5.2.3  配液釜自身配件清洁": "5.2.3 Cleaning of accessories on the solution-preparation reactor",
    "5.2.11 沉淀釜自身配件清洁": "5.2.11 Cleaning of accessories on the precipitation reactor",
    "5.2.11 沉淀釜设备安装确认": "5.2.11 Installation check of the precipitation reactor",
    "5.2.12 过滤干燥机主体清洗": "5.2.12 Cleaning of the filter-dryer body",
    "5.2.18 搅拌罐自身配件清洁": "5.2.18 Cleaning of accessories on the mixing tank",
    "表53，设备相关耗材清洁方式如下": "Table 53. Cleaning of equipment consumables",
    "表5，切割釜上安装的阀门及部件清洗": "Table 5. Cleaning of valves and components on the cleavage reactor",
    "回流进釜阀门及馏分转移进接收罐阀门": "Reflux-to-reactor valve and distillate-to-receiver valve",
    "手套箱使用压缩空气将手套箱整体吹干": "Blow the glove box dry with compressed air",
    "尾气管路未直接接触物料，不涉及清洗": "The vent line has no direct product contact; cleaning is not applicable",
    "该生产项目未使用接收罐，不涉及清洗": "The receiver is not used for this project; cleaning is not applicable",
    "使用不锈钢扳手将气动阀门执行器拆除": "Use a stainless steel wrench to remove the pneumatic-valve actuator",
    "5.2.19  搅拌罐设备安装确认": "5.2.19 Installation check of the mixing tank",
    "7.3.设备链中的每一个部件的清洁": "7.3. Cleaning of each component in the equipment train",
    "氮气管路管阀门未接触物料，不涉及清洗": "Valves on the nitrogen line have no product contact; cleaning is not applicable",
    "将原来滤芯拆除废弃，更换新的折叠滤芯": "Remove and discard the used cartridge, and install a new pleated cartridge",
    "表32，切割釜上安装的阀门及部件清洗": "Table 32. Cleaning of valves and components on the cleavage reactor",
    "手套箱使用压缩空气将手套箱整体吹干；": "Blow the glove box dry with compressed air;",
    "5.2.14 过滤干燥机自身配件清洁": "5.2.14 Cleaning of filter-dryer accessories",
    "5.2.15 过滤干燥机设备安装确认": "5.2.15 Installation check of the filter dryer",
    "表54，设备清洁后取样位置及方式如下": "Table 54. Sampling locations and methods after equipment cleaning",
    "3、 将拆解的αβ蝶阀依次完成复位组装": "3. Reassemble the split butterfly valve in the reverse order",
    "3.1.2 切割釜上安装的阀门及部件清洗": "3.1.2 Cleaning of valves and components on the cleavage reactor",
    "拆卸过程中检查并确认各密封垫/视镜完好。": "During disassembly, check that each gasket and sight glass is intact.",
    "此阀门在反应釜上口，未接触物料不涉及清洗": "This valve is on the reactor head and has no product contact; cleaning is not applicable",
    "粉尘捕集器已清洁，无物料痕迹、锈迹、水渍": "The dust collector is clean, with no material residue, rust, or water stains",
    "反应釜TJ4S-1213-R15氮气管路": "Nitrogen line of reactor TJ4S-1213-R15",
    "反应釜TJ4S-1213-R15尾气管路": "Vent line of reactor TJ4S-1213-R15",
    "（3）、将拆解的αβ蝶阀依次完成复位组装": "(3) Reassemble the split butterfly valve in the reverse order",
    "料残留，使用压缩气体交替吹干，QA验收；": "no material residue remains, blow dry alternately with compressed air, QA acceptance;",
    "使用压缩空气吹干，吹扫时间35±5min。": "Blow dry with compressed air. Purge for 35±5 min.",
    "将弯头进行复位安装，更换部件两端连接的垫片": "Reinstall the elbow and replace the gaskets at both ends",
    "压力表完好，在校验有效期内，正确连接可用。": "The pressure gauge is intact, within its calibration due date, correctly connected, and fit for use.",
    "（淋洗过程需保证泵后压力≥0.02MPa）": "(During rinsing, keep the pressure downstream of the pump at ≥0.02 MPa)",
    "该阀门在氮气管路中，未接触物料，不涉及清洗": "This valve is in the nitrogen line and has no product contact; cleaning is not applicable",
    "TJ4S-1213-R15到过滤干燥机管道": "Line from TJ4S-1213-R15 to the filter dryer",
    "各仪表完好，在校准有效期内，正确连接且可用": "Each instrument is intact, within its calibration due date, correctly connected, and fit for use",
    "表31，切割釜设备清洁拆卸过程中完整性检查": "Table 31. Integrity check during disassembly cleaning of the cleavage reactor",
    "将三通进行复位安装，更换部件两端连接的垫片": "Reinstall the tee and replace the gaskets at both ends",
    "表36，配液釜设备清洁拆卸过程中完整性检查": "Table 36. Integrity check during disassembly cleaning of the solution-preparation reactor",
    "将三通进行复位安装，更换部件三端连接的垫片": "Reinstall the tee and replace the gaskets on all three ends",
    "表41，压滤罐设备清洁拆卸过程中完整性检查": "Table 41. Integrity check during disassembly cleaning of the pressure filter",
    "表55，搅拌罐设备清洁拆卸过程中完整性检查": "Table 55. Integrity check during disassembly cleaning of the mixing tank",
    "TJ4S-1213-AT06搅拌罐主体管路": "Lines on mixing tank TJ4S-1213-AT06",
    "将拆解的人孔盖复位组装，将卡箍逐个紧固恢复": "Reinstall the manway cover and fasten each clamp",
    "将管件进行复位安装，更换部件两端连接的垫片": "Reinstall the fitting and replace the gaskets at both ends",
    "Flanged Ball valve法兰球阀": "Flanged ball valve",
    "此反应釜未涉及固体投料，未接触物料不涉及清洗": "This reactor is not used for solids charging and has no product contact; cleaning is not applicable",
    "氮气管路至TJ4S-1213-R15-V17": "Nitrogen line up to TJ4S-1213-R15-V17",
    "该管路为反应釜尾气管路，未接触物料不涉及清洗": "This is the reactor vent line and has no product contact; cleaning is not applicable",
    "过滤干燥机TJ4S-1212-FD06到储罐": "Filter dryer TJ4S-1212-FD06 to the tank",
    "该管路为氮气管路，未直接接触物料，不涉及清洗": "This is a nitrogen line and has no direct product contact; cleaning is not applicable",
    "表45，沉淀釜设备清洁拆卸过程中的完整性检查": "Table 45. Integrity check during disassembly cleaning of the precipitation reactor",
    "将过滤器进行复位安装，更换部件两端连接的垫片": "Reinstall the filter and replace the gaskets at both ends",
    "依次完成复位组装，更换管件两端连接的法兰垫片": "Reassemble in order and replace the flange gaskets at both ends of the fitting",
    "将Y型管进行复位安装，更换部件两端连接的垫片": "Reinstall the Y-pipe and replace the gaskets at both ends",
    "此生产项目未使用接收罐、未接触物料，不涉及清洗": "The receiver is not used for this project and has no product contact; cleaning is not applicable",
    "此阀门在反应釜尾气管路中，未接触物料不涉及清洗": "This valve is in the reactor vent line and has no product contact; cleaning is not applicable",
    "过滤干燥机TJ4S-1212-FD06到接收罐": "Filter dryer TJ4S-1212-FD06 to the receiver",
    "3.6. Sam设备清洁后的取样位置，取样方式": "3.6. Sampling locations and methods after equipment cleaning",
    "将管件进行复位安装，更换部件两端连接的快开垫片": "Reinstall the fitting and replace the tri-clamp gaskets at both ends",
    "5.2.6 压滤罐设备清洁拆卸过程中完整性检查": "5.2.6 Integrity check during disassembly cleaning of the pressure filter",
    "更换阀门两端连接的四氟法兰垫片，将隔膜阀门复位": "Replace the PTFE flange gaskets at both ends and reinstall the diaphragm valve",
    "7.6. Sam设备清洁后的取样位置，取样方式": "7.6. Sampling locations and methods after equipment cleaning",
    "该生产项目不涉及蒸馏，管路未接触物料，不涉及清洁": "Distillation is not used for this project; the line has no product contact and cleaning is not applicable",
    "新鲜溶剂管路至TJ4S-1213-R09-V71": "Fresh-solvent line up to TJ4S-1213-R09-V71",
    "此阀门在反应釜进氮气管路中，未接触物料不涉及清洗": "This valve is in the reactor nitrogen inlet line and has no product contact; cleaning is not applicable",
    "此阀门在接收罐进氮气管路中，未接触物料不涉及清洗": "This valve is in the receiver nitrogen inlet line and has no product contact; cleaning is not applicable",
    "接收罐废液管至TJ4S-1213-R10-V25": "Receiver waste line up to TJ4S-1213-R10-V25",
    "TJ4S-1213-F04母液至反应釜进料口阀门": "Valve from TJ4S-1213-F04 mother liquor to the reactor inlet",
    "冷凝器TJ4S-1212-FD06-C01到储罐": "Condenser TJ4S-1212-FD06-C01 to the tank",
    "冷凝器接收罐的管路阀门未直接接触物料，不涉及清洗": "Valves on the condenser-receiver line have no direct product contact; cleaning is not applicable",
    "3.3. 设备链中的每一个连接管道以及阀门的清洁": "3.3. Cleaning of each connecting line and valve in the equipment train",
    "过滤干燥机TJ4S-1212-FD06自带手套箱": "Glove box on filter dryer TJ4S-1212-FD06",
    "5.1.2 切割釜设备清洁拆卸过程中的完整性检查": "5.1.2 Integrity check during disassembly cleaning of the cleavage reactor",
    "使用不锈钢扳手将三通各端连接法兰断开，将管道拆除": "Use a stainless steel wrench to break the flange at each end of the tee and remove the pipe",
    "使用不锈钢扳手将弯头两端连接法兰断开，将弯头拆除": "Use a stainless steel wrench to break the flanges at both ends of the elbow and remove the elbow",
    "使用不锈钢扳手将管道两端连接法兰断开，将管道拆除": "Use a stainless steel wrench to break the flanges at both ends of the pipe and remove the pipe",
    "5.2.2 配液釜设备清洁拆卸过程中的完整性检查": "5.2.2 Integrity check during disassembly cleaning of the solution-preparation reactor",
    "使用不锈钢扳手将变径两端连接法兰断开，将管道拆除": "Use a stainless steel wrench to break the flanges at both ends of the reducer and remove the pipe",
    "使用不锈钢扳手将三通各端连接法兰断开，将三通拆除": "Use a stainless steel wrench to break the flange at each end of the tee and remove the tee",
    "表50，过滤干燥机设备清洁拆卸过程中的完整性检查": "Table 50. Integrity check during disassembly cleaning of the filter dryer",
    "使用不锈钢扳手将四通各端连接法兰断开，将管路拆除": "Use a stainless steel wrench to break the flange at each end of the cross and remove the line",
    "生产过程不涉及回流以及蒸馏，未接触物料，不涉及清洗": "Reflux and distillation are not used in this process; there is no product contact and cleaning is not applicable",
    "此阀门在反应釜进氮气管路上，未接触物料，不涉及清洗": "This valve is on the reactor nitrogen inlet and has no product contact; cleaning is not applicable",
    "使": "SKIP",
    "表3，设备清洁策略": "Table 3. Equipment cleaning strategy",
    "蝶阀": "Butterfly valve",
    "反应釜固体投料口": "Reactor solids-charging port",
    "气动球阀": "Pneumatic ball valve",
    "止回阀": "Check valve",
    "气动调节阀": "Pneumatic control valve",
    "人孔": "Manway",
    "快装球阀": "Tri-clamp ball valve",
    "拆卸过程中检查烧结板有无破损。": "During disassembly, check the sintered plate for damage.",
    "表14内列出的所有部件和配件以及螺栓和密封垫齐全无缺失，安全正确且完好无损。":
        "All parts, accessories, bolts, and gaskets listed in Table 14 are present, secure, and undamaged.",
    "表28，step1工序相关耗材": "Table 28. Consumables for the Step 1 operation",
    "反应釜氮气管路口至TJ4S-1213-R09-V06": "Reactor nitrogen nozzle up to TJ4S-1213-R09-V06",
    "TJ4S-1213-AT06搅拌罐到TJ4S-1213-AT06-PDP01": "Mixing tank TJ4S-1213-AT06 to TJ4S-1213-AT06-PDP01",
    "强光照射目视检查，确认设备内部清洁干燥，无异物（包括金属碎屑），灰尘和污渍":
        "Inspect under a strong light and confirm that the equipment interior is clean and dry, with no foreign matter (including metal shavings), dust, or stains.",
    "使用不锈钢扳手将与三通连接端的法兰固定螺杆逆时针方向旋转松动;":
        "Use a stainless steel wrench to loosen the flange bolts at the tee connection counterclockwise.",
    "在阀门TJ4S-1213-R09-V80处连接氮气管路，开启氮气，并打开TJ4S-1213-R09-V79和进液阀门TJ4S-1213-R09-V78，TJ4S-1213-R09-V72、TJ4S-1213-R09-V71，TJ4S-1213-R09-V65，TJ4S-1213-R09-V60使用氮气吹扫管路至无液滴":
        "Connect nitrogen at TJ4S-1213-R09-V80, open the nitrogen supply and valves TJ4S-1213-R09-V79, TJ4S-1213-R09-V78, TJ4S-1213-R09-V72, TJ4S-1213-R09-V71, TJ4S-1213-R09-V65, and TJ4S-1213-R09-V60, and purge the line with nitrogen until no droplets remain.",
    "TJ4S-1213-R15内纯化水体系经过下口阀门TJ4S-1213-R15-V46、TJ4S-1213-R15-V47，通过阀门139788-step02-P04-015放入废液罐；":
        "Discharge the purified water in TJ4S-1213-R15 to the waste tank through bottom valves TJ4S-1213-R15-V46 and TJ4S-1213-R15-V47 and valve 139788-step02-P04-015.",
    "通过哈氏合金压滤罐TJ4S-1200-F04上的固定淋洗口阀门TJ4S-1200-F04-V18，使用桶装甲醇通过专用新鲜溶剂加料泵向压滤罐内加入77.5Kg±2.5kg甲醇淋洗；通过哈氏合金压滤罐TJ4S-1200-F04上的固定淋洗口阀门TJ4S-1200-F04-V18，通过专用纯化水加料泵向压滤罐内加入95.0±5.0kg纯化水淋洗至无溶剂气味残留。":
        "Through fixed rinse valve TJ4S-1200-F04-V18 on Hastelloy pressure filter TJ4S-1200-F04, charge 77.5±2.5 kg of drummed methanol with the dedicated fresh-solvent charging pump and rinse the filter. Through the same valve, charge 95.0±5.0 kg of purified water with the dedicated purified-water charging pump and rinse until no solvent odor remains.",
    "将哈氏合金压滤罐TJ4S-1200-F04上方两个手孔以及收料口拆卸解体至最小节，将压滤罐上下封头解体、筛板抽出；使用纯化水依次将拆卸的所有设备内表面，擦洗至无可视物料残留。":
        "Disassemble the two handholes and the collection port on top of Hastelloy pressure filter TJ4S-1200-F04 down to the smallest parts, separate the upper and lower heads, and pull out the screen. Wipe every removed internal surface with purified water until no material residue is visible.",
    "通过哈氏合金压滤罐TJ4S-1200-F04上的固定淋洗口阀门TJ4S-1200-F04-V18，使用桶装甲醇通过专用新鲜溶剂加料泵向压滤罐内加入77.5Kg±2.5kg甲醇淋洗；通过管路139788-step02-P03转移至TJ4S-1213-R15、经过下口阀门TJ4S-1213-R15-V46、TJ4S-1213-R15-V47，通过阀门139788-step02-P04-015放入废液桶中":
        "Through fixed rinse valve TJ4S-1200-F04-V18 on Hastelloy pressure filter TJ4S-1200-F04, charge 77.5±2.5 kg of drummed methanol with the dedicated fresh-solvent charging pump and rinse the filter. Transfer it to TJ4S-1213-R15 through line 139788-step02-P03, then discharge it into a waste drum through bottom valves TJ4S-1213-R15-V46 and TJ4S-1213-R15-V47 and valve 139788-step02-P04-015.",
    "通过哈氏合金压滤罐TJ4S-1200-F04上的固定淋洗口阀门TJ4S-1200-F04-V18，通过专用纯化水加料泵向压滤罐内加入95.0±5.0kg纯化水淋洗至无溶剂气味残留，通过管路139788-step02-P03转移至TJ4S-1213-R15、经过下口阀门TJ4S-1213-R15-V46、TJ4S-1213-R15-V47，通过阀门139788-step02-P04-015放入废液桶中":
        "Through fixed rinse valve TJ4S-1200-F04-V18 on Hastelloy pressure filter TJ4S-1200-F04, charge 95.0±5.0 kg of purified water with the dedicated purified-water charging pump and rinse until no solvent odor remains. Transfer it to TJ4S-1213-R15 through line 139788-step02-P03, then discharge it into a waste drum through bottom valves TJ4S-1213-R15-V46 and TJ4S-1213-R15-V47 and valve 139788-step02-P04-015.",
}


def P(s):
    return "Keep the pressure downstream of the pump at ≥0.02 MPa during rinsing."


LONG = {
    "注：清洁过程中物料称重记录过程可记录在物料称重记录表（QF03429）内，称重记录表编号方式为：项目号-步骤号-表格编号-两位序号，如：CPo139788-03-01-4-01表示CPo139788-03第1步，表格4中第1次称重记录":
        "Note: Material weights taken during cleaning may be entered on the material weighing record (QF03429). Number the record as project No.-step No.-table No.-two-digit sequence. For example, CPo139788-03-01-4-01 is the first weighing record in Table 4 for step 1 of CPo139788-03.",
    "该设备的清洁从清洁时间节点上分为批次间清洁和每四批清洁及生产结束后清洁。":
        "By timing, cleaning of this equipment is divided into between-batch cleaning, cleaning every four batches, and cleaning at the end of production.",
    "Step 2所有相关的设备的清洁策略整体计划如下表3：":
        "The overall cleaning strategy for all Step 2 equipment is given in Table 3 below:",
    "目标：建立ASYM-139788产品生产用切割工序及其附属设备的清洗消毒规程，以指导生产过程中的设备清洗和消毒，包括批次间清洁和项目结束后清洁。":
        "Objective: Establish the cleaning and disinfection procedure for the cleavage step and its associated equipment used to manufacture ASYM-139788, covering equipment cleaning and disinfection during production, including between-batch cleaning and cleaning at the end of the project.",
    "范围：本清洗及消毒规程仅适用于切割工序及其辅助设备的清洗和消毒，包括配液釜，切割釜，压滤罐，沉淀釜，过滤干燥机，搅拌罐等，其他工序请参见其他工序的设备清洗消毒规程。":
        "Scope: This procedure applies only to cleaning and disinfection of the cleavage step and its auxiliary equipment, including the solution-preparation reactor, cleavage reactor, pressure filter, precipitation reactor, filter dryer, and mixing tank. For other steps, see the cleaning and disinfection procedure for that step.",
}


def kg(raw):
    return raw.replace("Kg", "kg").replace("kg", " kg").replace("  kg", " kg")


def valves(text):
    return text


# Process narratives keyed exactly.
def add_long():
    L = {}
    L["通过反应釜TJ4S-1213-R09上的固定淋洗管路上阀门TJ4S-1213-R09-V77和TJ4S-1213-R09-V10，使用桶装甲醇通过新鲜溶剂管路和新鲜溶剂专用气动隔膜泵加入175±10kg甲醇淋洗反应釜；（淋洗过程需保证泵后压力≥0.02MPa）"] = (
        "Through valves TJ4S-1213-R09-V77 and TJ4S-1213-R09-V10 on the fixed rinse line of reactor TJ4S-1213-R09, charge 175±10 kg of drummed methanol with the fresh-solvent line and the dedicated fresh-solvent pneumatic diaphragm pump, and rinse the reactor. Keep the pressure downstream of the pump at ≥0.02 MPa during rinsing."
    )
    L["TJ4S-1213-R09内淋洗甲醇通过下口阀门TJ4S-1213-R09-V46、TJ4S-1213-R09-V47以及转移管路139788-step02-P02转移至TJ4S-1200-F04；"] = (
        "Transfer the methanol rinse from TJ4S-1213-R09 to TJ4S-1200-F04 through bottom valves TJ4S-1213-R09-V46 and TJ4S-1213-R09-V47 and transfer line 139788-step02-P02."
    )
    L["TJ4S-1200-F04压滤罐内甲醇通过母液转移管路139788-step02-P03、袋式过滤器TJ4S-1212-BF005，使用气动隔膜泵TJ4S-1200-PDP035转移至TJ4S-1213-R15；"] = (
        "Transfer the methanol in pressure filter TJ4S-1200-F04 to TJ4S-1213-R15 through mother-liquor line 139788-step02-P03 and bag filter TJ4S-1212-BF005, using pneumatic diaphragm pump TJ4S-1200-PDP035."
    )
    L["TJ4S-1213-R15内甲醇体系经过下口阀门TJ4S-1213-R15-V46、TJ4S-1213-R15-V47，通过阀门139788-step02-P04-015放入废液罐；"] = (
        "Discharge the methanol in TJ4S-1213-R15 to the waste tank through bottom valves TJ4S-1213-R15-V46 and TJ4S-1213-R15-V47 and valve 139788-step02-P04-015."
    )
    L["使用纯化水专用管路和专用流量通过反应釜TJ4S-1213-R09上的固定淋洗管路上阀门TJ4S-1213-R09-V77和TJ4S-1213-R09-V10，加入175±10kg纯化水淋洗反应釜；"] = (
        "Through valves TJ4S-1213-R09-V77 and TJ4S-1213-R09-V10 on the fixed rinse line of reactor TJ4S-1213-R09, charge 175±10 kg of purified water by the dedicated purified-water line and dedicated flowmeter, and rinse the reactor."
    )
    L["TJ4S-1213-R09内淋洗纯化水通过下口阀门TJ4S-1213-R09-V46、TJ4S-1213-R09-V47以及转移管路139788-step02-P02转移至TJ4S-1200-F04；"] = (
        "Transfer the purified-water rinse from TJ4S-1213-R09 to TJ4S-1200-F04 through bottom valves TJ4S-1213-R09-V46 and TJ4S-1213-R09-V47 and transfer line 139788-step02-P02."
    )
    L["TJ4S-1200-F04压滤罐内纯化水通过母液转移管路139788-step02-P03、袋式过滤器TJ4S-1212-BF005，使用气动隔膜泵TJ4S-1200-PDP035转移至TJ4S-1213-R15"] = (
        "Transfer the purified water in pressure filter TJ4S-1200-F04 to TJ4S-1213-R15 through mother-liquor line 139788-step02-P03 and bag filter TJ4S-1212-BF005, using pneumatic diaphragm pump TJ4S-1200-PDP035."
    )
    L["TJ4S-1213-R15内纯化水体系经过下口阀门TJ4S-1213-R15-V46、TJ4S-1213-R15-V47，通过139788-step02-P04-015放入废液罐；"] = (
        "Discharge the purified water in TJ4S-1213-R15 to the waste tank through bottom valves TJ4S-1213-R15-V46 and TJ4S-1213-R15-V47 and 139788-step02-P04-015."
    )
    L["通过反应釜TJ4S-1213-R09上的固定氮气管路TJ4S-1213-R09-P02阀门TJ4S-1213-R09-V06使用氮气对反应釜及转移管路139788-step02-P02进行吹干"] = (
        "Through fixed nitrogen line TJ4S-1213-R09-P02 and valve TJ4S-1213-R09-V06 on reactor TJ4S-1213-R09, blow the reactor and transfer line 139788-step02-P02 dry with nitrogen."
    )
    L["在非专用溶剂管路处连接氮气管路，开启氮气，并打开TJ4S-1213-R09-V08，TJ4S-1213-R09-V48，TJ4S-1213-R09-V60，TJ4S-1213-R09-V61氮气吹扫至无液滴"] = (
        "Connect nitrogen at the non-dedicated solvent line, open the nitrogen supply and valves TJ4S-1213-R09-V08, TJ4S-1213-R09-V48, TJ4S-1213-R09-V60, and TJ4S-1213-R09-V61, and purge with nitrogen until no droplets remain."
    )
    L["加料结束后，在阀门TJ4S-1213-R09-V80处连接氮气管路，开启氮气，并打开TJ4S-1213-R09-V79和进液阀门TJ4S-1213-R09-V78，TJ4S-1213-R09-V72、TJ4S-1213-R09-V71，TJ4S-1213-R09-V65，TJ4S-1213-R09-V60使用氮气吹扫管路至无液滴"] = (
        "After charging, connect nitrogen at TJ4S-1213-R09-V80, open the nitrogen supply and valves TJ4S-1213-R09-V79, TJ4S-1213-R09-V78, TJ4S-1213-R09-V72, TJ4S-1213-R09-V71, TJ4S-1213-R09-V65, and TJ4S-1213-R09-V60, and purge the line with nitrogen until no droplets remain."
    )
    L["在阀门TJ4S-1213-R09-V79处连接氮气管路，开启氮气，并打开TJ4S-1213-R09-V78、TJ4S-1213-R09-V72、TJ4S-1213-R09-V71、TJ4S-1213-R09-V65和进料阀门TJ4S-1213-R09-V60使用氮气吹扫气动隔膜泵TJ4S-1200-PDP037及管路TJ4S-1213-R09-P04至无液滴"] = (
        "Connect nitrogen at TJ4S-1213-R09-V79, open the nitrogen supply and valves TJ4S-1213-R09-V78, TJ4S-1213-R09-V72, TJ4S-1213-R09-V71, TJ4S-1213-R09-V65, and inlet valve TJ4S-1213-R09-V60, and purge pneumatic diaphragm pump TJ4S-1200-PDP037 and line TJ4S-1213-R09-P04 with nitrogen until no droplets remain."
    )
    L["在阀门TJ4S-1213-R09-V80处连接氮气管路，开启氮气，打开TJ4S-1213-R09-V80，使用氮气吹扫插底管TJ4S-1200-DTu037至无液滴"] = (
        "Connect nitrogen at TJ4S-1213-R09-V80, open the nitrogen supply and TJ4S-1213-R09-V80, and purge dip pipe TJ4S-1200-DTu037 with nitrogen until no droplets remain."
    )
    L["回流进釜阀门及馏分转移进接收罐阀门TJ4S-1213-R09-V11、TJ4S-1213-R09-V21"] = (
        "Reflux-to-reactor valve and distillate-to-receiver valves TJ4S-1213-R09-V11 and TJ4S-1213-R09-V21"
    )
    L["通过反应釜清洗管线TJ4S-1213-R10-P02经过固定淋洗口TJ4S-1213-R10-V71、TJ4S-1213-R10-V65以及TJ4S-1213-R10-V61加入175±10kg甲醇淋洗反应釜（淋洗过程需保证泵后压力≥0.02MPa）"] = (
        "Through cleaning line TJ4S-1213-R10-P02 and fixed rinse ports TJ4S-1213-R10-V71, TJ4S-1213-R10-V65, and TJ4S-1213-R10-V61, charge 175±10 kg of methanol and rinse the reactor. Keep the pressure downstream of the pump at ≥0.02 MPa during rinsing."
    )
    L["TJ4S-1213-R10内淋洗甲醇通过下口阀门TJ4S-1213-R10-V46、TJ4S-1213-R10-V47、139788-step02-P01-004以及釜下转移管路139788-step02-P01，使用隔膜气动泵TJ4S-1200-PDP025通过反应釜进液管路TJ4S-1213-R09-P04及阀门TJ4S-1213-R09-V65、TJ4S-1213-R09-V60；TJ4S-1213-R09-V61转移至 TJ4S-1213-R09"] = (
        "Transfer the methanol rinse from TJ4S-1213-R10 to TJ4S-1213-R09 through bottom valves TJ4S-1213-R10-V46 and TJ4S-1213-R10-V47, 139788-step02-P01-004, and bottom transfer line 139788-step02-P01, using pneumatic diaphragm pump TJ4S-1200-PDP025, reactor inlet line TJ4S-1213-R09-P04, and valves TJ4S-1213-R09-V65, TJ4S-1213-R09-V60, and TJ4S-1213-R09-V61."
    )
    L["TJ4S-1213-R10内淋洗甲醇通过下口阀门TJ4S-1213-R10-V46、TJ4S-1213-R10-V47以及釜下转移管路139788-step02-P01，使用隔膜气动泵TJ4S-1200-PDP025通过反应釜进液管路TJ4S-1213-R09-P04及阀门TJ4S-1213-R09-V65、TJ4S-1213-R09-V60；TJ4S-1213-R09-V61转移至 TJ4S-1213-R09；"] = L[
        "TJ4S-1213-R10内淋洗甲醇通过下口阀门TJ4S-1213-R10-V46、TJ4S-1213-R10-V47、139788-step02-P01-004以及釜下转移管路139788-step02-P01，使用隔膜气动泵TJ4S-1200-PDP025通过反应釜进液管路TJ4S-1213-R09-P04及阀门TJ4S-1213-R09-V65、TJ4S-1213-R09-V60；TJ4S-1213-R09-V61转移至 TJ4S-1213-R09"
    ].replace("139788-step02-P01-004, and bottom transfer line", "bottom transfer line")
    # The sentence without -004 is different; set it explicitly below if needed.
    L["TJ4S-1213-R09内淋洗甲醇通过下口阀门TJ4S-1213-R09-V46、TJ4S-1213-R09-V47、TJ4S-1213-R09-V74以及转移管路139788-step02-P02转移至TJ4S-1200-F04；"] = (
        "Transfer the methanol rinse from TJ4S-1213-R09 to TJ4S-1200-F04 through bottom valves TJ4S-1213-R09-V46, TJ4S-1213-R09-V47, and TJ4S-1213-R09-V74 and transfer line 139788-step02-P02."
    )
    L["TJ4S-1200-F04压滤罐内甲醇通过TJ4S-1200-F04-V15、TJ4S-1200-F04-V16母液转移管路139788-step02-P03、袋式过滤器TJ4S-1212-BF005，使用隔膜气动泵TJ4S-1200-PDP035经过管路139788-step02-P03转移至TJ4S-1213-R15；"] = (
        "Transfer the methanol in pressure filter TJ4S-1200-F04 to TJ4S-1213-R15 through valves TJ4S-1200-F04-V15 and TJ4S-1200-F04-V16, mother-liquor line 139788-step02-P03, and bag filter TJ4S-1212-BF005, using pneumatic diaphragm pump TJ4S-1200-PDP035."
    )
    L["TJ4S-1213-R15内甲醇体系经过下口阀门TJ4S-1213-R15-V46、TJ4S-1213-R15-V47，通过139788-step02-P04-015放出废液桶中"] = (
        "Discharge the methanol in TJ4S-1213-R15 to a waste drum through bottom valves TJ4S-1213-R15-V46 and TJ4S-1213-R15-V47 and 139788-step02-P04-015."
    )
    L["使用纯化水专用管路和专用流量通过反应釜清洗管线TJ4S-1213-R10-P02经过固定淋洗口TJ4S-1213-R10-V71、TJ4S-1213-R10-V65以及TJ4S-1213-R10-V61加入175±10kg纯化水淋洗反应釜"] = (
        "Through cleaning line TJ4S-1213-R10-P02 and fixed rinse ports TJ4S-1213-R10-V71, TJ4S-1213-R10-V65, and TJ4S-1213-R10-V61, charge 175±10 kg of purified water by the dedicated purified-water line and dedicated flowmeter, and rinse the reactor."
    )
    L["TJ4S-1213-R09内淋洗纯化水通过下口阀门TJ4S-1213-R09-V46、TJ4S-1213-R09-V47、TJ4S-1213-R09-V74以及转移管路139788-step02-P02转移至TJ4S-1200-F04；"] = (
        "Transfer the purified-water rinse from TJ4S-1213-R09 to TJ4S-1200-F04 through bottom valves TJ4S-1213-R09-V46, TJ4S-1213-R09-V47, and TJ4S-1213-R09-V74 and transfer line 139788-step02-P02."
    )
    L["TJ4S-1200-F04压滤罐内纯化水通过TJ4S-1200-F04-V15、TJ4S-1200-F04-V16母液转移管路139788-step02-P03、袋式过滤器TJ4S-1212-BF005，使用隔膜气动泵TJ4S-1200-PDP035经过管路139788-step02-P03转移至TJ4S-1213-R15；"] = (
        "Transfer the purified water in pressure filter TJ4S-1200-F04 to TJ4S-1213-R15 through valves TJ4S-1200-F04-V15 and TJ4S-1200-F04-V16, mother-liquor line 139788-step02-P03, and bag filter TJ4S-1212-BF005, using pneumatic diaphragm pump TJ4S-1200-PDP035."
    )
    L["TJ4S-1213-R15内纯化水体系经过下口阀门TJ4S-1213-R15-V46、TJ4S-1213-R15-V47，通过139788-step02-P04-015放出废液罐；"] = (
        "Discharge the purified water in TJ4S-1213-R15 to the waste tank through bottom valves TJ4S-1213-R15-V46 and TJ4S-1213-R15-V47 and 139788-step02-P04-015."
    )
    L["通过反应釜TJ4S-1213-R10上的固定氮气使用点通过阀门TJ4S-1213-R09-V06使用氮气对反应釜及转移管路139788-step02-P01进行吹干"] = (
        "From the fixed nitrogen point on reactor TJ4S-1213-R10, blow reactor TJ4S-1213-R09 and transfer line 139788-step02-P01 dry with nitrogen through valve TJ4S-1213-R09-V06."
    )
    L["此阀门在新鲜溶剂管线中，加料结束后使用氮气对管路进行吹扫，开启阀门TJ4S-1213-R10-V08 、TJ4S-1213-R10-V48，并开启氮气，使用氮气吹扫管路及阀门至无液滴"] = (
        "This valve is in the fresh-solvent line. After charging, open valves TJ4S-1213-R10-V08 and TJ4S-1213-R10-V48 and the nitrogen supply, and purge the line and valves with nitrogen until no droplets remain."
    )
    L["此反应釜阀门同反应釜清洗管线TJ4S-1213-R10-P02 进行原位清洗，清洗方式详见表8中TJ4S-1213-R10-P02原位清洗部分"] = (
        "Clean this reactor valve in situ together with reactor cleaning line TJ4S-1213-R10-P02. See the in-situ cleaning of TJ4S-1213-R10-P02 in Table 8."
    )
    L["在阀门TJ4S-1213-R10-V75处连接氮气管路，开启氮气，并打开TJ4S-1213-R10-V77、TJ4S-1213-R10-V78、TJ4S-1213-R10-V73和进液阀门TJ4S-1213-R10-V10使用氮气吹扫气动隔膜泵TJ4S-1200-PDP005及管路TJ4S-1213-R10-P03至无液滴"] = (
        "Connect nitrogen at TJ4S-1213-R10-V75, open the nitrogen supply and valves TJ4S-1213-R10-V77, TJ4S-1213-R10-V78, TJ4S-1213-R10-V73, and inlet valve TJ4S-1213-R10-V10, and purge pneumatic diaphragm pump TJ4S-1200-PDP005 and line TJ4S-1213-R10-P03 with nitrogen until no droplets remain."
    )
    L["管线TJ4S-1213-R10-P02 进行原位清洗，清洗方式详见表7中TJ4S-1213-R10-P02原位清洗部分"] = (
        "Clean line TJ4S-1213-R10-P02 in situ. See the in-situ cleaning of TJ4S-1213-R10-P02 in Table 7."
    )
    L["使用外接氮气管路连接至阀门TJ4S-1213-R10-V75，开启氮气，再依次打开TJ4S-1213-R10-V77、TJ4S-1213-R10-V78、TJ4S-1213-R10-V73和进液阀门TJ4S-1213-R10-V10使用氮气吹扫气动隔膜泵TJ4S-1200-PDP005及管路TJ4S-1213-R10-P03，吹扫至目视无液滴"] = (
        "Connect an external nitrogen line to TJ4S-1213-R10-V75, open the nitrogen supply, then open TJ4S-1213-R10-V77, TJ4S-1213-R10-V78, TJ4S-1213-R10-V73, and inlet valve TJ4S-1213-R10-V10 in that order, and purge pneumatic diaphragm pump TJ4S-1200-PDP005 and line TJ4S-1213-R10-P03 with nitrogen until no droplets are visible."
    )
    L["通过哈式合金压滤罐TJ4S-1200-F04上的固定淋洗口阀门TJ4S-1200-F04-V18，使用桶装甲醇通过专用新鲜溶剂加料泵向压滤罐内加入77.5Kg±2.5kg甲醇淋洗；通过哈式合金压滤罐TJ4S-1200-F04上的固定淋洗口阀门TJ4S-1200-F04-V18，通过专用纯化水加料泵向压滤罐内加入95.0±5.0kg纯化水淋洗至无溶剂气味残留。"] = (
        "Through fixed rinse valve TJ4S-1200-F04-V18 on Hastelloy pressure filter TJ4S-1200-F04, charge 77.5±2.5 kg of drummed methanol with the dedicated fresh-solvent charging pump and rinse the filter. Through the same valve, charge 95.0±5.0 kg of purified water with the dedicated purified-water charging pump and rinse until no solvent odor remains."
    )
    L["将哈式合金压滤罐TJ4S-1200-F04上方两个手孔以及收料口拆卸解体至最小节，将压滤罐上下封头解体、筛板抽出；使用纯化水依次将拆卸的所有设备内表面，擦洗至无可视物料残留。"] = (
        "Disassemble the two handholes and the collection port on top of Hastelloy pressure filter TJ4S-1200-F04 down to the smallest parts, separate the upper and lower heads, and pull out the screen. Wipe every removed internal surface with purified water until no material residue is visible."
    )
    L["对照《设备内表面缺陷清单》（QF03143），通过目视检查设备内表面，检查设备是否有出现腐蚀、划痕 、涂层破损或脱落、露胎体的缺陷或现有内表面缺陷无任何变化"] = (
        "Against the Equipment Inner-Surface Defect List (QF03143), visually inspect the inner surface for corrosion, scratches, damaged or detached lining, exposed substrate, or any change in an existing inner-surface defect."
    )
    L["配件安装完毕后，使用氮气对设备进行打压试密，压滤罐打压0.2-0.3MPa，保压15-30min，压力变化≤0.01MPa。"] = (
        "After the fittings are installed, leak-test the equipment with nitrogen. Pressurize the pressure filter to 0.2–0.3 MPa, hold for 15–30 min, and confirm that the pressure change is ≤0.01 MPa."
    )
    L["强光照射目视检查，确认设备内部清洁干燥，异物，灰尘和污渍，安装后的烧结板无异物，无污染，损坏。"] = (
        "Inspect under a strong light and confirm that the equipment interior is clean and dry, with no foreign matter, dust, or stains, and that the sintered plate, once installed, is free of foreign matter, contamination, and damage."
    )
    L["通过反应釜TJ4S-1213-R15上的管线TJ4S-1213-R15-P02及进液口阀门TJ4S-1213-R15-V10、液体淋洗口阀门TJ4S-1213-R15-V66、TJ4S-1213-R15-V61、TJ4S-1213-R15-V60，使用桶装甲醇通过新鲜溶剂管路和新鲜溶剂专用气动隔膜泵加入175±10kg甲醇淋洗反应釜；"] = (
        "Through line TJ4S-1213-R15-P02, inlet valve TJ4S-1213-R15-V10, and liquid-rinse valves TJ4S-1213-R15-V66, TJ4S-1213-R15-V61, and TJ4S-1213-R15-V60 on reactor TJ4S-1213-R15, charge 175±10 kg of drummed methanol with the fresh-solvent line and the dedicated fresh-solvent pneumatic diaphragm pump, and rinse the reactor."
    )
    L["TJ4S-1213-R15内甲醇体系经过下口阀门TJ4S-1213-R15-V46、TJ4S-1213-R15-V47，通过阀门139788-step02-P04-015放入废液桶中；"] = (
        "Discharge the methanol in TJ4S-1213-R15 into a waste drum through bottom valves TJ4S-1213-R15-V46 and TJ4S-1213-R15-V47 and valve 139788-step02-P04-015."
    )
    L["使用纯化水专用管路和专用流量通过反应釜TJ4S-1213-R15上的管线TJ4S-1213-R15-P02及进液口阀门TJ4S-1213-R15-V10、液体淋洗口阀门TJ4S-1213-R15-V66、TJ4S-1213-R15-V61、TJ4S-1213-R15-V60，加入175±10kg纯化水淋洗反应釜；"] = (
        "Through line TJ4S-1213-R15-P02, inlet valve TJ4S-1213-R15-V10, and liquid-rinse valves TJ4S-1213-R15-V66, TJ4S-1213-R15-V61, and TJ4S-1213-R15-V60 on reactor TJ4S-1213-R15, charge 175±10 kg of purified water by the dedicated purified-water line and dedicated flowmeter, and rinse the reactor."
    )
    L["通过反应釜TJ4S-1213-R15上的氮气管路TJ4S-1213-R15-P01，经过氮气过滤器TJ4S-1213-R15-NF01，开启氮气阀门TJ4S-1213-R10-V06、TJ4S-1213-R10-V07使用氮气对反应釜及釜下阀门进行吹干"] = (
        "Through nitrogen line TJ4S-1213-R15-P01 and nitrogen filter TJ4S-1213-R15-NF01 on reactor TJ4S-1213-R15, open nitrogen valves TJ4S-1213-R10-V06 and TJ4S-1213-R10-V07 and blow the reactor and the valves below it dry with nitrogen."
    )
    L["此反应釜阀门同反应釜清洗管线TJ4S-1213-R15-P02 进行原位清洗，清洗方式详见表12中TJ4S-1213-R15-P02原位清洗部分"] = (
        "Clean this reactor valve in situ together with reactor cleaning line TJ4S-1213-R15-P02. See the in-situ cleaning of TJ4S-1213-R15-P02 in Table 12."
    )
    L["此阀门包含在表14中TJ4S-1213-R15-P03管线中，同该管线一并进行清洁"] = (
        "This valve is included with line TJ4S-1213-R15-P03 in Table 14 and is cleaned together with that line."
    )
    L["通过139788-STEP02-P04使用桶装甲醇通过专用新鲜溶剂气动隔膜泵经过TJ4S-1212-FD06-V93加入180kg±5kg甲醇淋洗过滤干燥机；淋洗甲醇通过母液管路TJ4S-1212-FD06-P06转移母液储罐TJ4S-1212-FD06-Tr01，再通过TJ4S-1212-FD06-P10使用气动隔膜泵TJ4S-1212-FD06-Tr01-PDP02、TJ4S-1212-FD06-P11转移至废液间;（淋洗过程需保证泵后压力≥0.02MPa）"] = (
        "Through 139788-STEP02-P04 and valve TJ4S-1212-FD06-V93, charge 180±5 kg of drummed methanol with the dedicated fresh-solvent pneumatic diaphragm pump and rinse the filter dryer. Transfer the methanol rinse to mother-liquor tank TJ4S-1212-FD06-Tr01 through mother-liquor line TJ4S-1212-FD06-P06, then to the waste room through TJ4S-1212-FD06-P10, pneumatic diaphragm pump TJ4S-1212-FD06-Tr01-PDP02, and TJ4S-1212-FD06-P11. Keep the pressure downstream of the pump at ≥0.02 MPa during rinsing."
    )
    L["通过过滤干燥机TJ4S-1212-FD06上口固定淋洗口，依次分别开启阀门TJ4S-1212-FD06-V25、TJ4S-1212-FD06-V26、TJ4S-1212-FD06-V32以及TJ4S-1212-FD06-V33、TJ4S-1212-FD06-V19，使用桶装甲醇通过专用新鲜溶剂气动隔膜泵经过管路139788-STEP02-P04加入180kg±5kg甲醇淋洗过滤干燥机主体以及捕集器；淋洗甲醇通过母液管路TJ4S-1212-FD06-P06转移母液储罐TJ4S-1212-FD06-Tr01，再通过TJ4S-1212-FD06-P10使用气动隔膜泵TJ4S-1212-FD06-Tr01-PDP02、TJ4S-1212-FD06-P11转移至废液间;（淋洗过程需保证泵后压力≥0.02MPa）"] = (
        "At the fixed rinse port on the head of filter dryer TJ4S-1212-FD06, open valves TJ4S-1212-FD06-V25, TJ4S-1212-FD06-V26, TJ4S-1212-FD06-V32, TJ4S-1212-FD06-V33, and TJ4S-1212-FD06-V19 in turn. Charge 180±5 kg of drummed methanol through line 139788-STEP02-P04 with the dedicated fresh-solvent pneumatic diaphragm pump, and rinse the filter-dryer body and the trap. Transfer the methanol rinse to mother-liquor tank TJ4S-1212-FD06-Tr01 through TJ4S-1212-FD06-P06, then to the waste room through TJ4S-1212-FD06-P10, pump TJ4S-1212-FD06-Tr01-PDP02, and TJ4S-1212-FD06-P11. Keep the pressure downstream of the pump at ≥0.02 MPa during rinsing."
    )
    L["通过过滤干燥机TJ4S-1212-FD06上口固定淋洗口，依次开启阀门TJ4S-1212-FD06-V25、TJ4S-1212-FD06-V26以及TJ4S-1212-FD06-V32，使用纯化水通过专用气动隔膜泵经过管路139788-STEP02-P04加入310kg±10kg纯化水淋洗过滤干燥机；淋洗纯化水通过母液管路TJ4S-1212-FD06-P06转移母液储罐TJ4S-1212-FD06-Tr01，再通过TJ4S-1212-FD06-P10使用气动隔膜泵TJ4S-1212-FD06-Tr01-PDP02、TJ4S-1212-FD06-P11转移至废液间"] = (
        "At the fixed rinse port on the head of filter dryer TJ4S-1212-FD06, open valves TJ4S-1212-FD06-V25, TJ4S-1212-FD06-V26, and TJ4S-1212-FD06-V32 in turn. Charge 310±10 kg of purified water through line 139788-STEP02-P04 with the dedicated pneumatic diaphragm pump and rinse the filter dryer. Transfer the rinse to mother-liquor tank TJ4S-1212-FD06-Tr01 through TJ4S-1212-FD06-P06, then to the waste room through TJ4S-1212-FD06-P10, pump TJ4S-1212-FD06-Tr01-PDP02, and TJ4S-1212-FD06-P11."
    )
    L["通过过滤干燥器TJ4S-1212-FD06上口固定氮气使用点阀门TJ4S-1212-FD06-V81、管路TJ4S-1212-FD06-P02使用氮气将过滤干燥器TJ4S-1212-FD06以及母液转移管路TJ4S-1212-FD06-P06；TJ4S-1212-FD06-P04、TJ4S-1212-FD06-P08进行吹干"] = (
        "Through fixed nitrogen point valve TJ4S-1212-FD06-V81 and line TJ4S-1212-FD06-P02 on the head of filter dryer TJ4S-1212-FD06, blow the filter dryer and mother-liquor transfer lines TJ4S-1212-FD06-P06, TJ4S-1212-FD06-P04, and TJ4S-1212-FD06-P08 dry with nitrogen."
    )
    L["将过滤干燥机出料口拆卸，使用洁净擦拭布蘸取纯化水对收料口内外表面进行擦拭清洗，直至目视无物料残留；使用压缩空气将收料口吹干，安装恢复；"] = (
        "Remove the filter-dryer discharge port. Wipe the inner and outer surfaces of the collection port with a clean wipe dipped in purified water until no material residue is visible, blow the port dry with compressed air, and reinstall it."
    )
    L["手套箱使用洁净擦拭布蘸取纯化水对手套箱内外表面进行擦拭清洗，直至目视无物料残留；"] = (
        "Wipe the inner and outer surfaces of the glove box with a clean wipe dipped in purified water until no material residue is visible."
    )
    L["再使用洁净擦拭布蘸取异丙醇对手套箱内外表面进行擦拭清洗，直至目视无物料残留；"] = (
        "Then wipe the inner and outer surfaces of the glove box with a clean wipe dipped in isopropanol until no material residue is visible."
    )
    L["过滤干燥机TJ4S-1212-FD06到冷凝器TJ4S-1212-FD06-C01"] = (
        "Filter dryer TJ4S-1212-FD06 to condenser TJ4S-1212-FD06-C01"
    )
    L["表21和表22中部分列出的所有部件和配件以及所有螺栓和密封垫齐全无缺失，安装正确且完好无损。"] = (
        "All parts, accessories, bolts, and gaskets listed in Tables 21 and 22 are present, correctly installed, and undamaged."
    )
    L["打压试密：向过滤干燥机中加入少量纯化水（加水量超过筛板）。氮气加压至 0.20~0.30MPa，保压 15~30min，压力变化≤0.01MPa。"] = (
        "Pressure leak test: add a small amount of purified water to the filter dryer, enough to cover the screen. Pressurize with nitrogen to 0.20–0.30 MPa, hold for 15–30 min, and confirm that the pressure change is ≤0.01 MPa."
    )
    L["过滤干燥机专用的物料接收罐所有部件和配件以及所有螺丝/螺栓和密封垫齐全无缺失，安装正确且完好无损。"] = (
        "On the material receiver dedicated to the filter dryer, all parts, accessories, screws/bolts, and gaskets are present, correctly installed, and undamaged."
    )
    L["过滤干燥机专用的物料接收罐所有端口密封完好，上口和下口无任何泄漏。"] = (
        "Every nozzle on the material receiver dedicated to the filter dryer is sealed, and there is no leak at the top or bottom outlet."
    )
    L["过滤干燥机专用的物料接收罐外表面清洁，无灰尘和污渍。"] = (
        "The outer surface of the material receiver dedicated to the filter dryer is clean, with no dust or stains."
    )
    L["过滤干燥机TJ4S-1212-FD06手套箱氮气管路TJ4S-1212-FD06-P01"] = (
        "Glove-box nitrogen line TJ4S-1212-FD06-P01 on filter dryer TJ4S-1212-FD06"
    )
    L["过滤干燥机TJ4S-1212-FD06氮气管路TJ4S-1212-FD06-P02"] = (
        "Nitrogen line TJ4S-1212-FD06-P02 on filter dryer TJ4S-1212-FD06"
    )
    L["过滤干燥机TJ4S-1212-FD06捕集器氮气管路TJ4S-1212-FD06-P03"] = (
        "Trap nitrogen line TJ4S-1212-FD06-P03 on filter dryer TJ4S-1212-FD06"
    )
    L["过滤干燥机TJ4S-1212-FD06尾气管路TJ4S-1212-FD06-P07"] = (
        "Vent line TJ4S-1212-FD06-P07 on filter dryer TJ4S-1212-FD06"
    )
    L["过滤干燥机TJ4S-1212-FD06手套箱尾气管路TJ4S-1212-FD06-P12"] = (
        "Glove-box vent line TJ4S-1212-FD06-P12 on filter dryer TJ4S-1212-FD06"
    )
    L["搅拌罐TJ4S-1213-AT06至反应釜TJ4S-1213-R15物料转移管路"] = (
        "Material transfer line from mixing tank TJ4S-1213-AT06 to reactor TJ4S-1213-R15"
    )
    L["压滤罐TJ4S-1200-F04至TJ4S-1213-R15物料转移管路"] = (
        "Material transfer line from pressure filter TJ4S-1200-F04 to TJ4S-1213-R15"
    )
    L["通过管路TJ4S-1213-R10-P02阀门TJ4S-1213-R10-V60、TJ4S-1213-R10-V61及喷淋球加入175±10kg的甲醇在63~75℃回流反应釜，回流时间2~4h，然后降温并通过釜下阀门TJ4S-1213-R10-V46和TJ4S-1213-R10-V47及工艺管路139788-step02-P01放出，取回流液进行Level Ⅱ检测"] = (
        "Charge 175±10 kg of methanol through line TJ4S-1213-R10-P02, valves TJ4S-1213-R10-V60 and TJ4S-1213-R10-V61, and the spray ball. Reflux the reactor at 63–75°C for 2–4 h, then cool and discharge through bottom valves TJ4S-1213-R10-V46 and TJ4S-1213-R10-V47 and process line 139788-step02-P01. Sample the reflux liquid for Level II testing."
    )
    L["通过阀门TJ4S-1213-R09-V77、TJ4S-1213-R09-V10及喷淋球加入175±10kg的甲醇在63~75℃回流反应釜，回流时间2~4h，然后降温并通过釜下阀门TJ4S-1213-R09-V46和TJ4S-1213-R09-V47及工艺管路139788-step02-P02放出，取回流液进行Level Ⅱ检测"] = (
        "Charge 175±10 kg of methanol through valves TJ4S-1213-R09-V77 and TJ4S-1213-R09-V10 and the spray ball. Reflux the reactor at 63–75°C for 2–4 h, then cool and discharge through bottom valves TJ4S-1213-R09-V46 and TJ4S-1213-R09-V47 and process line 139788-step02-P02. Sample the reflux liquid for Level II testing."
    )
    L["擦拭取样：使用甲醇溶剂按照取样点评估示意图QM04053依次擦拭内表面、物料进口、物料出口"] = (
        "Swab sampling: using methanol, swab the inner surface, material inlet, and material outlet in the order shown on sampling-point diagram QM04053."
    )
    L["淋洗取样：通过压滤罐喷淋口TJ4S-1200-F04-V18加入35±5kg甲醇淋洗压滤罐及其筛板，淋洗液取样检测Level Ⅱ残留"] = (
        "Rinse sampling: charge 35±5 kg of methanol through pressure-filter spray port TJ4S-1200-F04-V18, rinse the pressure filter and its screen, and sample the rinse for Level II residue."
    )
    L["淋洗取样：通过压滤罐喷淋口TJ4S-1200-F04-V18加入35±5kg甲醇淋洗压滤罐及其筛板，淋洗液取样检测LevelⅡ残留"] = L[
        "淋洗取样：通过压滤罐喷淋口TJ4S-1200-F04-V18加入35±5kg甲醇淋洗压滤罐及其筛板，淋洗液取样检测Level Ⅱ残留"
    ]
    L["通过反应釜TJ4S-1213-R15上的管线TJ4S-1213-R15-P02及进液口阀门TJ4S-1213-R15-V10、液体淋洗口阀门TJ4S-1213-R15-V66、TJ4S-1213-R15-V61、TJ4S-1213-R15-V60，使用桶装甲醇通过新鲜溶剂管路和新鲜溶剂专用气动隔膜泵加入175±10kg甲醇淋洗反应釜；（淋洗过程需保证泵后压力≥0.02MPa）"] = (
        "Through line TJ4S-1213-R15-P02, inlet valve TJ4S-1213-R15-V10, and liquid-rinse valves TJ4S-1213-R15-V66, TJ4S-1213-R15-V61, and TJ4S-1213-R15-V60 on reactor TJ4S-1213-R15, charge 175±10 kg of drummed methanol with the fresh-solvent line and the dedicated fresh-solvent pneumatic diaphragm pump, and rinse the reactor. Keep the pressure downstream of the pump at ≥0.02 MPa during rinsing."
    )
    L["经过下口阀门TJ4S-1213-R15-V46、TJ4S-1213-R15-V47，通过阀门139788-step02-P04-015放出，取淋洗液进行Level Ⅱ检测"] = (
        "Discharge through bottom valves TJ4S-1213-R15-V46 and TJ4S-1213-R15-V47 and valve 139788-step02-P04-015, and sample the rinse for Level II testing."
    )
    L["淋洗取样：通过过滤干燥机TJ4S-1212-FD06上口固定淋洗口，依次分别开启阀门TJ4S-1212-FD06-V25、TJ4S-1212-FD06-V26、TJ4S-1212-FD06-V32以及TJ4S-1212-FD06-V33、TJ4S-1212-FD06-V19，使用桶装甲醇通过专用新鲜溶剂气动隔膜泵经过管路139788-step02-P04加入180kg±5kg甲醇淋洗过滤干燥机主体以及捕集器；淋洗甲醇通过母液管路TJ4S-1212-FD06-P06转移母液储罐TJ4S-1212-FD06-Tr01放出（淋洗过程需保证泵后压力≥0.02MPa），取淋洗液进行Level Ⅱ检测；"] = (
        "Rinse sampling: at the fixed rinse port on the head of filter dryer TJ4S-1212-FD06, open valves TJ4S-1212-FD06-V25, TJ4S-1212-FD06-V26, TJ4S-1212-FD06-V32, TJ4S-1212-FD06-V33, and TJ4S-1212-FD06-V19 in turn. Charge 180±5 kg of drummed methanol through line 139788-step02-P04 with the dedicated fresh-solvent pneumatic diaphragm pump, and rinse the filter-dryer body and the trap. Discharge the methanol rinse from mother-liquor tank TJ4S-1212-FD06-Tr01 through TJ4S-1212-FD06-P06. Keep the pressure downstream of the pump at ≥0.02 MPa during rinsing, and sample the rinse for Level II testing."
    )
    L["擦拭取样：使用甲醇溶剂按照取样点评估示意图QM04053依次擦拭侧出料口、搅拌桨、取样手套箱、内表面"] = (
        "Swab sampling: using methanol, swab the side discharge port, agitator, sampling glove box, and inner surface in the order shown on sampling-point diagram QM04053."
    )
    L["1、通过反应釜TJ4S-1213-R15上的管线TJ4S-1213-R15-P02及进液口阀门TJ4S-1213-R15-V10、液体淋洗口阀门TJ4S-1213-R15-V66、TJ4S-1213-R15-V61、TJ4S-1213-R15-V60，使用桶装75%乙醇通过新鲜溶剂管路和新鲜溶剂专用气动隔膜泵加入175±10kg  75%乙醇淋洗反应釜（淋洗过程需保证泵后压力≥0.02MPa）；经过下口阀门TJ4S-1213-R15-V46、TJ4S-1213-R15-V47，通过阀门139788-step02-P04-015取淋洗液检测；再通过139788-step02-P04转移淋洗TJ4S-1212-FD06；"] = (
        "1. Through line TJ4S-1213-R15-P02, inlet valve TJ4S-1213-R15-V10, and liquid-rinse valves TJ4S-1213-R15-V66, TJ4S-1213-R15-V61, and TJ4S-1213-R15-V60 on reactor TJ4S-1213-R15, charge 175±10 kg of drummed 75% ethanol with the fresh-solvent line and the dedicated fresh-solvent pneumatic diaphragm pump, and rinse the reactor. Keep the pressure downstream of the pump at ≥0.02 MPa during rinsing. Take a rinse sample for testing through bottom valves TJ4S-1213-R15-V46 and TJ4S-1213-R15-V47 and valve 139788-step02-P04-015, then transfer the rinse to TJ4S-1212-FD06 through 139788-step02-P04."
    )
    L["2、通过反应釜TJ4S-1213-R15上的氮气管路TJ4S-1213-R15-P01，经过氮气过滤器TJ4S-1213-R15-NF01，开启氮气阀门TJ4S-1213-R10-V06、TJ4S-1213-R10-V07使用氮气对反应釜及釜下阀门进行吹干"] = (
        "2. Through nitrogen line TJ4S-1213-R15-P01 and nitrogen filter TJ4S-1213-R15-NF01 on reactor TJ4S-1213-R15, open nitrogen valves TJ4S-1213-R10-V06 and TJ4S-1213-R10-V07 and blow the reactor and the valves below it dry with nitrogen."
    )
    L["TJ4S-1213-R15内75%消毒液通过过滤干燥机TJ4S-1212-FD06上口固定淋洗口，依次分别开启阀门TJ4S-1212-FD06-V25、TJ4S-1212-FD06-V26、TJ4S-1212-FD06-V32以及TJ4S-1212-FD06-V33、TJ4S-1212-FD06-V19，淋洗过滤干燥机主体以及捕集器；通过母液管路TJ4S-1212-FD06-P06转移母液储罐TJ4S-1212-FD06-Tr01放出（淋洗过程需保证泵后压力≥0.02MPa），取淋洗液检测；"] = (
        "Send the 75% ethanol solution from TJ4S-1213-R15 to the fixed rinse port on the head of filter dryer TJ4S-1212-FD06. Open valves TJ4S-1212-FD06-V25, TJ4S-1212-FD06-V26, TJ4S-1212-FD06-V32, TJ4S-1212-FD06-V33, and TJ4S-1212-FD06-V19 in turn, and rinse the filter-dryer body and the trap. Discharge through mother-liquor line TJ4S-1212-FD06-P06 to tank TJ4S-1212-FD06-Tr01. Keep the pressure downstream of the pump at ≥0.02 MPa during rinsing, and sample the rinse for testing."
    )
    L["通过反应釜TJ4S-1213-R09上的固定淋洗口TJ4S-1213-R09-V77、TJ4S-1213-R09-V75、TJ4S-1213-R09-V10以及TJ4S-1213-R09-V76，使用桶装甲醇通过新鲜溶剂管路和新鲜溶剂专用气动隔膜泵加入175±10kg甲醇淋洗反应釜；（淋洗过程需保证泵后压力≥0.02MPa）"] = (
        "Through fixed rinse ports TJ4S-1213-R09-V77, TJ4S-1213-R09-V75, TJ4S-1213-R09-V10, and TJ4S-1213-R09-V76 on reactor TJ4S-1213-R09, charge 175±10 kg of drummed methanol with the fresh-solvent line and the dedicated fresh-solvent pneumatic diaphragm pump, and rinse the reactor. Keep the pressure downstream of the pump at ≥0.02 MPa during rinsing."
    )
    L["TJ4S-1200-F04压滤罐内甲醇通过母液转移管路139788-step02-P03、袋式过滤器TJ4S-1212-BF005，使用隔膜气动泵TJ4S-1200-PDP035经过管路139788-step02-P03转移至TJ4S-1213-R15；"] = (
        "Transfer the methanol in pressure filter TJ4S-1200-F04 to TJ4S-1213-R15 through mother-liquor line 139788-step02-P03 and bag filter TJ4S-1212-BF005, using pneumatic diaphragm pump TJ4S-1200-PDP035."
    )
    L["使用纯化水专用管路和专用质量流量计通过反应釜TJ4S-1213-R09上的固定淋洗口TJ4S-1213-R09-V77、TJ4S-1213-R09-V75、TJ4S-1213-R09-V10以及TJ4S-1213-R09-V76，加入175±10kg纯化水淋洗反应釜；"] = (
        "Through fixed rinse ports TJ4S-1213-R09-V77, TJ4S-1213-R09-V75, TJ4S-1213-R09-V10, and TJ4S-1213-R09-V76 on reactor TJ4S-1213-R09, charge 175±10 kg of purified water by the dedicated purified-water line and dedicated mass flowmeter, and rinse the reactor."
    )
    L["TJ4S-1200-F04压滤罐内纯化水通过母液转移管路139788-step02-P03、袋式过滤器TJ4S-1212-BF005，使用隔膜气动泵TJ4S-1200-PDP035经过管路139788-step02-P03转移至TJ4S-1213-R15"] = (
        "Transfer the purified water in pressure filter TJ4S-1200-F04 to TJ4S-1213-R15 through mother-liquor line 139788-step02-P03 and bag filter TJ4S-1212-BF005, using pneumatic diaphragm pump TJ4S-1200-PDP035."
    )
    L["TJ4S-1200-F04压滤罐内纯化水通过母液转移管路139788-step02-P03、袋式过滤器TJ4S-1212-BF005，使用隔膜气动泵TJ4S-1200-PDP035经过管路139788-step02-P03转移至TJ4S-1213-R15；"] = L[
        "TJ4S-1200-F04压滤罐内纯化水通过母液转移管路139788-step02-P03、袋式过滤器TJ4S-1212-BF005，使用隔膜气动泵TJ4S-1200-PDP035经过管路139788-step02-P03转移至TJ4S-1213-R15"
    ]
    L["冲洗完成后按照P&ID图所示，依次将反应釜上的所有管路和阀门进行拆卸，拆卸清洁要求参见如下表所示，反应釜的氮气/尾气管路及阀门、爆破片、固体加料阀、其他物料加料阀、视镜拆卸至最小节，使用纯化水擦拭至无物料残留，压缩气体吹干。"] = (
        "After the flush, remove every line and valve on the reactor in the order shown on the P&ID. Disassembly-cleaning requirements are in the table below. Disassemble the reactor nitrogen/vent lines and valves, rupture disc, solids-charging valve, other material-charging valves, and sight glass down to the smallest parts. Wipe with purified water until no material residue remains, and blow dry with compressed air."
    )
    L["设备罐体及配件验收完毕后，反向依次将视镜、其他物料加料阀、固体加料阀、爆破片、氮气/尾气管路及阀门进行安装，安装完毕后，使用氮气将搅拌罐和转移管路吹干"] = (
        "After the vessel and accessories are accepted, install the sight glass, other material-charging valves, solids-charging valve, rupture disc, and nitrogen/vent lines and valves in the reverse order. After installation, blow the mixing tank and transfer lines dry with nitrogen."
    )
    L["设备罐体及配件验收完毕后，反向依次将视镜、其他物料加料阀、固体加料阀、爆破片、氮气/尾气管路及阀门进行安装，安装完毕后，使用氮气将搅拌罐和转移管路吹干。"] = L[
        "设备罐体及配件验收完毕后，反向依次将视镜、其他物料加料阀、固体加料阀、爆破片、氮气/尾气管路及阀门进行安装，安装完毕后，使用氮气将搅拌罐和转移管路吹干"
    ]
    L["处理完成后，按照P&ID图所示，依次将搅拌罐上的所有管路和阀门进行拆卸，拆卸清洁要求参见如下表所示，搅拌罐的氮气/尾气管路及阀门、爆破片、固体加料阀、其他物料加料阀、"] = (
        "After this step, remove every line and valve on the mixing tank in the order shown on the P&ID. Disassembly-cleaning requirements are in the table below, covering the mixing-tank nitrogen/vent lines and valves, rupture disc, solids-charging valve, and other material-charging valves."
    )
    L["拆卸至最小节，使用纯化水擦拭至无物料残留，压缩气体吹干。"] = (
        "Disassemble down to the smallest parts, wipe with purified water until no material residue remains, and blow dry with compressed air."
    )
    L["配件安装完毕后，使用氮气对设备进行打压试密，氮气加压至0.07-0.09Mpa，保压≥30min，压力变化≤0.01Mpa"] = (
        "After the fittings are installed, leak-test the equipment with nitrogen. Raise the nitrogen pressure to 0.07–0.09 MPa, hold for ≥30 min, and confirm that the pressure change is ≤0.01 MPa."
    )
    L["通过专用纯化水气动隔膜泵和转移管路向搅拌罐内加入105±5kg 纯化水淋洗储罐内表面，通过下口管路TJ4S-1213-AT06-P04、TJ4S-1213-AT06-P05、TJ4S-1213-AT06-P03通过循环泵TJ4S-1213-AT06-MP01循环淌洗15~30min；再通过管路TJ4S-1213-AT06-P04的排净阀门TJ4S-1213-AT06-V03放出至废液罐。"] = (
        "Charge 105±5 kg of purified water into the mixing tank with the dedicated purified-water pneumatic diaphragm pump and the transfer line, and rinse the inner surface. Circulate through bottom lines TJ4S-1213-AT06-P04, TJ4S-1213-AT06-P05, and TJ4S-1213-AT06-P03 with circulation pump TJ4S-1213-AT06-MP01 for 15–30 min, then discharge to the waste tank through drain valve TJ4S-1213-AT06-V03 on line TJ4S-1213-AT06-P04."
    )
    L["通过哈式合金压滤罐TJ4S-1200-F04上的固定淋洗口阀门TJ4S-1200-F04-V18，使用桶装甲醇通过专用新鲜溶剂加料泵向压滤罐内加入77.5Kg±2.5kg甲醇淋洗；通过管路139788-step02-P03转移至TJ4S-1213-R15、经过下口阀门TJ4S-1213-R15-V46、TJ4S-1213-R15-V47，通过阀门139788-step02-P04-015放入废液桶中"] = (
        "Through fixed rinse valve TJ4S-1200-F04-V18 on Hastelloy pressure filter TJ4S-1200-F04, charge 77.5±2.5 kg of drummed methanol with the dedicated fresh-solvent charging pump and rinse the filter. Transfer it to TJ4S-1213-R15 through line 139788-step02-P03, then discharge it into a waste drum through bottom valves TJ4S-1213-R15-V46 and TJ4S-1213-R15-V47 and valve 139788-step02-P04-015."
    )
    L["通过哈式合金压滤罐TJ4S-1200-F04上的固定淋洗口阀门TJ4S-1200-F04-V18，通过专用纯化水加料泵向压滤罐内加入95.0±5.0kg纯化水淋洗至无溶剂气味残留，通过管路139788-step02-P03转移至TJ4S-1213-R15、经过下口阀门TJ4S-1213-R15-V46、TJ4S-1213-R15-V47，通过阀门139788-step02-P04-015放入废液桶中"] = (
        "Through fixed rinse valve TJ4S-1200-F04-V18 on Hastelloy pressure filter TJ4S-1200-F04, charge 95.0±5.0 kg of purified water with the dedicated purified-water charging pump and rinse until no solvent odor remains. Transfer it to TJ4S-1213-R15 through line 139788-step02-P03, then discharge it into a waste drum through bottom valves TJ4S-1213-R15-V46 and TJ4S-1213-R15-V47 and valve 139788-step02-P04-015."
    )
    L["该管路为长距离转移管路，中间无法兰接口，此管路在TJ4S-1200-F04主体设备清洁内，清洗方式，详见表34中139788-step02-P03在线清洗部分"] = (
        "This is a long transfer line with no flange joint along its length. It is cleaned with the main equipment TJ4S-1200-F04. See the online cleaning of 139788-step02-P03 in Table 34."
    )
    L["TJ4S-1213-R10-V47至TJ4S-1213-R09-P04-001"] = "TJ4S-1213-R10-V47 to TJ4S-1213-R09-P04-001"
    L["TJ4S-1213-R09-V74至TJ4S-1200-F04罐上口"] = "TJ4S-1213-R09-V74 to the top nozzle of tank TJ4S-1200-F04"
    L["TJ4S-1212-F04下口至TJ4S-1213-R15-V62上口阀门"] = "Bottom outlet of TJ4S-1212-F04 to the upstream side of valve TJ4S-1213-R15-V62"
    L["TJ4S-1213-R15-V47至TJ4S-1212-FD06"] = "TJ4S-1213-R15-V47 to TJ4S-1212-FD06"
    L["溶剂管道釜上口阀门TJ4S-1213-R10-V10接移动隔膜泵"] = (
        "Reactor-head solvent-line valve TJ4S-1213-R10-V10, connected to a mobile diaphragm pump"
    )
    L["反应釜氮气管路至TJ4S-1213-R09-V06"] = "Reactor nitrogen line up to TJ4S-1213-R09-V06"
    L["反应釜氮气进釜管路至TJ4S-1213-R10-V06"] = "Reactor nitrogen inlet line up to TJ4S-1213-R10-V06"
    L["接收罐废液管道口至TJ4S-1213-R09-V25"] = "Receiver waste-line nozzle up to TJ4S-1213-R09-V25"
    L["TJ4S-1213-AT06搅拌罐到TJ4S-1213-AT06-E01换热器"] = (
        "Mixing tank TJ4S-1213-AT06 to heat exchanger TJ4S-1213-AT06-E01"
    )
    L["TJ4S-1213-R15管路"] = "TJ4S-1213-R15 line"
    return L


LONG.update(add_long())

# The methanol transfer without -004 was overwritten incorrectly above. Set the exact sentence.
LONG["TJ4S-1213-R10内淋洗甲醇通过下口阀门TJ4S-1213-R10-V46、TJ4S-1213-R10-V47以及釜下转移管路139788-step02-P01，使用隔膜气动泵TJ4S-1200-PDP025通过反应釜进液管路TJ4S-1213-R09-P04及阀门TJ4S-1213-R09-V65、TJ4S-1213-R09-V60；TJ4S-1213-R09-V61转移至 TJ4S-1213-R09"] = (
    "Transfer the methanol rinse from TJ4S-1213-R10 to TJ4S-1213-R09 through bottom valves TJ4S-1213-R10-V46 and TJ4S-1213-R10-V47 and bottom transfer line 139788-step02-P01, using pneumatic diaphragm pump TJ4S-1200-PDP025, reactor inlet line TJ4S-1213-R09-P04, and valves TJ4S-1213-R09-V65, TJ4S-1213-R09-V60, and TJ4S-1213-R09-V61."
)
LONG["TJ4S-1213-R10内淋洗甲醇通过下口阀门TJ4S-1213-R10-V46、TJ4S-1213-R10-V47以及釜下转移管路139788-step02-P01，使用隔膜气动泵TJ4S-1200-PDP025通过反应釜进液管路TJ4S-1213-R09-P04及阀门TJ4S-1213-R09-V65、TJ4S-1213-R09-V60；TJ4S-1213-R09-V61转移至 TJ4S-1213-R09；"] = LONG[
    "TJ4S-1213-R10内淋洗甲醇通过下口阀门TJ4S-1213-R10-V46、TJ4S-1213-R10-V47以及釜下转移管路139788-step02-P01，使用隔膜气动泵TJ4S-1200-PDP025通过反应釜进液管路TJ4S-1213-R09-P04及阀门TJ4S-1213-R09-V65、TJ4S-1213-R09-V60；TJ4S-1213-R09-V61转移至 TJ4S-1213-R09"
]


def see_table(s):
    if "详见" not in s and "详见表" not in s:
        return None
    m = re.search(r"表\s*(\d+)", s)
    if not m:
        return None
    table = m.group(1)
    ids = re.findall(r"(?:TJ4S|139788)[-A-Za-z0-9]+", s)
    ident = ids[-1] if ids else ""
    if s.startswith("此阀门"):
        who = "This valve"
    elif s.startswith("此管路"):
        who = "This line"
    elif s.startswith("此部件"):
        who = "This part"
    elif s.startswith("此泵"):
        who = "This pump"
    elif s.startswith("此视镜"):
        who = "This sight glass"
    else:
        who = "This item"
    how = "by nitrogen purge" if "氮气吹扫" in s else "in the online cleaning section" if "在线" in s else "in the cleaning section"
    where = f" for {ident}" if ident else ""
    return f"{who} is cleaned with the main equipment. See Table {table}{where}, {how}."


def not_applicable(s):
    if "不涉及" not in s:
        return None
    bits = []
    if "回流" in s and "蒸馏" in s:
        bits.append("reflux and distillation are not used")
    elif "蒸馏" in s:
        bits.append("distillation is not used")
    if "未使用接收罐" in s or "未使用接收" in s:
        bits.append("the receiver is not used for this project")
    if "未接触物料" in s or "未直接接触" in s or "不直接接触" in s:
        bits.append("there is no product contact")
    if "固体投料" in s:
        bits.append("solids charging is not used")
    if not bits:
        bits.append("it is outside the scope of this cleaning")
    return "Cleaning is not applicable: " + "; ".join(bits) + "."


WIPE = [
    ("解体后阀门及球体，表面", "the surfaces of the disassembled valve and ball"),
    ("解体后阀门阀体、阀盖以及阀芯，表面", "the surfaces of the disassembled valve body, bonnet, and plug"),
    ("解体后玻璃以及密封垫片表面，两端不锈钢头", "the disassembled glass, gasket surfaces, and stainless-steel ends"),
    ("解体后玻璃以及密封垫片表面，两端法兰", "the disassembled glass, gasket surfaces, and flanges at both ends"),
    ("解体后玻璃、人孔法兰以及密封垫片表面", "the disassembled glass, manway flange, and gasket surfaces"),
    ("拆除后泵腔、膜片、阻液球内外表面", "the inner and outer surfaces of the removed pump chamber, diaphragm, and check balls"),
    ("拆除的过滤器底座及滤壳内表面", "the inner surfaces of the removed filter base and housing"),
    ("拆除后管件内外表面，密封件表面", "the inner and outer surfaces of the removed fitting and the seal surfaces"),
    ("拆除的三通，密封件表面", "the removed tee and seal surfaces"),
    ("拆除的弯头，密封件表面", "the removed elbow and seal surfaces"),
    ("拆底管内外表面", "the inner and outer surfaces of the removed dip pipe"),
    ("拆除后Y型管", "the removed Y-pipe"),
    ("异径管部件", "the reducer"),
    ("变径部件", "the reducer"),
    ("弯头部件", "the elbow"),
    ("三通部件", "the tee"),
    ("四通部件", "the cross"),
    ("管道部件", "the pipe"),
    ("管径部件", "the pipe fitting"),
    ("套筒内外表面以及密封连接处", "the inner and outer surfaces of the housing and the sealing joints"),
    ("人孔内外表面、人孔垫片", "the inner and outer surfaces of the manway and the manway gasket"),
    ("隔膜阀门两端法兰将阀门内表面", "the inner surface of the diaphragm valve from both flanges"),
]


def wipe(s):
    if "擦拭" not in s and "擦洗" not in s:
        return None
    qa = "QA acceptance" if ("QA验收" in s or s.rstrip().endswith("验收")) else ""
    end = ""
    if "三通三端" in s:
        blow = "blow dry with compressed air alternately from all three ends of the tee"
    elif "三通两端" in s:
        blow = "blow dry with compressed air alternately from both ends of the tee"
    elif "弯头两端" in s:
        blow = "blow dry with compressed air alternately from both ends of the elbow"
    elif "Y型管两端" in s:
        blow = "blow dry with compressed air alternately from both ends of the Y-pipe"
    elif "管件两端" in s or "加料管路两端" in s:
        blow = "blow dry with compressed air alternately from both ends"
    elif "交替" in s:
        blow = "blow dry alternately with compressed air"
    else:
        blow = "blow dry with compressed air"
    obj = None
    for key, name in WIPE:
        if key in s:
            obj = name
            break
    if obj is None:
        return None
    cloth = "a clean wipe moistened with purified water" if "浸湿" in s else "a clean wipe dipped in purified water"
    extra = ""
    if "膜片" in s and "破损" in s:
        extra = " Check the diaphragm for damage, aging, or pinholes, and replace it if any is found."
    tail = ", " + qa if qa else ""
    punct = ";" if s.rstrip().endswith("；") else ""
    return f"Wipe {obj} with {cloth} until no material residue is visible, {blow}{tail}{extra}{punct}"


AB = (
    "Lock the active and passive halves of the split butterfly valve together. Wipe the inner and outer surfaces and the perimeter of the valve all over with a clean wipe well soaked in purified water. Open the disc while wiping and wipe the disc faces, until no material residue is visible under a high-intensity flashlight."
)


def alpha_beta(s):
    if "αβ" not in s and "αβ" not in s:
        return None
    if "被动阀拆卸" in s:
        return (
            "Remove the passive half of the split butterfly valve. Wipe the sealing face and the gasket with a clean wipe well soaked in purified water until no material residue is visible under a high-intensity flashlight. Remove residual water from the removed parts with a dry clean wipe, then blow them completely dry with compressed air."
        )
    if "复位" in s:
        return "Reassemble the split butterfly valve in the reverse order."
    if "主动阀" in s:
        return AB
    return None


def wrench(s):
    if not s.startswith("使用不锈钢扳手") and not s.startswith("使用扳手"):
        return None
    known = {
        "使用不锈钢扳手将阀门四个固定螺杆的螺母逆时针方向拧松，取下四个螺杆，将阀门解体并将阀球取出；":
            "Use a stainless steel wrench to loosen the nuts on the four fixing bolts counterclockwise, remove the four bolts, disassemble the valve, and take out the ball.",
        "使用不锈钢扳手将阀门四个固定螺杆的螺母逆时针方向拧松，取下四个螺杆，将视镜解体并将玻璃垫片取出；":
            "Use a stainless steel wrench to loosen the nuts on the four fixing bolts counterclockwise, remove the four bolts, disassemble the sight glass, and take out the glass and gasket.",
        "使用不锈钢扳手将阀门四个固定螺杆逆时针方向拧松，取下四个螺杆及下方螺母，将球阀解体并将阀球取出；":
            "Use a stainless steel wrench to loosen the four fixing bolts counterclockwise, remove the four bolts and the nuts underneath, disassemble the ball valve, and take out the ball.",
        "使用不锈钢扳手将阀门四个固定螺杆的螺母逆时针方向拧松，取下四个螺杆，将视镜解体并将玻璃管取出；":
            "Use a stainless steel wrench to loosen the nuts on the four fixing bolts counterclockwise, remove the four bolts, disassemble the sight glass, and take out the glass tube.",
        "使用不锈钢扳手将固定视镜的8个螺母逆时针方向拧松，取下螺杆，将视镜解体并将玻璃垫片取出；":
            "Use a stainless steel wrench to loosen the eight nuts holding the sight glass counterclockwise, remove the bolts, disassemble the sight glass, and take out the glass and gasket.",
        "使用不锈钢扳手将固定视镜8个螺母逆时针方向拧松，取下螺杆，将视镜解体并将玻璃垫片取出；":
            "Use a stainless steel wrench to loosen the eight nuts holding the sight glass counterclockwise, remove the bolts, disassemble the sight glass, and take out the glass and gasket.",
        "使用不锈钢扳手将氮气过滤器套筒快开卡箍固定螺栓逆时针松动取下，打开过滤器套筒，将滤芯取下废弃":
            "Use a stainless steel wrench to loosen and remove the tri-clamp bolt on the nitrogen-filter housing counterclockwise, open the housing, and remove and discard the cartridge.",
        "使用不锈钢扳手将袋式过滤器套筒卡箍固定螺栓逆时针松动取下，打开过滤器套筒，将滤袋取下废弃":
            "Use a stainless steel wrench to loosen and remove the clamp bolt on the bag-filter housing counterclockwise, open the housing, and remove and discard the filter bag.",
        "使用不锈钢扳手将异径管两端连接法兰断开，将异径管拆除":
            "Use a stainless steel wrench to break the connecting flanges at both ends of the reducer and remove the reducer.",
        "使用不锈钢扳手将与弯头连接端的法兰固定螺杆逆时针方向旋转松动;":
            "Use a stainless steel wrench to loosen the flange bolts at the elbow connection counterclockwise.",
        "使用不锈钢扳手将弯头上下两端的快开卡盘松动拆卸，取下弯头部件;":
            "Use a stainless steel wrench to loosen and remove the tri-clamp connections at both ends of the elbow, and remove the elbow.",
        "使用不锈钢扳手将气动隔膜泵上下侧进出液腔连接法兰螺栓逆时针拧下，取下进出液腔管，取下阻液球，使用扳手将两端泵壳螺栓逆时针拧下，打开泵腔，取下膜片":
            "Use a stainless steel wrench to unscrew the flange bolts on the inlet and outlet chambers of the pneumatic diaphragm pump counterclockwise, remove the chamber pipes and the check balls, unscrew the housing bolts at both ends counterclockwise, open the pump chamber, and remove the diaphragm.",
        "使用不锈钢扳手将球阀丝扣连接端逆时针拧开取下，取出阀芯球体":
            "Use a stainless steel wrench to unscrew the threaded end of the ball valve counterclockwise and take out the ball.",
        "使用不锈钢扳手将阀门两端连接法兰断开，将隔膜阀体取出；":
            "Use a stainless steel wrench to break the flanges at both ends of the valve and remove the diaphragm-valve body.",
        "使用不锈钢扳手将流体过滤器的卡盘松开，取下底座和滤壳":
            "Use a stainless steel wrench to release the clamp of the liquid filter and remove the base and housing.",
        "使用不锈钢扳手将止回阀阀体上的四个螺母逆时针方向拧松，解体拆除阀体和阀盖，取出止回阀芯":
            "Use a stainless steel wrench to loosen the four nuts on the check-valve body counterclockwise, separate the body and cover, and take out the check-valve plug.",
        "使用不锈钢扳手将人孔盖周围十二个固定卡箍拧松，逐个取下卡箍，取下人孔盖；":
            "Use a stainless steel wrench to loosen the twelve clamps around the manway cover, remove the clamps one by one, and remove the cover.",
        "使用不锈钢扳手将与阀门的法兰固定螺杆逆时针方向旋转松动;":
            "Use a stainless steel wrench to loosen the flange bolts at the valve connection counterclockwise.",
        "使用不锈钢扳手将氮气过滤器与管路上下两端的快开卡盘松动拆卸，取下管径部件;":
            "Use a stainless steel wrench to loosen and remove the tri-clamp connections between the nitrogen filter and the line at both ends, and remove the pipe fitting.",
        "使用不锈钢扳手将管路上下两端的快开卡盘松动拆卸，取下管道部件;":
            "Use a stainless steel wrench to loosen and remove the tri-clamp connections at both ends of the line, and remove the pipe section.",
        "使用不锈钢扳手将管道两端连接快开卡箍断开，将管道拆除":
            "Use a stainless steel wrench to open the tri-clamps at both ends of the pipe and remove the pipe.",
        "使用不锈钢扳手将变径两端连接快开卡箍断开，将管道拆除":
            "Use a stainless steel wrench to open the tri-clamps at both ends of the reducer and remove the pipe.",
    }
    return known.get(s.rstrip()) or known.get(s)


REASS = {
    "将拆解的阀门依次完成复位组装，更换阀门两端连接的四氟法兰垫片":
        "Reassemble the valve in order and replace the PTFE flange gaskets at both ends.",
    "将拆解的阀门依次完成复位组装，更换阀门两端连接的法兰垫片":
        "Reassemble the valve in order and replace the flange gaskets at both ends.",
    "将拆解的阀门依次完成复位组装，更换阀门两端连接的快开垫片":
        "Reassemble the valve in order and replace the tri-clamp gaskets at both ends.",
    "将拆解的视镜依次完成复位组装，更换阀门两端连接的法兰垫片":
        "Reassemble the sight glass in order and replace the flange gaskets at both ends of the connection.",
    "将拆解的视镜依次完成复位组装，更换视镜两侧连接的法兰垫片":
        "Reassemble the sight glass in order and replace the flange gaskets on both sides.",
    "将拆解的视镜依次完成复位组装，更换阀门两端连接的快开垫片":
        "Reassemble the sight glass in order and replace the tri-clamp gaskets at both ends of the connection.",
    "将拆解的异径管依次完成复位组装，更换管件两端连接的法兰垫片":
        "Reassemble the reducer in order and replace the flange gaskets at both ends.",
    "将拆解的弯头依次完成复位组装，更换管件两端连接的快开垫片":
        "Reassemble the elbow in order and replace the tri-clamp gaskets at both ends.",
    "将拆解的弯头依次完成复位组装，更换管件两端连接的法兰垫片":
        "Reassemble the elbow in order and replace the flange gaskets at both ends.",
    "将拆解的弯头依次完成复位组装，更换管件各端连接的法兰垫片":
        "Reassemble the elbow in order and replace the flange gasket at each end.",
    "将拆解的三通依次完成复位组装，更换管件各端连接的法兰垫片":
        "Reassemble the tee in order and replace the flange gasket at each end.",
    "将拆解的管道依次完成复位组装，更换管件两端连接的法兰垫片":
        "Reassemble the pipe in order and replace the flange gaskets at both ends.",
    "将拆解的管道依次完成复位组装，更换管件两端连接的快开垫片":
        "Reassemble the pipe in order and replace the tri-clamp gaskets at both ends.",
    "将拆解的管道依次完成复位组装，更换管件各端连接的法兰垫片":
        "Reassemble the pipe in order and replace the flange gasket at each end.",
    "将拆解的变径依次完成复位组装，更换管件各端连接的法兰垫片":
        "Reassemble the reducer in order and replace the flange gasket at each end.",
    "将拆解的四通依次完成复位组装，更换管件各端连接的法兰垫片":
        "Reassemble the cross in order and replace the flange gasket at each end.",
    "将拆解的管径依次完成复位组装，更换管件两端连接的快开垫片":
        "Reassemble the pipe fitting in order and replace the tri-clamp gaskets at both ends.",
    "安装新的折叠滤芯，将套筒依次完成复位组装，并更换过滤器两端连接的垫片":
        "Install a new pleated cartridge, reassemble the housing in order, and replace the gaskets at both ends of the filter.",
    "安装新的滤芯，将套筒依次完成复位组装，并更换过滤器两端连接的垫片":
        "Install a new cartridge, reassemble the housing in order, and replace the gaskets at both ends of the filter.",
    "安装新的滤袋，将套筒依次完成复位组装，并更换过滤器两端连接的垫片":
        "Install a new filter bag, reassemble the housing in order, and replace the gaskets at both ends of the filter.",
    "更换阀门两端连接的四氟法兰垫片，将隔膜阀门复位组装。":
        "Replace the PTFE flange gaskets at both ends and reassemble the diaphragm valve.",
    "PTFE lined衬PTFE 将拆解的阀门依次完成复位组装，更换阀门两端连接的四氟法兰垫片":
        "Reassemble the valve in order and replace the PTFE flange gaskets at both ends.",
}


def table_title(s):
    m = re.match(r"^表\s*(\d+)[，,](.+)$", s)
    if not m:
        return None
    n, rest = m.group(1), m.group(2).strip()
    # equipment train captions
    if "设备链" in rest:
        rest_en = rest
        rest_en = rest_en.replace("搪玻璃反应釜", "glass-lined reactor ")
        rest_en = rest_en.replace("至", " to ")
        rest_en = rest_en.replace("到", " to ")
        rest_en = rest_en.replace("压滤罐", "pressure filter ")
        rest_en = rest_en.replace("反应釜", "reactor ")
        rest_en = rest_en.replace("过滤干燥机", "filter dryer ")
        rest_en = rest_en.replace("设备链管道以及阀门部件清洁如下", "equipment-train lines and valve parts, as follows")
        rest_en = rest_en.replace("设备链管道以及阀门部件的清洁如下", "equipment-train lines and valve parts, as follows")
        rest_en = rest_en.replace("设备链管道以及附阀门部件清洁如下", "equipment-train lines and associated valve parts, as follows")
        rest_en = re.sub(r"\s+", " ", rest_en).strip()
        return f"Table {n}. Cleaning of {rest_en}"
    mapped = {
        "切割釜主体清洗": "Cleaning of the cleavage-reactor body",
        "切割釜上安装的阀门及部件清洗": "Cleaning of valves and components on the cleavage reactor",
        "切割釜自身配备的管路清洗": "Cleaning of lines on the cleavage reactor",
        "切割釜设备安装确认": "Installation check of the cleavage reactor",
        "配液釜主体清洗": "Cleaning of the solution-preparation reactor body",
        "配液釜阀门及部件的清洗": "Cleaning of valves and components on the solution-preparation reactor",
        "配液釜自身配备的管路清洗": "Cleaning of lines on the solution-preparation reactor",
        "配液釜自身管件的清洗": "Cleaning of fittings on the solution-preparation reactor",
        "配液釜设备安装确认": "Installation check of the solution-preparation reactor",
        "压滤罐主体清洗": "Cleaning of the pressure-filter body",
        "压滤罐设备完整性检查": "Integrity check of the pressure filter",
        "压滤罐自身配件清洗": "Cleaning of pressure-filter accessories",
        "压滤罐设备安装确认": "Installation check of the pressure filter",
        "沉淀釜主体清洗": "Cleaning of the precipitation-reactor body",
        "沉淀釜阀门及配件清洗": "Cleaning of valves and fittings on the precipitation reactor",
        "沉淀釜自身管线清洗": "Cleaning of precipitation-reactor lines",
        "沉淀釜设备安装确认": "Installation check of the precipitation reactor",
        "过滤干燥机主体清洗": "Cleaning of the filter-dryer body",
        "过滤干燥机主体配件清洁": "Cleaning of filter-dryer accessories",
        "过滤干燥机自身管路清洁": "Cleaning of filter-dryer lines",
        "过滤干燥机设备安装确认": "Installation check of the filter dryer",
        "搅拌罐主体清洗": "Cleaning of the mixing-tank body",
        "搅拌罐主体配件清洁": "Cleaning of mixing-tank accessories",
        "搅拌罐自身管路清洁": "Cleaning of mixing-tank lines",
        "搅拌罐设备安装确认": "Installation check of the mixing tank",
        "设备及设备链消毒": "Disinfection of the equipment and equipment train",
        "设备清洁后取样信息": "Sampling after equipment cleaning",
        "设备相关耗材清洁方式如下": "Cleaning of equipment consumables",
        "设备清洁后取样位置及方式如下": "Sampling locations and methods after equipment cleaning",
        "切割釜设备清洁拆卸过程中完整性检查": "Integrity check during disassembly cleaning of the cleavage reactor",
        "配液釜设备清洁拆卸过程中完整性检查": "Integrity check during disassembly cleaning of the solution-preparation reactor",
        "压滤罐设备清洁拆卸过程中完整性检查": "Integrity check during disassembly cleaning of the pressure filter",
        "沉淀釜设备清洁拆卸过程中的完整性检查": "Integrity check during disassembly cleaning of the precipitation reactor",
        "过滤干燥机设备清洁拆卸过程中的完整性检查": "Integrity check during disassembly cleaning of the filter dryer",
        "搅拌罐设备清洁拆卸过程中完整性检查": "Integrity check during disassembly cleaning of the mixing tank",
    }
    if rest in mapped:
        return f"Table {n}. {mapped[rest]}"
    return None


def section_heading(s):
    m = re.match(r"^(\d+(?:\.\d+)*)\s*(.+)$", s)
    if not m:
        return None
    num, rest = m.group(1), m.group(2).strip()
    rest = re.sub(r"\s+", " ", rest)
    known = {
        "切割釜主体清洗": "Cleaning of the cleavage-reactor body",
        "切割釜主体的清洁": "Cleaning of the cleavage-reactor body",
        "切割釜上安装的阀门及部件清洗": "Cleaning of valves and components on the cleavage reactor",
        "切割釜设备安装确认": "Installation check of the cleavage reactor",
        "切割釜设备清洁拆卸过程中的完整性检查": "Integrity check during disassembly cleaning of the cleavage reactor",
        "切割釜自身配件清洁": "Cleaning of accessories on the cleavage reactor",
        "配液釜主体清洗": "Cleaning of the solution-preparation reactor body",
        "配液釜设备清洁拆卸过程中的完整性检查": "Integrity check during disassembly cleaning of the solution-preparation reactor",
        "配液釜自身配件清洁": "Cleaning of accessories on the solution-preparation reactor",
        "配液釜设备安装确认": "Installation check of the solution-preparation reactor",
        "压滤罐主体清洗": "Cleaning of the pressure-filter body",
        "压滤罐设备清洁拆卸过程中完整性检查": "Integrity check during disassembly cleaning of the pressure filter",
        "压滤罐自身配件清洗": "Cleaning of pressure-filter accessories",
        "压滤罐设备安装确认": "Installation check of the pressure filter",
        "沉淀釜主体清洗": "Cleaning of the precipitation-reactor body",
        "沉淀釜设备清洁拆卸过程中的完整性检查": "Integrity check during disassembly cleaning of the precipitation reactor",
        "沉淀釜自身配件清洁": "Cleaning of accessories on the precipitation reactor",
        "沉淀釜设备安装确认": "Installation check of the precipitation reactor",
        "过滤干燥机主体清洗": "Cleaning of the filter-dryer body",
        "过滤干燥机设备清洁拆卸过程中的完整性检查": "Integrity check during disassembly cleaning of the filter dryer",
        "过滤干燥机自身配件清洁": "Cleaning of filter-dryer accessories",
        "过滤干燥机设备安装确认": "Installation check of the filter dryer",
        "搅拌罐主体清洗": "Cleaning of the mixing-tank body",
        "搅拌罐设备清洁拆卸过程中的完整性检查": "Integrity check during disassembly cleaning of the mixing tank",
        "搅拌罐自身配件清洁": "Cleaning of accessories on the mixing tank",
        "搅拌罐设备安装确认": "Installation check of the mixing tank",
        "设备链中的每一个部件的清洁": "Cleaning of each component in the equipment train",
    }
    if rest in known:
        return f"{num} {known[rest]}"
    return None


_NORM = None


def english_for(s):
    global _NORM
    s = s.strip().lstrip("/")
    if not s or s == "使" or not HAN.search(s):
        return None
    if _NORM is None:
        _NORM = {}
        for d in (LONG, SHORT, REASS):
            for k, v in d.items():
                if v and v != "SKIP":
                    _NORM[re.sub(r"\s+", "", k)] = v
    hit = _NORM.get(re.sub(r"\s+", "", s))
    if hit:
        return hit
    if s in LONG:
        return LONG[s]
    if s in SHORT:
        return SHORT[s]
    if s in REASS:
        return REASS[s]
    hit = wrench(s) or alpha_beta(s) or wipe(s) or see_table(s) or not_applicable(s) or table_title(s) or section_heading(s)
    if hit:
        return hit
    # Chinese after a translated label, e.g. 目标：...
    if s.startswith("目标："):
        return LONG.get(s)
    if s.startswith("范围："):
        return LONG.get(s)
    return None


RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
             "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
             "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
             "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
             "specVanish", "oMath"]


def child(parent, name):
    el = parent.find(W + name)
    if el is not None:
        return el
    el = etree.Element(W + name)
    rank = RPR_ORDER.index(name)
    for i, c in enumerate(parent):
        local = etree.QName(c).localname
        if local in RPR_ORDER and RPR_ORDER.index(local) > rank:
            parent.insert(i, el)
            return el
    parent.append(el)
    return el


def ptext(p):
    return "".join(t.text or "" for t in p.iter(W + "t"))


def in_table(p):
    return any(a.tag == W + "tc" for a in p.iterancestors())


def replace_span(p, old, new):
    done = 0
    start = 0
    while True:
        nodes = list(p.iter(W + "t"))
        blob = "".join(t.text or "" for t in nodes)
        i = blob.find(old, start)
        if i < 0:
            return done
        remaining, pos, first = len(old), 0, True
        for t in nodes:
            txt = t.text or ""
            if remaining and pos + len(txt) > i:
                k = max(i - pos, 0)
                take = min(len(txt) - k, remaining)
                t.text = txt[:k] + (new if first else "") + txt[k + take:]
                if t.text[:1].isspace() or t.text[-1:].isspace():
                    t.set(XML_SPACE, "preserve")
                first = False
                remaining -= take
            pos += len(txt)
        done += 1
        start = i + len(new)


def english_run(ref, text, table):
    r = etree.Element(W + "r")
    ref_rpr = ref.find(W + "rPr")
    rpr = copy.deepcopy(ref_rpr) if ref_rpr is not None else etree.Element(W + "rPr")
    r.append(rpr)
    fonts = child(rpr, "rFonts")
    for a in ("ascii", "hAnsi", "cs"):
        fonts.set(W + a, EN_FONT)
    fonts.attrib.pop(W + "hint", None)
    if table:
        child(rpr, "sz").set(W + "val", TABLE_SZ)
        child(rpr, "szCs").set(W + "val", TABLE_SZ)
    t = etree.SubElement(r, W + "t")
    t.text = text
    t.set(XML_SPACE, "preserve")
    return r


def insert_at(p, offset, text, table):
    pos = 0
    nodes = list(p.iter(W + "t"))
    if not nodes:
        return False
    for t in nodes:
        n = len(t.text or "")
        if pos <= offset <= pos + n:
            k, r = offset - pos, t.getparent()
            if k:
                left = copy.deepcopy(r)
                idx = list(r).index(t)
                for c in list(left)[idx + 1:]:
                    left.remove(c)
                left[idx].text = (t.text or "")[:k]
                left[idx].set(XML_SPACE, "preserve")
                for c in list(r)[:idx]:
                    if c.tag != W + "rPr":
                        r.remove(c)
                t.text = (t.text or "")[k:]
                t.set(XML_SPACE, "preserve")
                r.addprevious(left)
            r.addprevious(english_run(r, text, table))
            return True
        pos += n
    return False


def apply_paragraph(p, missing):
    raw = ptext(p)
    s = raw.strip()
    if not s or not HAN.search(s):
        return
    if "/" in s and re.search(r"[A-Za-z]", s.split("/", 1)[0]) and len(s.split("/", 1)[1]) < 24 and "。" not in s and "；" not in s:
        return
    if ALREADY.search(s):
        # still translate a Chinese sentence that follows a short label
        m = re.match(r"^[^/\u4e00]*[A-Za-z][^/]*/[^：:\n]{0,12}[：:]", s)
        if m and HAN.search(s[m.end():]) and len(s[m.end():]) > 12:
            en = english_for(s[m.end():])
            if en:
                insert_at(p, raw.find(s[m.end():]), en + "/", in_table(p))
            else:
                missing.append(s[m.end():])
        return
    han = HAN.search(s)
    lead = s[: han.start()] if han else ""
    # A word-only prefix such as "Butterfly valve" is a label glued to Chinese.
    # Do not treat an equipment number or "Step 2" as a label.
    word_lead = bool(re.fullmatch(r"[A-Za-z][A-Za-z .+\-]*", lead.strip())) if lead.strip() else False
    target = s[han.start():] if word_lead else s
    en = english_for(target)
    if not en:
        missing.append(s)
        return
    if word_lead:
        replace_span(p, lead, en + "/")
    else:
        insert_at(p, raw.find(s), en + "/", in_table(p))


def main():
    data = SRC.read_bytes()
    zin = zipfile.ZipFile(BytesIO(data))
    xml = zin.read("word/document.xml")
    root = etree.fromstring(xml)
    for p in list(root.iter(W + "p")):
        for old, new in FIXES:
            if old in ptext(p):
                replace_span(p, old, new)
    missing = []
    for p in list(root.iter(W + "p")):
        apply_paragraph(p, missing)
    # second pass in case a fix exposed Chinese
    if missing:
        still = []
        seen = set(missing)
        for p in list(root.iter(W + "p")):
            s = ptext(p).strip()
            if s in seen and HAN.search(s) and not ALREADY.search(s):
                en = english_for(s.lstrip("/"))
                if en:
                    apply_paragraph(p, [])
                else:
                    still.append(s)
        missing = still
    uniq = []
    got = set()
    for s in missing:
        if s not in got:
            got.add(s)
            uniq.append(s)
    out = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w") as zout:
        for item in zin.infolist():
            payload = out if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(item, payload)
    zin.close()
    SRC.write_bytes(buf.getvalue())
    print("missing", len(uniq))
    Path("/tmp/sop2_missing.txt").write_text("\n---\n".join(uniq))
    if uniq:
        for s in uniq[:40]:
            print("MISS", s[:180])


if __name__ == "__main__":
    main()
