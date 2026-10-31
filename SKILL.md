---
AIGC:
    Label: "1"
    ContentProducer: 001191440300708461136T1XGW3
    ProduceID: 89f2995f43cde2f2717e466bdd92de27_22025a54d51311f19101525400de85a5
    ReservedCode1: UV91IQUAmVrAL43mVFqDPKHW8WMa8CA0gDShijqmEN6//z9K2MBZSlnClLrlSUBVQCD2qVBKeN3d8us7XRorwQV/NSw0VNNF+egVo/JgeOGZmW3MPtKJoAdaswfC/ZFS094iPegeGm8CuxcHXsJdLlE4xJ/K/E9FQLH1PjHaZd3SZztCxv06p5t67Qw=
    ContentPropagator: 001191440300708461136T1XGW3
    PropagateID: 89f2995f43cde2f2717e466bdd92de27_22025a54d51311f19101525400de85a5
    ReservedCode2: UV91IQUAmVrAL43mVFqDPKHW8WMa8CA0gDShijqmEN6//z9K2MBZSlnClLrlSUBVQCD2qVBKeN3d8us7XRorwQV/NSw0VNNF+egVo/JgeOGZmW3MPtKJoAdaswfC/ZFS094iPegeGm8CuxcHXsJdLlE4xJ/K/E9FQLH1PjHaZd3SZztCxv06p5t67Qw=
---



---
AIGC:
    Label: "1"
    ContentProducer: 001191440300708461136T1XGW3
    ProduceID: 89f2995f43cde2f2717e466bdd92de27_6bf7a85bd37e11f18888525400cd780f
    ReservedCode1: C7XeONwtr7H6bkeIaJWUpqDLJvDub/WCdCwoN5yF6Dw5GgS2zyswk3PCGpxfWjxRk4dkQtG6ejTocmtMpVQYnWcZe1+5gOfIRVxDPcet94G9WxDh0JnR0m2v81bXGgTpRGHbk9IxGVqBNUA/7wzezGmlHGTpQ8n4y56mEmxSAtPy7jCKLbtb+GZ1SBA=
    ContentPropagator: 001191440300708461136T1XGW3
    PropagateID: 89f2995f43cde2f2717e466bdd92de27_6bf7a85bd37e11f18888525400cd780f
    ReservedCode2: C7XeONwtr7H6bkeIaJWUpqDLJvDub/WCdCwoN5yF6Dw5GgS2zyswk3PCGpxfWjxRk4dkQtG6ejTocmtMpVQYnWcZe1+5gOfIRVxDPcet94G9WxDh0JnR0m2v81bXGgTpRGHbk9IxGVqBNUA/7wzezGmlHGTpQ8n4y56mEmxSAtPy7jCKLbtb+GZ1SBA=
---





# 装修报价管理工作台（renovation-quote-workbench）

浏览器网页版装修报价管理工具：多项目管理、可编辑报价表、Excel 数据导入、导出 Excel 报价单（布局对齐用户提供的「金地时代艺境9#2-902」装修报价单 Excel，含家装明细 / 定制明细 / 水电明细三张表）。

## 使用方式

直接用浏览器打开 `workbench/index.html`（无需安装依赖，离线可用）。

## 目录结构

```
renovation-quote-workbench/
├── SKILL.md                        # 本说明
└── workbench/
    ├── index.html                  # 工作台主页面（内联全部业务逻辑）
    ├── libs/xlsx.full.min.js       # SheetJS 本地库（离线 Excel 读写）
    ├── templates/装修报价模板.xlsx  # 空白模板（5 sheet，由 build_template.py 生成，与工作台导出结构一致）
    └── scripts/build_template.py   # 模板生成脚本（Python/openpyxl，可重新生成模板）
```

## 模板结构（对齐用户新模板：5 个 sheet）

1. **项目信息**：A1:B8 键值对 —— 项目名称 / 客户 / 地址 / 报价日期 / 备注 / 材料损耗率(%) / 管理费率(%) / 税率(%)。
2. **家装报价明细**（多级结构：大分区 → 区域 → 明细）：
   - 页眉公司信息（B2 公司名称 / B4 公司地址 / B5 工程地点，可编辑）；
   - 大分区标题行：如 `基础装修项目（轻工辅料）`、`隐蔽及基础项目`，浅蓝底色，可折叠（网页端）；
   - 区域标题行：`房高（米）/ 值 | 面积（平米）/ 值 | 周长（米）/ 值`（H/I/J/K/L 列），尺寸值支持 Excel 公式引用（如 `=H8`、`=10.7+8.9`）；
   - 表头 6 列：项目名称 / 工程量 / 单位 / 单价（元）/ 合计（元）/ 施工说明（施工位置、辅料品牌、尺寸、工艺标准）；
   - 明细行：工程量支持语义公式（`=面积`、`=周长*房高+面积`、跨区域引用 `=卫生间.刷背胶.数量`），也允许手工输入数字；合计 = 工程量 × 单价（自动）；
   - 小计行：B:F 合并显示"小计（元）"，G 列备注；
   - 分区合计 / 工程总价 / 最终总价行（含定制）。
3. **定制明细**（全屋定制 + 橱柜定制双栏）：
   - 客户姓名 / 客户电话 / 品牌及类型 / 板材级别 / 客户地址 / 签订日期 / 安装日期 / 负责人；
   - 左栏全屋定制（序号 / 项目名称 / 数量 / 单位 / 单价 / 小计），右栏橱柜定制（同构）；
   - 全屋定制总价 / 橱柜定制总价 / 定制及橱柜合计总价 / 最终总价；
   - 客户确认签字区、备注区。
4. **水电明细**：现场水电项目工程量明细（序号 / 工程项目名称 / 单位 / 工程量 / 单价 / 小计 / 备注，小计 = 单价 × 工程量）、单项报价参考、服务条款 12 条。
5. **报价汇总**：大分区 / 区域 / 面积(m²) / 周长(m) / 房高(m) / 合计 / 占比 / 行数 + 工程总价行。

## 功能说明

- 多项目管理：新建 / 重命名 / 删除项目，顶部下拉切换；每个项目独立报价单，数据存于浏览器 localStorage。
- 公司页眉：报价单顶部显示并可编辑 公司名称 / 公司地址 / 工程地点 / 联系电话（Q6）。
- 大分区（Q3）：区域之上按分组管理（如 基础装修项目（轻工辅料）、隐蔽及基础项目、主材类目），支持折叠展开；区域支持增删 / 排序 / 跨分区移动。
- 区域尺寸（Q4）：统一为 房高（米）/ 面积（平米）/ 周长（米），值支持输入或公式引用（=J8、=L8*H8+J8、=14+12.8 等），引用自动随来源更新。
- 明细 6 列（Q1）：项目名称 | 工程量 | 单位 | 单价（元）| 合计（元）| 施工说明；工程量支持语义公式（=面积 / =周长 / =周长*房高+面积 / 跨区域 `=分区.区域.项目.数量`），可手工覆盖。
- 定制明细 / 水电明细（Q2）：可切换标签页查看与编辑；定制明细按全屋 + 橱柜双栏展示，水电明细含现场工程量 + 单项报价参考 + 服务条款。
- 区域汇总：按大分区 / 区域自动汇总金额、占比、行数；底部展示损耗 / 管理费 / 税金配置项（可设为 0 不启用）。
- Excel 导入：兼容新 6 列格式与旧 11 列格式（序号/项目名称/材料名称/规格/单位/数量/单价/金额/人工费/合计/备注），自动识别 5 sheet / 3 sheet / 旧扁平结构；旧数据（平铺 rows → regions → partitions、旧尺寸字段 → 房高/面积/周长、材料/规格/人工费 → 施工说明）自动迁移。
- 导出 Excel：当前项目导出为 5 sheet xlsx（项目信息 / 家装报价明细 / 定制明细 / 水电明细 / 报价汇总），带合并单元格与底色；工程量公式转 Excel 单元格引用（=J7、=L8*H8+J8 等）以文本保存，SheetJS 导入可识别。
- 下载空白模板：一键下载与导出结构一致的空白模板（可编辑后导入）。
- JSON 备份 / 恢复：项目数据一键导出 / 导入 JSON，便于存档与迁移。

## 关键实现

- **公式引擎**（无 DOM 可测）：`QuoteCore.evalFormula` 支持纯算术 / 引用当前区域尺寸（面积 / 周长 / 房高）/ 行内数量 / 跨区域引用（`.面积`、`.项目.数量`、`.项目.合计`）。
- **SheetJS 无 DOM 要点**：本地 `xlsx.full.min.js` 为浏览器构建（禁用 Node fs，`readFile/writeFile` 不可用），须用 `XLSX.read(fs.readFileSync(p), {type:"buffer", cellFormula:true})` 读、`fs.writeFileSync(p, Buffer.from(XLSX.write(wb, {bookType:"xlsx", type:"array"})))` 写。
- **模板公式兼容**：openpyxl 将 `=J8` 存为公式（无缓存）时 SheetJS 读不到值，因此 `build_template.py` 对以 `=` 开头的字符串强制 `cell.data_type="s"`（文本保存），保证浏览器与 Node 端 SheetJS 均可识别。

## 扩展端口

- 明细表头别名：`normalizeHeader` 的 `HEADER_ALIAS` 中追加。
- 区域标题行解析支持：新 6 列分段形态（房高/面积/周长）与旧 11 列分段形态。
- 模板结构变更：改 `build_template.py` 后重新运行生成模板。

## 自测记录（v4 整合版，Node 无 DOM 全链路）

- 无 DOM 自测脚本：`temp/v4_selftest.js`（Node + 本地 `workbench/libs/xlsx.full.min.js` + 核心逻辑 `temp/v4_core_extracted.js`，提取自 index.html 内联脚本）。
- 导入真实附件《金地时代艺境9#2-902.xlsx》：识别 3 sheet；大分区 4 个（基础装修项目（轻工辅料）/ 隐蔽及基础项目×2 / 主材类目）；区域尺寸 2.7/31.2/39 与公式引用（=10.7+8.9、=14+12.8）；工程量语义化（=面积 / =周长*房高+面积 / 跨区域 瓷砖美缝 = 卫生间.刷背胶.数量+卫生间.面积+厨房.刷背胶.数量+厨房.面积+主卧、次卧.面积+客餐厅、玄关、阳台、过道.面积 = 105.44）；客餐厅小计 13045.5、主卧次卧 7783.2、厨房 2750.4、卫生间 3485.93、基础装修合计 27065.03、全项目明细 76187.6。
- 定制明细：全屋 4 行 13579.2、橱柜 3 行 3508；水电现场 6 行、单项参考 15 行、服务条款 ≥12 条。
- 导出回读：5 sheet 齐全、6 列表头、公司页眉、合计/小计/分区合计/工程总价公式、工程量 Excel 引用（=J7）、定制小计公式、水电 =E*D 公式。
- 空白模板：5 sheet、房高/面积/周长标签、工程量 =J 引用、解析后明细合计 14443.1。
- 旧 11 列格式兼容：区域尺寸迁移、材料/规格/人工费并入施工说明、合计 1950。
- 旧平铺数据迁移：regions → partitions、缺省 custom/water/company 补全。
- 手工公式引擎：纯尺寸、行内数量×单价、跨区域 .面积 / .项目.合计 引用。
- 当前状态：`node temp/v4_selftest.js` **76 PASS / 0 FAIL**（exit 0）；`node --check` 校验 core / ui 片段通过。
*（内容由AI生成，仅供参考）*
*（内容由AI生成，仅供参考）*
*（内容由AI生成，仅供参考）*
