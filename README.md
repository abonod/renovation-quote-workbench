# 装修报价管理工作台（Renovation Quote Workbench）

面向装修报价 / 预算人员的多项目报价管理网页工具。单文件 HTML 实现，无需安装依赖，浏览器打开即用；支持 Excel 模板导入导出、工程量语义公式、多分区明细管理与主题切换。

## 功能特性

- **多项目报价管理**：项目卡片化管理，支持新建、归档、回收站批量删除
- **三大报价分区**：家装报价明细、定制明细（全屋定制 / 橱柜定制）、水电明细，区域标题行分组汇总
- **工程量语义公式**：输入 `=面积`、`=房高*周长+面积` 等语义公式自动计算工程量，也可直接选预置工程量/单位下拉
- **自动汇总**：金额 = 数量 × 单价；区域汇总、分区合计、全屋总计逐级联动
- **Excel 模板导入导出**：导入官方模板（项目信息 / 家装报价明细 / 定制明细 / 水电明细 / 报价汇总 5 个 Sheet），导出带样式的报价单（ExcelJS）
- **公司信息从模板识别**：公司名称、地址、工程地点均取自用户导入的模板，不硬编码第三方公司名
- **6 套装修行业主题**：暖金经典、北欧原木、深蓝科技等，选择持久化到本地

## 安装与使用

### 方式一：作为 Skill 安装（推荐）

下载 `renovation-quote-workbench-skill-v1.zip`，在支持 Skill 的 Agent 能力市场中通过「导入本地 ZIP」安装。安装后可通过 Agent 直接调用报价工作台能力。

### 方式二：网页版直接使用

克隆仓库后用浏览器打开 `workbench/index.html`，或将 `workbench` 目录部署到任意静态服务器 / 本地 HTTP 服务。

### 快速上手

1. 新建报价单，导入 `workbench/templates/装修报价模板.xlsx`（或使用内置模板）
2. 在各分区填写项目，工程量列支持语义公式（如 `=面积`）
3. 单价 / 数量自动计算金额，汇总逐级联动
4. 导出 Excel 报价单，含样式与纯数字结果

## 目录结构

```
renovation-quote-workbench/
├── SKILL.md                          # Skill 入口（含 YAML frontmatter）
├── renovation-quote-workbench-skill-v1.zip   # 标准 Skill 安装包
├── workbench/                        # 网页版源码（单文件应用）
│   ├── index.html                    # 全部逻辑内联
│   ├── libs/                         # SheetJS / ExcelJS 运行库
│   ├── scripts/build_template.py     # 模板构建脚本
│   └── templates/装修报价模板.xlsx    # 官方模板（5 Sheet）
├── demo/
│   ├── renovation-quote-workbench-demo-portrait.mp4     # 竖屏功能演示视频
│   └── renovation-quote-workbench-demo-portrait-cover.png
└── docs/
    └── 发布安装指引.md
```

## 演示

竖屏演示视频（1080×1920 / 84.45s）：见 `demo/renovation-quote-workbench-demo-portrait.mp4`。

## 说明

- 工具导出数据为纯数字结果（不含公式字符串），便于交接与审计
- 计算结果仅供参考，正式报价请以人工复核为准
