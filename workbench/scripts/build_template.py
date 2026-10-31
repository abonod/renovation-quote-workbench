#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_template.py
根据「金地时代艺境9#2-902.xlsx」附件三 sheet 结构 + 用户确认的 Q1-Q6 方案，
生成工作台配套的空白模板：workbench/templates/装修报价模板.xlsx。

模板结构（5 个 sheet，与 index.html「下载空白模板」产物一致，可被工作台直接导入）：
  1) 项目信息   ：公司信息 + 项目名称/客户/地址/报价日期/备注/损耗率/管理费率/税率
  2) 家装报价明细：页眉公司信息 + 大分区（B:L 合并行，支持折叠）+
                   区域标题行（B 列区域名 | G 房高（米）H 值 | I 面积（平米）J 值 | K 周长（米）L 值）+
                   6 列表头（B 项目名称 | C 工程量 | D 单位 | E 单价（元） | F 合计（元） |
                   G 施工位置、辅料品牌、尺寸、工艺标准）+
                   明细行（C 工程量支持 =J8 房高/面积/周长引用公式，F =C*E 合计公式）+
                   小计行（=SUM 公式）+ 分区合计行 + 工程总价行 + 最终总价行（含定制）
  3) 定制明细   ：客户信息 + 全屋定制/橱柜定制双栏 + 总价行 + 客户确认签字区 + 备注
  4) 水电明细   ：现场水电项目工程量明细（=E*D 公式）+ 合计 + 单项报价参考 + 服务条款
  5) 报价汇总   ：大分区/区域数/行数/合计/占比 + 全项目总计

运行：python build_template.py
输出：workbench/templates/装修报价模板.xlsx
"""
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.normpath(os.path.join(HERE, "..", "templates"))
OUT_PATH = os.path.join(OUT_DIR, "装修报价模板.xlsx")

PARTITION_COLOR = "DDEBF7"   # 大分区行浅蓝
REGION_COLOR = "E2EFDA"      # 区域标题行浅绿
HEADER_COLOR = "D9E2F3"      # 表头浅蓝
SUBTOTAL_COLOR = "FCE4D6"    # 小计/总计行浅橙

thin = Side(style="thin", color="8EAADB")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def style_cell(cell, bold=False, fill=None, align=CENTER):
    cell.border = BORDER
    cell.alignment = align
    cell.font = Font(name="微软雅黑", size=10, bold=bold)
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)


def row_values(ws, r, vals, fill=None, bold=False, aligns=None):
    for i, v in enumerate(vals, start=1):
        cell = ws.cell(row=r, column=i, value=v)
        # 以 = 开头的公式字符串以文本形式保存：避免 openpyxl 写成无缓存值的公式，
        # 否则浏览器 SheetJS 导入时公式单元格读不到，工作台无法识别引用
        if isinstance(v, str) and v.startswith("="):
            cell.data_type = "s"
        a = CENTER
        if aligns and i - 1 < len(aligns) and aligns[i - 1]:
            a = aligns[i - 1]
        style_cell(cell, bold=bold, fill=fill, align=a)


def merge(ws, r, c1, c2):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)


def build():
    os.makedirs(OUT_DIR, exist_ok=True)
    wb = Workbook()

    # ---------- 1. 项目信息 ----------
    ws_info = wb.active
    ws_info.title = "项目信息"
    info_rows = [
        ("项目名称", "示例项目"),
        ("客户", ""),
        ("地址", ""),
        ("报价日期", ""),
        ("备注", ""),
        ("公司名称", "xx"),
        ("公司地址", "xx"),
        ("工程地点", "xx"),
        ("联系电话", ""),
        ("材料损耗率(%)", 0),
        ("管理费率(%)", 0),
        ("税率(%)", 0),
    ]
    for i, (k, v) in enumerate(info_rows, start=1):
        kc = ws_info.cell(row=i, column=1, value=k)
        vc = ws_info.cell(row=i, column=2, value=v)
        style_cell(kc, bold=True, fill=HEADER_COLOR)
        style_cell(vc, align=LEFT)
    # A14 AI 备注行（与用户手动模板一致）
    ai_note = ws_info.cell(row=14, column=1, value="（内容由AI生成，仅供参考）")
    style_cell(ai_note, align=LEFT)
    ws_info.column_dimensions["A"].width = 16
    ws_info.column_dimensions["B"].width = 44

    # ---------- 2. 家装报价明细 ----------
    ws = wb.create_sheet("家装报价明细")
    for c, w in enumerate([6, 30, 12, 8, 10, 12, 60, 8, 12, 10, 10, 10], start=1):
        ws.column_dimensions[get_column_letter(c)].width = w

    r = 1
    # 页眉（公司信息）
    row_values(ws, r, [""] * 12); r += 1
    row_values(ws, r, ["", "xx"] + [""] * 10); merge(ws, r, 2, 12); ws.row_dimensions[r].height = 20; r += 1
    row_values(ws, r, ["", "公司地址：xx"] + [""] * 10); merge(ws, r, 2, 12); r += 1
    row_values(ws, r, ["", "工程地点：xx"] + [""] * 10); merge(ws, r, 2, 12); r += 1
    row_values(ws, r, ["", "联系电话："] + [""] * 10); merge(ws, r, 2, 12); r += 1
    row_values(ws, r, [""] * 12); r += 1

    # 大分区
    row_values(ws, r, ["", "基础装修项目（轻工辅料）"] + [""] * 10, fill=PARTITION_COLOR, bold=True)
    merge(ws, r, 2, 12); ws.row_dimensions[r].height = 20; r += 1

    # 区域1：客餐厅、玄关、阳台、过道
    row_values(ws, r, ["", "客餐厅、玄关、阳台、过道", "", "", "", "", "房高（米）", 2.7, "面积（平米）", 31.2, "周长（米）", 39], fill=REGION_COLOR)
    merge(ws, r, 2, 3); ws.row_dimensions[r].height = 20
    region1_title = r; r += 1
    row_values(ws, r, ["", "项目名称", "工程量", "单位", "单价（元）", "合计（元）", "施工位置、辅料品牌、尺寸、工艺标准"], fill=HEADER_COLOR, bold=True)
    header1 = r; r += 1
    row_values(ws, r, ["", "地面铺砖", "=J{}".format(region1_title), "平米", 50, "=C{}*E{}".format(r, r),
                       "1、800*800常规瓷砖正常铺贴，门窗口不减面积；（注：地面拼花、造型费用另计）"],
               aligns=[None, LEFT, CENTER, CENTER, CENTER, CENTER, LEFT])
    ws.row_dimensions[r].height = 18; r += 1
    row_values(ws, r, ["", "墙顶基层处理及亮白辅料", "=L{}*H{}+J{}".format(region1_title, region1_title, region1_title), "平米", 30, "=C{}*E{}".format(r, r),
                       "1.采用（美巢）墙锢所有墙顶面滚刷一遍；2、批刮（美巢）腻子"],
               aligns=[None, LEFT, CENTER, CENTER, CENTER, CENTER, LEFT])
    ws.row_dimensions[r].height = 18
    detail_end1 = r; r += 1
    row_values(ws, r, ["", "小计（元）", "", "", "", "=SUM(F{}:F{})".format(header1 + 1, detail_end1), "备注"], fill=SUBTOTAL_COLOR, bold=True)
    merge(ws, r, 2, 5)
    subtotal1 = r; r += 1

    # 区域2：厨房
    row_values(ws, r, ["", "厨房", "", "", "", "", "房高（米）", "=H{}".format(region1_title), "面积（平米）", 3.93, "周长（米）", 9], fill=REGION_COLOR)
    merge(ws, r, 2, 3); ws.row_dimensions[r].height = 20
    region2_title = r; r += 1
    row_values(ws, r, ["", "项目名称", "工程量", "单位", "单价（元）", "合计（元）", "施工位置、辅料品牌、尺寸、工艺标准"], fill=HEADER_COLOR, bold=True)
    header2 = r; r += 1
    row_values(ws, r, ["", "地面铺砖", "=L{}*H{}+J{}".format(region2_title, region2_title, region2_title), "平米", 50, "=C{}*E{}".format(r, r),
                       "1、800*800、300*600常规瓷砖正常铺贴，门窗口不减面积"],
               aligns=[None, LEFT, CENTER, CENTER, CENTER, CENTER, LEFT])
    ws.row_dimensions[r].height = 18; r += 1
    row_values(ws, r, ["", "防水", 1, "项", 600, "=C{}*E{}".format(r, r),
                       "1、滚刷（美巢或者东方雨虹防水固）防水涂料2遍，高度0.4米；2、闭水测试24小时"],
               aligns=[None, LEFT, CENTER, CENTER, CENTER, CENTER, LEFT])
    ws.row_dimensions[r].height = 18
    detail_end2 = r; r += 1
    row_values(ws, r, ["", "小计（元）", "", "", "", "=SUM(F{}:F{})".format(header2 + 1, detail_end2), "备注"], fill=SUBTOTAL_COLOR, bold=True)
    merge(ws, r, 2, 5)
    subtotal2 = r; r += 1

    # 分区合计 + 工程总价 + 最终总价
    row_values(ws, r, ["", "基础装修项目（轻工辅料）合计（元）", "", "", "", "=F{}+F{}".format(subtotal1, subtotal2)], fill=SUBTOTAL_COLOR, bold=True)
    merge(ws, r, 2, 5)
    part_sum = r; r += 1
    row_values(ws, r, ["", "工程总价（元）", "", "", "", "=F{}".format(part_sum)], fill=SUBTOTAL_COLOR, bold=True)
    merge(ws, r, 2, 5); r += 1
    row_values(ws, r, ["", "最终总价（元）（含定制）", "", "", "", 0], fill=SUBTOTAL_COLOR, bold=True)
    merge(ws, r, 2, 5)
    ws.freeze_panes = "A2"

    # ---------- 3. 定制明细 ----------
    ws_c = wb.create_sheet("定制明细")
    cvals = [
        ["全屋定制报价明细"],
        ["客户姓名", "", "客户电话", "", "品牌及类型", "", "板材级别", ""],
        ["客户地址", "", "", "", "签订日期", "", "安装日期", ""],
        ["负责人", "", "电话", ""],
        ["品类明细", "", "全屋定制", "", "五金", "", "", "", "橱柜", "台面"],
        ["", "", "门板", "", "滑轨", "", "玻璃门", "", "", "门板"],
        ["", "", "层板", "", "拉手", "", "水槽", "", "", "柜体"],
        ["", "", "背板", "", "拉直器", "", "灯线", "", "", "背板"],
        ["全屋定制部分", "", "", "", "", "", "", "", "橱柜定制部分"],
        ["项目", "序号", "项目名称", "数量", "", "单位", "单价（元）", "小计（元）", "", "序号", "项目名称", "", "数量", "单位", "单价（元）", "小计（元）"],
        ["", 1, "电视柜", 3.06, "", "㎡", 720, "=G11*D11", "", 1, "橱柜（地柜）", "", 2.86, "延米", 800, "=O11*M11"],
        ["", 2, "主卧衣柜", 4.9, "", "㎡", 720, "=G12*D12", "", 2, "橱柜（吊柜）", "", 1.45, "延米", 600, "=O12*M12"],
        ["", 3, "次卧衣柜", 6.2, "", "㎡", 720, "=G13*D13", "", 3, "包管", "", 1, "项", 350, "=O13*M13"],
        ["全屋定制总价（元）", "", "", "=SUM(H11:H13)", "", "", "", "橱柜定制总价（元）", "", "", "=SUM(P11:P13)", "", "", "", ""],
        ["定制及橱柜合计总价（元）", "", "", "=D27+K27", "", "", "", "", "", "最终总价（元）", "", "", "", "", ""],
        ["客户确认签字："],
        ["备注："],
    ]
    for i, row in enumerate(cvals, start=1):
        for j, v in enumerate(row, start=1):
            cell = ws_c.cell(row=i, column=j, value=v)
            if isinstance(v, str) and v.startswith("="):
                cell.data_type = "s"  # 公式文本保存，兼容工作台 SheetJS 导入
            style_cell(cell, bold=(i == 1 or i == 9 or i in (14, 15)), fill=HEADER_COLOR if i == 1 or i == 9 else (SUBTOTAL_COLOR if i in (14, 15) else None), align=LEFT if j in (3, 10) else CENTER)
    merge(ws_c, 1, 1, 15)
    merge(ws_c, 9, 1, 8); merge(ws_c, 9, 9, 15)
    merge(ws_c, 2, 5, 6); merge(ws_c, 2, 7, 8)
    merge(ws_c, 3, 3, 4); merge(ws_c, 3, 5, 6)
    merge(ws_c, 4, 3, 4)
    merge(ws_c, 10, 4, 5); merge(ws_c, 10, 8, 9)
    merge(ws_c, 14, 7, 9)
    for c, w in enumerate([16, 6, 16, 8, 8, 10, 10, 4, 6, 16, 8, 8, 10, 10, 8], start=1):
        ws_c.column_dimensions[get_column_letter(c)].width = w

    # ---------- 4. 水电明细 ----------
    ws_w = wb.create_sheet("水电明细")
    wvals = [
        ["水电价格表（合同附件）"],
        ["现场水电项目工程量明细"],
        ["序号", "工程项目名称", "工程量", "单位", "单价", "小计", "备注"],
        [1, "电路改造  2.5²", 0, "米", 50, "=E4*D4", ""],
        [2, "灯线改造", 0, "米", 45, "=E5*D5", ""],
        [3, "水路改造", 0, "米", 55, "=E6*D6", ""],
        ["合计", "", "", "", "", "=SUM(F4:F6)", ""],
        ["单项报价参考"],
        ["序号", "项目名称", "工程量", "单位", "单价", "备注"],
        [1, "水路改造", 0, "米", 55, ""],
        [2, "灯线改造", 0, "米", 45, ""],
        [3, "墙排工艺", 0, "项", 200, ""],
        ["服务条款"],
        ["1、为了维护您的利益，请你不要接受任何口头承诺。"],
        ["2、实际发生项目若与报价不符，一切以实际项目发生为准。"],
        ["3、以上报价不含物业收取的装修管理费。管理费费用务必客户自理。"],
    ]
    for i, row in enumerate(wvals, start=1):
        for j, v in enumerate(row, start=1):
            cell = ws_w.cell(row=i, column=j, value=v)
            if isinstance(v, str) and v.startswith("="):
                cell.data_type = "s"  # 公式文本保存，兼容工作台 SheetJS 导入
            style_cell(cell, bold=(i in (1, 2, 3, 8, 9, 13)), fill=HEADER_COLOR if i in (3, 9) else None, align=LEFT if i >= 13 else CENTER)
    merge(ws_w, 1, 1, 7); merge(ws_w, 2, 1, 7)
    merge(ws_w, 7, 1, 3)
    merge(ws_w, 8, 1, 7); merge(ws_w, 13, 1, 7)
    for i in range(14, 17):
        merge(ws_w, i, 1, 7)
    for c, w in enumerate([6, 26, 8, 10, 8, 10, 50], start=1):
        ws_w.column_dimensions[get_column_letter(c)].width = w

    # ---------- 5. 报价汇总 ----------
    ws_sum = wb.create_sheet("报价汇总")
    sum_headers = ["大分区", "区域数", "行数", "合计(元)", "占比"]
    for c, h in enumerate(sum_headers, start=1):
        cell = ws_sum.cell(row=1, column=c, value=h)
        style_cell(cell, bold=True, fill=HEADER_COLOR)
    sum_rows = [
        ("基础装修项目（轻工辅料）", 2, 4, 0, ""),
        ("全项目总计", "", "", 0, "100%"),
    ]
    for i, row in enumerate(sum_rows, start=2):
        is_total = row[0] == "全项目总计"
        for c, v in enumerate(row, start=1):
            cell = ws_sum.cell(row=i, column=c, value=v)
            style_cell(cell, bold=is_total, fill=SUBTOTAL_COLOR if is_total else None)
    for c, w in enumerate([26, 10, 10, 12, 10], start=1):
        ws_sum.column_dimensions[get_column_letter(c)].width = w

    wb.save(OUT_PATH)
    print("OK ->", OUT_PATH)
    print("sheets:", wb.sheetnames)


if __name__ == "__main__":
    build()
