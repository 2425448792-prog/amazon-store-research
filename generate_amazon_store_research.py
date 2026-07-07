"""生成《亚马逊品牌旗舰店竞品调研.xlsx》。

优先使用 openpyxl；若当前离线环境未安装 openpyxl，则使用标准库生成等价 xlsx，
保证脚本可重复运行并完成基础验证。
"""
from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape

OUTPUT = Path("亚马逊品牌旗舰店竞品调研.xlsx")
FONT = "微软雅黑"
HEADER_COLOR = "1F4E78"
CONCLUSION_PRACTICES_START_ROW = 26
CONCLUSION_PRACTICES_END_ROW = 45

SHEETS = {
    "02 竞品清单": ["编号","品牌名称","亚马逊站点","店铺类型","主营品类","品牌旗舰店链接","选择原因","重点参考内容","调研优先级","调研状态","备注"],
    "03 店铺调研": ["编号","品牌名称","店铺类型","亚马逊站点","主营品类","品牌旗舰店链接","首页参考图","品牌定位","目标消费者","首页模块结构","导航与产品分类方式","首屏设计特点","整体视觉风格","色彩与字体特点","图片使用方式","品牌故事表达","产品卖点表达","产品展示方式","场景化展示","页面跳转与转化路径","优点","不足","可借鉴点","不建议照搬的内容","首屏吸引力评分","导航分类评分","视觉统一性评分","卖点表达评分","品牌表达评分","转化路径评分","综合评分","借鉴优先级","调研状态","调研日期","备注"],
    "04 优秀案例": ["案例编号","品牌名称","模块类型","参考截图","店铺链接","所在页面","案例说明","优秀之处","可以借鉴什么","适用我们的哪个页面","借鉴优先级","是否建议采用","备注"],
    "05 调研结论": [],
    "06 装修建议": ["建议编号","装修模块","当前问题或机会点","装修建议","参考品牌","参考案例编号","参考截图","店铺链接","预期作用","实施难度","建议优先级","是否采纳","上级意见","备注"],
}

DROPDOWNS = {
    ("02 竞品清单", "D"): ["直接竞品","同品类优秀店铺","跨品类参考店铺"],
    ("02 竞品清单", "I"): ["高","中","低"],
    ("02 竞品清单", "J"): ["未开始","进行中","已完成","暂不调研"],
    ("03 店铺调研", "Y:AD"): ["1","2","3","4","5"],
    ("03 店铺调研", "AF"): ["P0-重点参考","P1-建议参考","P2-一般参考"],
    ("03 店铺调研", "AG"): ["未开始","进行中","已完成"],
    ("04 优秀案例", "C"): ["首页首屏","导航栏","产品分类","主推产品","使用场景","品牌故事","功能卖点","产品对比","人群分类","活动模块","视频模块","其他"],
    ("04 优秀案例", "K"): ["P0","P1","P2"],
    ("04 优秀案例", "L"): ["建议采用","调整后采用","暂不采用","待上级确认"],
    ("05 调研结论", "F"): ["高","中","低"],
    ("05 调研结论", "G"): ["建议采用","调整后采用","暂不采用","待上级确认"],
    ("06 装修建议", "B"): ["首页首屏","导航栏","产品分类","核心产品","品牌故事","使用场景","功能卖点","产品对比","图片风格","视频模块","其他"],
    ("06 装修建议", "J"): ["低","中","高"],
    ("06 装修建议", "K"): ["P0-必须优化","P1-建议优化","P2-后续优化"],
    ("06 装修建议", "L"): ["采纳","部分采纳","暂不采纳","待确认"],
}


def build_with_openpyxl() -> None:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.cell import range_boundaries

    wb = Workbook()
    wb.remove(wb.active)
    header_fill = PatternFill("solid", fgColor=HEADER_COLOR)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style_table(ws, rows, cols, freeze="A2"):
        ws.freeze_panes = freeze
        ws.auto_filter.ref = f"A1:{get_column_letter(cols)}{rows}"
        for row in ws.iter_rows(min_row=1, max_row=rows, max_col=cols):
            for c in row:
                c.font = Font(name=FONT, size=10, bold=c.row == 1, color="FFFFFF" if c.row == 1 else "000000")
                c.fill = header_fill if c.row == 1 else PatternFill(fill_type=None)
                c.alignment = Alignment(horizontal="center" if c.row == 1 or c.column <= 6 else "left", vertical="center", wrap_text=c.column > 6)
                c.border = border
        ws.row_dimensions[1].height = 24

    def add_dv(ws, col_expr, values, start_row, end_row):
        if end_row < start_row:
            raise ValueError(
                f"Invalid data validation range: start_row={start_row}, end_row={end_row}"
            )
        dv = DataValidation(type="list", formula1='"' + ','.join(values) + '"', allow_blank=True)
        ws.add_data_validation(dv)
        if ":" not in col_expr:
            dv.add(f"{col_expr}{start_row}:{col_expr}{end_row}")
            return
        start_col, end_col = col_expr.split(":", 1)
        dv.add(f"{start_col}{start_row}:{end_col}{end_row}")

    for title, headers in SHEETS.items():
        ws = wb.create_sheet(title)
        if title == "05 调研结论":
            ws.append(["指标", "数值"])
            metrics = [
                ("计划调研店铺数", 20),("已录入店铺数", '=COUNTA(\'03 店铺调研\'!B2:B21)'),("已完成调研数", '=COUNTIF(\'03 店铺调研\'!AG2:AG21,"已完成")'),
                ("直接竞品数量", '=COUNTIF(\'03 店铺调研\'!C2:C21,"直接竞品")'),("同品类优秀店铺数量", '=COUNTIF(\'03 店铺调研\'!C2:C21,"同品类优秀店铺")'),("跨品类参考店铺数量", '=COUNTIF(\'03 店铺调研\'!C2:C21,"跨品类参考店铺")'),
                ("平均综合评分", '=IFERROR(ROUND(AVERAGE(\'03 店铺调研\'!AE2:AE21),1),"")'),("最高评分店铺", '=IFERROR(INDEX(\'03 店铺调研\'!B2:B21,MATCH(MAX(\'03 店铺调研\'!AE2:AE21),\'03 店铺调研\'!AE2:AE21,0)),"暂无数据")'),("已收集优秀案例数量", '=COUNTA(\'04 优秀案例\'!A2:A41)')]
            for m in metrics: ws.append(list(m))
            ws.append([]); ws.append(["分析维度","竞品普遍做法","优秀店铺做法","常见问题","对我们的启示"])
            for d in ["首页首屏","页面结构","导航设计","产品分类","视觉风格","品牌故事","场景图片","产品卖点","产品展示","页面跳转","转化路径"]: ws.append([d])
            ws.append([]); ws.append(["编号","优秀做法","参考品牌","参考案例编号","适用页面","重要程度","是否建议采用"])
            for _ in range(CONCLUSION_PRACTICES_END_ROW - CONCLUSION_PRACTICES_START_ROW + 1): ws.append([])
            rows, cols = ws.max_row, 7
            assert rows == CONCLUSION_PRACTICES_END_ROW
            ws.freeze_panes="A2"; ws.auto_filter.ref=f"A13:E23"
        else:
            ws.append(headers); rows = {"02 竞品清单":21,"03 店铺调研":21,"04 优秀案例":41,"06 装修建议":31}[title]; cols=len(headers)
            for r in range(2, rows+1):
                vals = [r-1] + [""]*(cols-1)
                ws.append(vals)
                if title == "03 店铺调研": ws[f"AE{r}"] = f'=IF(COUNT(Y{r}:AD{r})=0,"",ROUND(AVERAGE(Y{r}:AD{r}),1))'
            style_table(ws, rows, cols, "C2" if title == "03 店铺调研" else "A2")
        if title == "05 调研结论": style_table(ws, rows, cols, "A2")
        for c in range(1, cols+1): ws.column_dimensions[get_column_letter(c)].width = 18
        for wide in ["D","G"] if title in ["04 优秀案例","06 装修建议"] else ["G"] if title=="03 店铺调研" else []: ws.column_dimensions[wide].width=28
        for key, vals in DROPDOWNS.items():
            if key[0] == title:
                start_row = 26 if title == "05 调研结论" else 2
                end_row = 45 if title == "05 调研结论" else rows
                add_dv(ws, key[1], vals, start_row, end_row)
        if title == "03 店铺调研": ws.conditional_formatting.add("AE2:AE21", ColorScaleRule(start_type='num', start_value=1, start_color='F8696B', mid_type='num', mid_value=3, mid_color='FFEB84', end_type='num', end_value=5, end_color='63BE7B'))
    wb.save(OUTPUT)

    assert OUTPUT.exists() and OUTPUT.stat().st_size > 0
    reloaded = load_workbook(OUTPUT)
    try:
        assert reloaded.sheetnames == list(SHEETS)
        conclusion_dv_ranges = set()
        for ws in reloaded.worksheets:
            for dv in ws.data_validations.dataValidation:
                for cell_range in dv.cells.ranges:
                    min_col, min_row, max_col, max_row = range_boundaries(str(cell_range))
                    if min_row > max_row:
                        raise ValueError(
                            f"Invalid data validation range: start_row={min_row}, end_row={max_row}"
                        )
                    if ws.title == "05 调研结论":
                        conclusion_dv_ranges.add(str(cell_range))
        expected_conclusion_ranges = {
            f"F{CONCLUSION_PRACTICES_START_ROW}:F{CONCLUSION_PRACTICES_END_ROW}",
            f"G{CONCLUSION_PRACTICES_START_ROW}:G{CONCLUSION_PRACTICES_END_ROW}",
        }
        missing_ranges = expected_conclusion_ranges - conclusion_dv_ranges
        if missing_ranges:
            raise ValueError(
                "Missing conclusion data validation ranges: "
                + ", ".join(sorted(missing_ranges))
            )
    finally:
        reloaded.close()


def fallback_minimal_xlsx() -> None:
    """离线兜底：生成基础 xlsx，保留工作表、表头、空行和公式。"""
    # 使用最小 OpenXML 结构，避免在无 openpyxl 的受限环境中无法交付文件。
    sheets = list(SHEETS)
    content_types = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' + ''.join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1,6)) + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
    rels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    wbxml = '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + ''.join(f'<sheet name="{escape(n)}" sheetId="{i}" r:id="rId{i}"/>' for i,n in enumerate(sheets,1)) + '</sheets></workbook>'
    wbrels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ''.join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1,6)) + '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    styles = '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="10"/><name val="微软雅黑"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="微软雅黑"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/></patternFill></fill></fills><borders count="2"><border/><border><left style="thin"><color rgb="FFD9D9D9"/></left><right style="thin"><color rgb="FFD9D9D9"/></right><top style="thin"><color rgb="FFD9D9D9"/></top><bottom style="thin"><color rgb="FFD9D9D9"/></bottom></border></borders><cellXfs count="3"><xf fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf></cellXfs></styleSheet>'
    def cell(col,row,val,style=2):
        ref=f"{col}{row}"; v=escape(str(val))
        if isinstance(val,str) and val.startswith('='): return f'<c r="{ref}" s="{style}"><f>{escape(val[1:])}</f></c>'
        return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{v}</t></is></c>'
    def sheet_xml(title):
        headers = SHEETS[title]
        if title == '05 调研结论':
            rows = [["指标","数值"],["计划调研店铺数",20],["已录入店铺数",'=COUNTA(\'03 店铺调研\'!B2:B21)'],["已完成调研数",'=COUNTIF(\'03 店铺调研\'!AG2:AG21,"已完成")'],["直接竞品数量",'=COUNTIF(\'03 店铺调研\'!C2:C21,"直接竞品")'],["同品类优秀店铺数量",'=COUNTIF(\'03 店铺调研\'!C2:C21,"同品类优秀店铺")'],["跨品类参考店铺数量",'=COUNTIF(\'03 店铺调研\'!C2:C21,"跨品类参考店铺")'],["平均综合评分",'=IFERROR(ROUND(AVERAGE(\'03 店铺调研\'!AE2:AE21),1),"")'],["最高评分店铺",'=IFERROR(INDEX(\'03 店铺调研\'!B2:B21,MATCH(MAX(\'03 店铺调研\'!AE2:AE21),\'03 店铺调研\'!AE2:AE21,0)),"暂无数据")'],["已收集优秀案例数量",'=COUNTA(\'04 优秀案例\'!A2:A41)'],[],["分析维度","竞品普遍做法","优秀店铺做法","常见问题","对我们的启示"]] + [[d] for d in ["首页首屏","页面结构","导航设计","产品分类","视觉风格","品牌故事","场景图片","产品卖点","产品展示","页面跳转","转化路径"]] + [[],["编号","优秀做法","参考品牌","参考案例编号","适用页面","重要程度","是否建议采用"]] + [[] for _ in range(CONCLUSION_PRACTICES_END_ROW - CONCLUSION_PRACTICES_START_ROW + 1)]
        else:
            n={"02 竞品清单":20,"03 店铺调研":20,"04 优秀案例":40,"06 装修建议":30}[title]
            rows=[headers]+[[i]+[""]*(len(headers)-1) for i in range(1,n+1)]
            if title=='03 店铺调研':
                for r in range(2,22): rows[r-1][30]=f'=IF(COUNT(Y{r}:AD{r})=0,"",ROUND(AVERAGE(Y{r}:AD{r}),1))'
        body=''
        for r, vals in enumerate(rows,1):
            cells=''.join(cell(chr(65+c),r,v,1 if r==1 or (title=='05 调研结论' and r in [12,25]) else 2) for c,v in enumerate(vals))
            body += f'<row r="{r}">{cells}</row>'
        return f'<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData>{body}</sheetData><autoFilter ref="A1:Z1"/></worksheet>'
    with ZipFile(OUTPUT, 'w', ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types); z.writestr('_rels/.rels', rels); z.writestr('xl/workbook.xml', wbxml); z.writestr('xl/_rels/workbook.xml.rels', wbrels); z.writestr('xl/styles.xml', styles)
        for i,s in enumerate(sheets,1): z.writestr(f'xl/worksheets/sheet{i}.xml', sheet_xml(s))


def main() -> None:
    try:
        build_with_openpyxl()
        engine = "openpyxl"
    except ModuleNotFoundError as exc:
        if exc.name != "openpyxl":
            raise
        fallback_minimal_xlsx()
        engine = "stdlib fallback（当前环境无法安装 openpyxl）"
    assert OUTPUT.exists() and OUTPUT.stat().st_size > 0
    with ZipFile(OUTPUT) as z:
        assert 'xl/workbook.xml' in z.namelist()
    print(f"已使用 {engine} 生成并验证：{OUTPUT.resolve()}")

if __name__ == "__main__":
    main()
