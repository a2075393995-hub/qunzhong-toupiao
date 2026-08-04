from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
SCREENSHOT_DIR = ROOT / "qa" / "delivery_screenshots"
OUTPUT = ROOT / "qa" / "群众选票格式化打印工具_使用说明.docx"


def set_run_font(run, name="Microsoft YaHei UI", size=10.5, bold=False, color=None):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def add_para(document, text="", size=10.5, bold=False, color=None, style=None):
    p = document.add_paragraph(style=style)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    return p


def add_heading(document, text, level=1):
    p = document.add_paragraph()
    p.paragraph_format.space_before = Pt(12 if level == 1 else 8)
    p.paragraph_format.space_after = Pt(5)
    run = p.add_run(text)
    set_run_font(run, size=15 if level == 1 else 12.5, bold=True, color="1F4D78" if level == 1 else "2E74B5")
    return p


def add_bullet(document, text):
    p = document.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(text)
    set_run_font(run, size=10.5)
    return p


def shade_cell(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_text(cell, text, bold=False):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(text)
    set_run_font(run, size=10, bold=bold)


def add_image(document, path, caption):
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(str(path), width=Inches(6.4))
    cap = document.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(8)
    r = cap.add_run(caption)
    set_run_font(r, size=9, color="6B7280")


def build():
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    styles = document.styles
    normal = styles["Normal"]
    normal.font.name = "Microsoft YaHei UI"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei UI")
    normal.font.size = Pt(10.5)

    document.core_properties.title = "群众选票格式化打印工具 使用说明"
    document.core_properties.author = "群众选票格式化打印工具"
    document.core_properties.subject = "发布版操作说明"

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("群众选票格式化打印工具")
    set_run_font(r, size=22, bold=True, color="111827")
    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = subtitle.add_run("发布版使用说明")
    set_run_font(r, size=12, color="4B5563")

    add_para(document, "用途：把 Word 表决票模板和短信/表格投票数据合并，可按每位选民生成一份 DOCX，也可合并导出到一个 DOCX，并输出投票结果汇总。", size=10.5)

    add_heading(document, "一、交付物", 1)
    table = document.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    set_cell_text(table.rows[0].cells[0], "文件/目录", True)
    set_cell_text(table.rows[0].cells[1], "说明", True)
    shade_cell(table.rows[0].cells[0], "E8EEF5")
    shade_cell(table.rows[0].cells[1], "E8EEF5")
    rows = [
        ("群众选票格式化打印工具.exe", "主程序，双击运行。"),
        ("模板\\百家湖西花园小区表决票模板.docx", "示例 Word 模板，可直接用于上传模板。"),
        ("测试数据\\图三_投票数据.csv", "示例投票数据，可直接用于上传数据文件。"),
        ("output", "导出目录。每次导出会自动生成一个按时间命名的子目录。"),
        ("群众选票格式化打印工具_使用说明.docx", "当前说明文档。"),
    ]
    for left, right in rows:
        cells = table.add_row().cells
        set_cell_text(cells[0], left)
        set_cell_text(cells[1], right)

    add_heading(document, "二、主界面", 1)
    add_para(document, "顶部是操作流程按钮，左侧是字段/投票结果/格式刷，右侧是模板预览，底部是数据预览和调试信息。")
    loaded = SCREENSHOT_DIR / "main_loaded_sample.png"
    blank = SCREENSHOT_DIR / "main_release.png"
    if loaded.exists():
        add_image(document, loaded, "图 1：使用桌面模板和测试数据打开后的主界面")
    elif blank.exists():
        add_image(document, blank, "图 1：发布版主界面")

    add_heading(document, "三、快速流程", 1)
    steps = [
        "点击“1 上传模板”，选择 Word 模板文件。",
        "点击“2 上传数据文件”，选择 CSV 或 Excel 投票数据，并设置投票结果数量检查规则。",
        "点击“3 调试模板”，在左侧选择字段或投票结果，再在右侧模板预览中点击判断区和标记区。",
        "点击“4 调试完成”。",
        "点击“5 导出前预览”，在程序内部检查第一份生成效果。",
        "预览确认后点击“6 开始导出”，选择多文件/单文件导出模式，按需勾选纯净模式。",
    ]
    for index, step in enumerate(steps, start=1):
        add_para(document, f"{index}. {step}")

    add_heading(document, "四、调试模板要点", 1)
    bullets = [
        "字段填入：房号/地址、姓名、电话号码可以在左侧选择后，直接点击模板中的下划线或表格位置。",
        "投票结果：结果1、结果2、结果3来自数据文件中的投票选项顺序。",
        "判断区：点击模板里能代表选项的文字区域，例如“选项①”“同意”“反对”。",
        "标记区：点击真正需要打勾的位置，一般是“表决意见”所在空格或下划线。",
        "多选集合：按 Ctrl 或 Shift 选择多个结果，会显示“集合1”。集合表示这些结果共用同一套判断区/标记区规则。",
        "精确型数量检查会保留原始投票内容且不会去重：如果一条数据里 3 个投票结果相同，仍按 3 个结果处理。",
        "范围型数量检查会保留原始投票内容：例如数据写 1、2、3，程序预览和导出匹配时仍以 1、2、3 为准。",
        "用于匹配模板判断区时，选项1 和 选项① 仍可识别为同一个判断条件。",
        "如果要重来，点击顶部“重置”，会清空当前模板、数据、标注和预览，但不会删除已经导出的文件。",
    ]
    for item in bullets:
        add_bullet(document, item)

    add_heading(document, "五、导出结果", 1)
    add_para(document, "导出时，程序会在输出目录下自动新建时间文件夹，例如：")
    add_para(document, "output\\2026-7-21-14-31-55", size=10.5, bold=True, color="111827")
    for item in [
        "多文件导出：每位选民生成一个 DOCX，文件名默认使用姓名，也可以在“文件名前缀”中追加前缀。",
        "单文件导出：所有正常数据合并到一个 DOCX，每条数据从新页开始追加。上传模板时请检查模板最好只占 1 页；如果模板超过 1 页，程序仍会从上一条完整内容后另起新页追加。",
        "纯净模式：导出的可见内容只保留程序写入的房号/地址、姓名、电话号码和打勾，模板原有文字会被隐藏，表格结构会被删除，适合套打到已经印好的纸质模板上。",
        "投票结果汇总.xlsx：包含“成功导出的数据”和“未导出的数据”两个工作表。",
        "未导出的数据：票数过多或过少的数据会进入该工作表，结果列显示“废票”。",
    ]:
        add_bullet(document, item)

    add_heading(document, "六、测试数据说明", 1)
    add_para(document, "交付包中已经放入桌面测试模板和测试数据的副本。第一次试用建议直接选择：")
    add_bullet(document, "模板：模板\\百家湖西花园小区表决票模板.docx")
    add_bullet(document, "数据：测试数据\\图三_投票数据.csv")
    add_bullet(document, "数量检查：精确型，填写 3 个。示例数据中第 5 行有 4 个结果，会进入“未导出的数据”。")

    add_heading(document, "七、常见问题", 1)
    qas = [
        ("为什么预览只看第一条数据？", "导出前预览会优先展示第一条正常数据，用于检查模板位置是否正确。批量导出会按全部正常数据生成。"),
        ("为什么某个结果没有打勾？", "请检查该条数据里对应结果实际是什么；集合只是共享规则，不表示固定打所有选项。"),
        ("可以重复调试吗？", "可以。导出前预览点“返回修改”，或点击“上一步”“重置”重新调整。"),
    ]
    for q, a in qas:
        add_para(document, q, bold=True)
        add_para(document, a)

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    document.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    build()
