#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate Word summary document for 郝宇青."""

from pathlib import Path

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

OUTPUT_PATH = Path(__file__).parent / "郝宇青-学术合作与研究贡献总结.docx"


def set_cell_font(cell, text, bold=False, size=10.5):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")


def add_title(doc, text, level=0):
    if level == 0:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(22)
        run.font.name = "黑体"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
    else:
        doc.add_heading(text, level=level)


def add_body(doc, text, bold=False, indent=False):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.74)
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(12)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    p.paragraph_format.line_spacing = 1.5
    return p


def add_table(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        set_cell_font(table.rows[0].cells[i], h, bold=True)
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            set_cell_font(table.rows[r_idx + 1].cells[c_idx], str(val))
    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph()
    return table


def build_document():
    doc = Document()

    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.17)
        section.right_margin = Cm(3.17)

    add_title(doc, "郝宇青学术概况与合作研究总结")
    add_body(doc, "（周敏晖各论课申报材料配套文档）")
    doc.add_paragraph()

    add_title(doc, "一、基本概况", level=1)
    add_body(
        doc,
        "郝宇青，男，1970年生，山东人，法学博士，华东师范大学公共管理学院教授、博士生导师，"
        "当代中国政治发展与战略研究所所长。长期从事国际共产主义运动史（苏联政治方向）、"
        "当代中国政治、基层社会治理等领域研究，发表学术论文百余篇，出版学术著作多部，"
        "主持国家社科基金、教育部人文社科研究等各类课题十余项。",
        indent=True,
    )
    add_body(
        doc,
        "现任中国国际共运史学会理事、中国苏联东欧史研究会理事、中国科学社会主义学会理事，"
        "上海市科学社会主义学会副会长、上海市政治学会常务理事、上海市马克思主义研究会理事，"
        "上海市习近平新时代中国特色社会主义思想研究中心研究员，全国基层党建研究中心特邀研究员，"
        "上海市闵行区党建研究中心顾问等。",
        indent=True,
    )

    add_title(doc, "二、研究方向与学术特色", level=1)
    add_body(
        doc,
        "郝宇青教授的研究横跨三个相互关联的维度：一是苏联政治与社会主义运动史，"
        "以执政合法性、政治信任、政党建设等议题见长；二是当代中国政治，"
        "关注社会转型中的政治信任、新媒体政治、社会情绪治理等问题；"
        "三是基层社会治理，特别是城市社区党建、基层党组织建设质量提升等前沿议题。",
        indent=True,
    )
    add_body(
        doc,
        "其学术特色在于：善于运用比较政治视野，将苏联及国际共运经验与中国基层治理实践相结合；"
        "注重从“主体—制度/过程—文化”三维框架分析组织建设问题；"
        "长期参与上海市及全国基层党建政策咨询与理论研究，具有扎实的田野调研基础。",
        indent=True,
    )

    add_title(doc, "三、与周敏晖的合作研究", level=1)
    add_body(
        doc,
        "郝宇青教授与课程负责人周敏晖（上海商学院马克思主义学院副教授）"
        "在城市社区党建、基层治理领域保持长期稳定的学术合作关系。"
        "周敏晖博士师从郝宇青教授，在华东师范大学攻读政治学理论相关学位期间，"
        "即以上海市闵行区江川路街道为长期田野样本点，开展团队党建、社区共同体建设等专题研究。"
        "二人合作成果已发表于《党政研究》《当代世界社会主义问题》等CSSCI期刊，"
        "为周敏晖各论课“从组织路线到组织再造”提供了重要的理论支撑与学术背书。",
        indent=True,
    )

    add_title(doc, "四、合作论文核心理论贡献", level=1)

    add_body(doc, "（一）《主体·制度·文化：城市社区支部建设质量的三维分析》", bold=True)
    add_body(
        doc,
        "发表于《党政研究》2020年第2期。提出提高城市社区支部建设质量应从主体、制度和文化三个维度入手："
        "主体是核心——支部主体强健则领导力强；制度是保障——为支部建设提供行动指南；"
        "文化是支撑——保持党支部先进性和纯洁性。该文为后续“主体—过程—文化”分析框架奠定了理论基础。",
        indent=True,
    )

    add_body(doc, "（二）《“团队党建”：城市社区党建工作的新探索——以上海市江川路街道为例》", bold=True)
    add_body(
        doc,
        "发表于《当代世界社会主义问题》2019年第3期。系统总结上海市江川路街道“团队党建”创新实践，"
        "提出“支部领导团队、党员融入团队、团队凝聚群众”三机制，"
        "突破传统以地缘、业缘为边界的党组织设置方式，"
        "以“趣缘”为纽带实现混合式、组合式党支部创新，"
        "为破解城市社区“组织弱化”问题提供了可复制的组织再造路径。",
        indent=True,
    )

    add_body(doc, "（三）《构建城市社区共同体：党建何以引领——基于“主体—过程—文化”三维视角的分析》", bold=True)
    add_body(
        doc,
        "发表于《党政研究》2025年第2期。在既有研究基础上，将三维框架从“支部建设质量”"
        "拓展至“社区共同体构建”：主体之维强调社区党组织引领意识与引领能力；"
        "过程之维关注党建引领方式与议题设置；文化之维聚焦主流文化、文化资源与公共精神。"
        "该文直接支撑周敏晖各论课第8周“党建引领：主体—过程—文化”专题教学。",
        indent=True,
    )

    add_body(doc, "（四）其他合作成果", bold=True)
    add_body(
        doc,
        "二人还合作发表《实现美好生活必须构建城市社区共同体——学习习近平关于社区治理的重要论述》"
        "（《党政研究》2022年第6期）等论文，形成从支部建设→团队党建→社区共同体的"
        "递进式研究链条，与周敏晖专著《城市社区治理的组织再造研究》（2022）形成专著—论文互证体系。",
        indent=True,
    )

    add_title(doc, "五、合作研究成果对照表", level=1)
    add_table(
        doc,
        ["论文/成果", "发表信息", "核心理论工具", "与课程关系"],
        [
            [
                "主体·制度·文化：城市社区支部建设质量的三维分析",
                "《党政研究》2020年第2期",
                "主体—制度—文化三维",
                "分析框架源头，第8讲理论基础",
            ],
            [
                "“团队党建”：城市社区党建工作的新探索",
                "《当代世界社会主义问题》2019年第3期",
                "团队党建三机制",
                "核心案例，第5—6讲",
            ],
            [
                "构建城市社区共同体：党建何以引领",
                "《党政研究》2025年第2期",
                "主体—过程—文化三维",
                "落地链工具，第8讲",
            ],
            [
                "实现美好生活必须构建城市社区共同体",
                "《党政研究》2022年第6期",
                "社区共同体理论",
                "原著导读与论述解读",
            ],
        ],
        col_widths=[4.5, 4, 3.5, 4],
    )

    add_title(doc, "六、对周敏晖各论课建设的意义", level=1)
    points = [
        (
            "学术指导与质量保障",
            "郝宇青教授作为周敏晖的学术导师与合作者，"
            "为课程“谱系—再造—落地”三链模型提供了经过同行评议的理论验证，"
            "确保课程内容不是简单的政策解读，而是有扎实学术根基的研究转化。",
        ),
        (
            "案例资源的学理深度",
            "江川路街道“团队党建”案例经郝宇青—周敏晖合作研究多年，"
            "已形成从概念提炼、机制分析到共同体构建的完整学理链条，"
            "使课程核心案例具有不可替代的学术价值。",
        ),
        (
            "分析工具的原创性",
            "“主体—过程—文化”三维框架、“团队党建三机制”等课堂分析工具"
            "均源于二人合作发表的CSSCI论文，学生可直接运用这些工具进行案例分析，"
            "实现“研教一体”。",
        ),
        (
            "跨校协同的示范效应",
            "华东师范大学（郝宇青）与上海商学院（周敏晖）的跨校合作，"
            "体现了高校马克思主义学院之间理论研究—教学实践协同创新的良好范式，"
            "有利于课程申报中展示团队学术背景与协作网络。",
        ),
    ]
    for i, (title, content) in enumerate(points, 1):
        add_body(doc, f"{i}. {title}", bold=True)
        add_body(doc, content, indent=True)

    add_title(doc, "七、郝宇青教授主要代表性成果（精选）", level=1)
    add_table(
        doc,
        ["类别", "成果名称", "说明"],
        [
            ["著作", "《苏联国家与社会的关系研究》", "国家社科基金项目成果"],
            ["著作", "《严肃党内政治生活的顶层设计》（合著）", "党建理论研究"],
            ["论文", "《执政合法性资源的再生产：中国共产党的重要课题》", "《探索》2007年第5期"],
            ["论文", "《社会情绪民粹化：形成机理及其消解》", "《人民论坛·学术前沿》2019年第17期"],
            ["论文", "《以内生发展驱动“三新”党建工作提质增效》", "《人民论坛·学术前沿》2025年"],
            ["课题", "研究阐释党的十九大精神国家社科基金专项", "18VSJ102，基层党组织建设"],
            ["课题", "教育部重大攻关项目子课题", "“提高党的建设科学化水平研究”"],
        ],
        col_widths=[2, 7, 7],
    )

    add_title(doc, "八、总结", level=1)
    add_body(
        doc,
        "郝宇青教授是城市社区党建与基层治理领域具有重要影响力的学者，"
        "与课程负责人周敏晖的合作研究形成了“导师—学生—合作者”三位一体的学术关系。"
        "二人合作发表的多篇CSSCI论文，特别是“团队党建三机制”和“主体—过程—文化”三维框架，"
        "构成了周敏晖各论课“从组织路线到组织再造”的核心理论资源与课堂分析工具。"
        "在课程建设中，应充分标注合作论文来源，体现学术传承与跨校协同，"
        "同时建议在课程申报材料的“学术团队”或“研究基础”部分明确郝宇青教授的学术指导与合作关系。",
        indent=True,
    )

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run("整理人：______________\n日期：2026年7月4日")
    run.font.size = Pt(12)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    doc.save(OUTPUT_PATH)
    return OUTPUT_PATH


if __name__ == "__main__":
    path = build_document()
    print(f"Generated: {path}")
