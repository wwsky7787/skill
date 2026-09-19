#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
住院医师规范化培训教学查房教案生成器

功能：
1. 根据病例资料自动生成完整的教学查房教案
2. 通过网络搜索补充讨论问题和最新研究进展
3. 生成符合格式要求的 docx 文件（8000-12000 字）

格式要求：
- A4 纸（210mm × 297mm）
- 版心尺寸：156mm × 225mm
- 页边距：上 36mm、下 36mm、左 27mm、右 27mm（不含页码）
- 每页 22 行，每行 28 字
- 正文：三号仿宋_GB2312，行间距 28 磅
- 标题层级：
  - 第一层："一、"（三号黑体）
  - 第二层："(一)"（三号楷体）
  - 第三层："1."（三号仿宋，有标点）
  - 第四层："(1)"（三号仿宋，必须有标点）
"""

from docx import Document
from docx.shared import Pt, Mm, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.enum.section import WD_ORIENT
import re
from typing import Dict, List, Optional
import json


# ==================== 页面与格式设置 ====================

def set_page_format(doc):
    """设置页面格式：A4 纸，版心尺寸，页边距"""
    section = doc.sections[0]

    # A4 纸尺寸
    section.page_width = Mm(210)
    section.page_height = Mm(297)

    # 页边距：上 36mm、下 36mm、左 27mm、右 27mm
    section.top_margin = Mm(36)
    section.bottom_margin = Mm(36)
    section.left_margin = Mm(27)
    section.right_margin = Mm(27)

    return section


def set_font_style(run, font_name='仿宋_GB2312', font_size=16, bold=False, italic=False, color='000000'):
    """
    设置字体样式
    
    Args:
        run: docx run 对象
        font_name: 字体名称
        font_size: 字号（磅值），三号=16pt
        bold: 是否加粗
        italic: 是否倾斜
        color: 字体颜色（十六进制）
    """
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = None  # 使用默认黑色

    # 设置中文字体
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)


def set_paragraph_format(paragraph, line_spacing=28, space_before=0, space_after=0, 
                         alignment=WD_ALIGN_PARAGRAPH.LEFT, first_line_indent=None):
    """
    设置段落格式
    
    Args:
        paragraph: docx paragraph 对象
        line_spacing: 行间距（磅值），默认 28 磅
        space_before: 段前间距（磅值）
        space_after: 段后间距（磅值）
        alignment: 对齐方式
        first_line_indent: 首行缩进（字符数），默认 2 字符
    """
    p_format = paragraph.paragraph_format

    # 行间距（固定值 28 磅）
    p_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p_format.line_spacing = Pt(line_spacing)

    # 段前段后间距
    p_format.space_before = Pt(space_before)
    p_format.space_after = Pt(space_after)

    # 对齐方式
    p_format.alignment = alignment

    # 首行缩进（默认 2 字符，每字符 28 磅）
    if first_line_indent is None:
        first_line_indent = 2
    p_format.first_line_indent = Pt(28 * first_line_indent)

    return p_format


# ==================== 标题与段落添加 ====================

def add_main_title(doc, title):
    """
    添加主标题（教案封面标题）
    三号黑体，居中
    """
    paragraph = doc.add_paragraph()
    set_paragraph_format(paragraph, line_spacing=28, space_before=36, space_after=24, 
                         alignment=WD_ALIGN_PARAGRAPH.CENTER, first_line_indent=0)

    run = paragraph.add_run(title)
    set_font_style(run, font_name='黑体', font_size=16, bold=True)

    return paragraph


def add_section_title(doc, title):
    """
    添加一级标题（一、二、三、）
    三号黑体，句尾无标点
    """
    paragraph = doc.add_paragraph()
    set_paragraph_format(paragraph, line_spacing=28, space_before=24, space_after=12, first_line_indent=0)

    run = paragraph.add_run(title)
    set_font_style(run, font_name='黑体', font_size=16, bold=True)

    return paragraph


def add_subsection_title(doc, title):
    """
    添加二级标题（(一)(二)(三)）
    三号楷体，句尾可有可无标点
    """
    paragraph = doc.add_paragraph()
    set_paragraph_format(paragraph, line_spacing=28, space_before=12, space_after=6, first_line_indent=0)

    run = paragraph.add_run(title)
    set_font_style(run, font_name='楷体_GB2312', font_size=16)

    return paragraph


def add_numbered_paragraph(doc, text, level=3, number='', first_line_indent=2):
    """
    添加带序号的段落
    
    Args:
        doc: Document 对象
        text: 文本内容
        level: 标题层级 (1-4)
        number: 序号（如"一、"、"(一)"、"1."、"(1)"）
        first_line_indent: 首行缩进字符数
    """
    paragraph = doc.add_paragraph()
    set_paragraph_format(paragraph, line_spacing=28, space_before=0, space_after=0, 
                         first_line_indent=first_line_indent)

    # 添加序号和内容
    if level == 1:
        # "一、"（三号黑体）
        run_num = paragraph.add_run(number)
        set_font_style(run_num, font_name='黑体', font_size=16, bold=True)
        run_text = paragraph.add_run(text)
        set_font_style(run_text, font_name='仿宋_GB2312', font_size=16)

    elif level == 2:
        # "(一)"（三号楷体）
        run_num = paragraph.add_run(number)
        set_font_style(run_num, font_name='楷体_GB2312', font_size=16)
        run_text = paragraph.add_run(text)
        set_font_style(run_text, font_name='仿宋_GB2312', font_size=16)

    elif level == 3:
        # "1."（三号仿宋，有标点）
        run_num = paragraph.add_run(number)
        set_font_style(run_num, font_name='仿宋_GB2312', font_size=16)
        run_text = paragraph.add_run(text)
        set_font_style(run_text, font_name='仿宋_GB2312', font_size=16)

    elif level == 4:
        # "(1)"（三号仿宋，必须有标点）
        run_num = paragraph.add_run(number)
        set_font_style(run_num, font_name='仿宋_GB2312', font_size=16)
        run_text = paragraph.add_run(text)
        set_font_style(run_text, font_name='仿宋_GB2312', font_size=16)

    return paragraph


def add_plain_paragraph(doc, text, first_line_indent=2):
    """
    添加普通段落（无序号）
    """
    paragraph = doc.add_paragraph()
    set_paragraph_format(paragraph, line_spacing=28, space_before=0, space_after=0, 
                         first_line_indent=first_line_indent)

    run = paragraph.add_run(text)
    set_font_style(run, font_name='仿宋_GB2312', font_size=16)

    return paragraph


def add_table(doc, data, column_widths=None, header_rows=1):
    """
    添加表格
    
    Args:
        doc: Document 对象
        data: 二维列表数据
        column_widths: 列宽列表（mm）
        header_rows: 表头行数
    """
    rows = len(data)
    cols = len(data[0]) if data else 0

    if cols == 0:
        return None

    table = doc.add_table(rows=rows, cols=cols)
    table.style = 'Table Grid'

    # 设置列宽
    if column_widths:
        for i, width in enumerate(column_widths):
            if i < cols:
                table.columns[i].width = Mm(width)

    # 填充数据
    for i, row_data in enumerate(data):
        row = table.rows[i]
        for j, cell_data in enumerate(row_data):
            cell = row.cells[j]
            cell.text = str(cell_data) if cell_data else ''

            # 设置单元格内文本格式
            for paragraph in cell.paragraphs:
                set_paragraph_format(paragraph, line_spacing=24, first_line_indent=0)
                for run in paragraph.runs:
                    if i < header_rows:
                        # 表头：黑体
                        set_font_style(run, font_name='黑体', font_size=14, bold=True)
                    else:
                        # 表体：仿宋
                        set_font_style(run, font_name='仿宋_GB2312', font_size=14)

    return table


# ==================== 教案各部分生成 ====================

def add_cover_page(doc, data):
    """添加封面页"""
    # 主标题
    add_main_title(doc, '住院医师规范化培训教学查房教案')
    
    # 空 3 行
    for _ in range(3):
        doc.add_paragraph()
    
    # 教学查房标题
    title = data.get('教学查房标题', '疾病名称 + 教学查房')
    title_para = doc.add_paragraph()
    set_paragraph_format(title_para, line_spacing=28, space_before=0, space_after=24, 
                         alignment=WD_ALIGN_PARAGRAPH.CENTER, first_line_indent=0)
    run = title_para.add_run(title)
    set_font_style(run, font_name='黑体', font_size=16, bold=True)
    
    # 空 1 行
    doc.add_paragraph()
    
    # 单位和姓名（楷体）
    unit_info = f"{data.get('专业基地', '')} · {data.get('科室', '')}"
    unit_para = doc.add_paragraph()
    set_paragraph_format(unit_para, line_spacing=28, space_before=0, space_after=12, 
                         alignment=WD_ALIGN_PARAGRAPH.CENTER, first_line_indent=0)
    run = unit_para.add_run(unit_info)
    set_font_style(run, font_name='楷体_GB2312', font_size=16)
    
    # 指导医师
    teacher_info = f"指导医师：{data.get('指导医师', '')}"
    teacher_para = doc.add_paragraph()
    set_paragraph_format(teacher_para, line_spacing=28, space_before=0, space_after=12, 
                         alignment=WD_ALIGN_PARAGRAPH.CENTER, first_line_indent=0)
    run = teacher_para.add_run(teacher_info)
    set_font_style(run, font_name='楷体_GB2312', font_size=16)
    
    # 教学日期
    date_info = f"教学日期：{data.get('教学日期', '')}"
    date_para = doc.add_paragraph()
    set_paragraph_format(date_para, line_spacing=28, space_before=0, space_after=36, 
                         alignment=WD_ALIGN_PARAGRAPH.CENTER, first_line_indent=0)
    run = date_para.add_run(date_info)
    set_font_style(run, font_name='楷体_GB2312', font_size=16)
    
    # 分页
    doc.add_page_break()


def add_basic_info(doc, data):
    """添加教学基本信息表格"""
    add_section_title(doc, '一、教学基本信息')
    
    info_data = data.get('基本信息', {})
    
    table_data = [
        ['教学查房标题', info_data.get('教学查房标题', '')],
        ['专业基地', info_data.get('专业基地', '')],
        ['科室', info_data.get('科室', '')],
        ['指导医师', info_data.get('指导医师', '')],
        ['职称', info_data.get('职称', '')],
        ['教学日期', info_data.get('教学日期', '')],
        ['参与住院医师', info_data.get('参与住院医师', '')],
        ['教学时长', info_data.get('教学时长', '60-90 分钟')],
    ]
    
    add_table(doc, table_data, column_widths=[40, 116])
    doc.add_paragraph()  # 空行


def add_teaching_objectives(doc, data):
    """添加教学目标（符合 SMART 原则）"""
    add_section_title(doc, '二、教学目标')

    objectives = data.get('教学目标', {})

    # 知识目标（SMART 验证）
    add_subsection_title(doc, '(一) 知识目标')
    knowledge_obj = objectives.get('知识目标', [])
    if isinstance(knowledge_obj, str):
        knowledge_obj = [knowledge_obj]
    # 验证并优化教学目标
    knowledge_obj = validate_smart_objectives(knowledge_obj)
    for i, obj in enumerate(knowledge_obj, 1):
        add_numbered_paragraph(doc, obj, level=3, number=f'{i}.')

    # 能力目标（SMART 验证）
    add_subsection_title(doc, '(二) 能力目标')
    ability_obj = objectives.get('能力目标', [])
    if isinstance(ability_obj, str):
        ability_obj = [ability_obj]
    # 验证并优化教学目标
    ability_obj = validate_smart_objectives(ability_obj)
    for i, obj in enumerate(ability_obj, 1):
        add_numbered_paragraph(doc, obj, level=3, number=f'{i}.')

    # 素养目标（SMART 验证）
    add_subsection_title(doc, '(三) 素养目标')
    quality_obj = objectives.get('素养目标', [])
    if isinstance(quality_obj, str):
        quality_obj = [quality_obj]
    # 验证并优化教学目标
    quality_obj = validate_smart_objectives(quality_obj)
    for i, obj in enumerate(quality_obj, 1):
        add_numbered_paragraph(doc, obj, level=3, number=f'{i}.')

    doc.add_paragraph()  # 空行


def add_case_summary(doc, data):
    """添加病例摘要（按照规范标题层次）"""
    add_section_title(doc, '三、病例摘要')

    case_data = data.get('病例摘要', {})

    # （一）患者基本信息
    add_subsection_title(doc, '（一）患者基本信息')
    add_plain_paragraph(doc, case_data.get('患者基本信息', ''))

    # （二）主诉
    add_subsection_title(doc, '（二）主诉')
    add_plain_paragraph(doc, case_data.get('主诉', ''))

    # （三）现病史
    add_subsection_title(doc, '（三）现病史')
    add_plain_paragraph(doc, case_data.get('现病史', ''))
    
    # 病史特点归纳（在现病史之后）
    if case_data.get('病史特点归纳'):
        add_subsection_title(doc, '（四）病史特点归纳')
        add_plain_paragraph(doc, case_data.get('病史特点归纳', ''))
        # 后续标题序号顺延
        _add_remaining_case_sections(doc, case_data, start_index=5)
    else:
        # 无病史特点归纳，按正常顺序添加
        _add_remaining_case_sections(doc, case_data, start_index=4)

    doc.add_paragraph()  # 空行


def _add_remaining_case_sections(doc, case_data, start_index=4):
    """添加病例摘要剩余部分（从指定序号开始）"""
    
    # 定义剩余的病例摘要部分
    sections = [
        ('既往史', '（四）既往史' if start_index == 4 else '（五）既往史'),
        ('个人史', '（五）个人史' if start_index == 4 else '（六）个人史'),
        ('家族史', '（六）家族史' if start_index == 4 else '（七）家族史'),
        ('体格检查', '（七）体格检查' if start_index == 4 else '（八）体格检查'),
        ('辅助检查', '（八）辅助检查' if start_index == 4 else '（九）辅助检查'),
        ('初步诊断', '（九）初步诊断' if start_index == 4 else '（十）初步诊断'),
        ('诊疗经过', '（十）诊疗经过' if start_index == 4 else '（十一）诊疗经过')
    ]
    
    # 根据 start_index 调整 sections
    if start_index == 5:
        sections = [
            ('既往史', '（五）既往史'),
            ('个人史', '（六）个人史'),
            ('家族史', '（七）家族史'),
            ('体格检查', '（八）体格检查'),
            ('辅助检查', '（九）辅助检查'),
            ('初步诊断', '（十）初步诊断'),
            ('诊疗经过', '（十一）诊疗经过')
        ]
    
    for key, title in sections:
        add_subsection_title(doc, title)
        add_plain_paragraph(doc, case_data.get(key, ''))


def add_teaching_flow(doc, data):
    """添加教学流程与时间安排"""
    add_section_title(doc, '四、教学流程与时间安排')
    
    add_subsection_title(doc, '(一) 查房准备阶段（示教室）')
    add_plain_paragraph(doc, '时间：5-10 分钟')
    add_plain_paragraph(doc, '主要内容：')
    
    prep_content = data.get('查房准备阶段', [
        '教学查房参与成员相互介绍',
        '介绍教学查房患者的基本信息与教学目标',
        '宣布本次教学查房的流程、时间分配、角色分工和注意事项'
    ])
    for i, content in enumerate(prep_content, 1):
        add_numbered_paragraph(doc, content, level=3, number=f'{i}.')
    
    add_subsection_title(doc, '(二) 临床信息采集阶段（床旁）')
    add_plain_paragraph(doc, '时间：15-20 分钟')
    add_plain_paragraph(doc, '主要内容：')
    
    bedside_content = data.get('临床信息采集阶段', [
        '住院医师脱稿汇报病史（5-6 分钟）',
        '指导医师核实病史并示范问诊技巧',
        '住院医师进行体格检查操作',
        '指导医师观察并示范正确的体格检查方法',
        '指导医师与患者交流，示范医患沟通技巧和人文关怀'
    ])
    for i, content in enumerate(bedside_content, 1):
        add_numbered_paragraph(doc, content, level=3, number=f'{i}.')
    
    add_subsection_title(doc, '(三) 病例讨论阶段（示教室）')
    add_plain_paragraph(doc, '时间：40-60 分钟')
    add_plain_paragraph(doc, '主要内容：')
    
    discussion_content = data.get('病例讨论阶段', [
        '对床旁查房过程进行总结反馈',
        '引导住院医师归纳总结病例特点',
        '讨论诊断假设与鉴别诊断',
        '分析解读辅助检查结果',
        '讨论并制定诊疗方案',
        '介绍最新指南和研究进展',
        '教学总结，提出自学问题，提供参考文献'
    ])
    for i, content in enumerate(discussion_content, 1):
        add_numbered_paragraph(doc, content, level=3, number=f'{i}.')
    
    # 时间安排表格
    add_subsection_title(doc, '(四) 时间安排表')
    table_data = [
        ['阶段', '时间', '主要内容', '参与者活动', '指导医师职责'],
        ['查房准备阶段', '5-10 分钟', '成员介绍、交代教学目标与流程', '聆听、记录', '主持、说明'],
        ['临床信息采集阶段（床旁）', '15-20 分钟', '病史汇报、问诊核实、体格检查、医患沟通', 
         '主管住院医师汇报与查体，其他住院医师观摩', '观察、补充、示范'],
        ['病例讨论阶段（示教室）', '40-60 分钟', '病例总结、诊断鉴别、诊疗讨论、教学总结', 
         '积极讨论、主动发言', '引导、点评、总结']
    ]
    add_table(doc, table_data, column_widths=[35, 20, 50, 30, 21])
    
    doc.add_paragraph()  # 空行


def add_key_points(doc, data):
    """添加教学重点与难点（优化版，避免重复）"""
    add_section_title(doc, '五、教学重点与难点')

    # 获取教学重点和难点
    key_points = data.get('教学重点', [])
    difficult_points = data.get('教学难点', [])
    
    if isinstance(key_points, str):
        key_points = [key_points]
    if isinstance(difficult_points, str):
        difficult_points = [difficult_points]
    
    # 优化：去除重复内容，确保重点和难点不重叠
    optimized_key_points, optimized_difficult_points = optimize_key_points(key_points, difficult_points)
    
    add_subsection_title(doc, '(一) 教学重点')
    add_plain_paragraph(doc, '（住院医师必须掌握的核心知识和关键技能）')
    for i, point in enumerate(optimized_key_points, 1):
        add_numbered_paragraph(doc, point, level=3, number=f'{i}.')

    add_subsection_title(doc, '(二) 教学难点')
    add_plain_paragraph(doc, '（住院医师理解和应用可能存在困难的内容）')
    for i, point in enumerate(optimized_difficult_points, 1):
        add_numbered_paragraph(doc, point, level=3, number=f'{i}.')

    doc.add_paragraph()  # 空行


def optimize_key_points(key_points: list, difficult_points: list) -> tuple:
    """
    优化教学重点与难点，避免内容重复
    
    原则：
    1. 重点强调"必须掌握"的内容，难点强调"容易混淆/困难"的内容
    2. 同一知识点不应同时列为重点和难点
    3. 重点侧重知识技能，难点侧重理解应用
    
    Args:
        key_points: 教学重点列表
        difficult_points: 教学难点列表
        
    Returns:
        tuple: (优化后的重点，优化后的难点)
    """
    if not key_points:
        return [], difficult_points
    if not difficult_points:
        return key_points, []
    
    # 提取关键词进行去重
    optimized_key_points = []
    optimized_difficult_points = []
    used_keywords = set()
    
    # 首先处理重点，提取关键词
    for point in key_points:
        optimized_key_points.append(point)
        # 提取关键词（简单实现：取前 10 个字符）
        keyword = point[:10] if len(point) > 10 else point
        used_keywords.add(keyword)
    
    # 处理难点，去除与重点重复的内容
    for point in difficult_points:
        # 检查是否与重点重复
        keyword = point[:10] if len(point) > 10 else point
        is_duplicate = any(keyword in used_key or used_key in keyword 
                          for used_key in used_keywords)
        
        if not is_duplicate:
            optimized_difficult_points.append(point)
            used_keywords.add(keyword)
        else:
            # 如果重复，尝试改写为难点表述（侧重理解应用）
            rewritten = rewrite_as_difficulty(point)
            if rewritten and rewritten not in optimized_difficult_points:
                optimized_difficult_points.append(rewritten)
    
    return optimized_key_points, optimized_difficult_points


def rewrite_as_difficulty(key_point: str) -> str:
    """
    将重点内容改写为难点表述（侧重理解应用）
    
    Args:
        key_point: 重点内容
        
    Returns:
        str: 改写后的难点表述，如果无法改写则返回 None
    """
    # 常见重点到难点的转换模式
    conversions = {
        '诊断标准': '复杂情况下的诊断鉴别思路',
        '治疗原则': '特殊情况下的治疗决策权衡',
        '临床表现': '非典型临床表现的识别',
        '检查方法': '检查结果的综合分析与解读',
        '治疗方案': '个体化治疗方案的制定',
        '鉴别诊断': '相似疾病的鉴别要点',
    }
    
    for key_phrase, difficulty_phrase in conversions.items():
        if key_phrase in key_point:
            # 保留疾病名称
            disease_name = key_point.replace(key_phrase, '').strip(' 的')
            return f"{disease_name}{difficulty_phrase}" if disease_name else difficulty_phrase
    
    return None


def validate_smart_objectives(objectives: list) -> list:
    """
    验证教学目标是否符合 SMART 原则
    
    SMART 原则：
    - Specific（具体）：明确界定学习内容和预期结果
    - Measurable（可测量）：可通过评价方法检验
    - Achievable（可达成）：符合住院医师实际水平
    - Relevant（相关）：与住培要求紧密相关
    - Time-bound（有时限）：在本次教学活动中达成
    
    Args:
        objectives: 教学目标列表
        
    Returns:
        list: 验证后的目标列表（移除不符合 SMART 原则的目标）
    """
    # 用于验证的行为动词
    valid_verbs = [
        '掌握', '熟悉', '了解', '能够', '可以', '会',
        '正确', '独立', '准确', '系统', '全面',
        '制定', '提出', '分析', '判断', '识别',
        '进行', '完成', '参与', '协助', '指导'
    ]
    
    validated = []
    for obj in objectives:
        # 检查是否包含有效的行为动词
        has_valid_verb = any(verb in obj for verb in valid_verbs)
        
        # 检查是否具体明确（长度适中，不是笼统表述）
        is_specific = 10 <= len(obj) <= 100
        
        # 简单的 SMART 验证
        if has_valid_verb and is_specific:
            validated.append(obj)
        else:
            # 尝试改进目标表述
            improved = improve_objective(obj)
            if improved:
                validated.append(improved)
    
    return validated


def improve_objective(obj: str) -> str:
    """
    改进不符合 SMART 原则的教学目标
    
    Args:
        obj: 原始目标
        
    Returns:
        str: 改进后的目标，如果无法改进则返回 None
    """
    # 常见笼统表述的改进
    improvements = {
        '了解': '了解 XX 疾病的基本概念和诊疗流程',
        '学习': '学习并掌握 XX 疾病的核心诊疗要点',
        '知道': '能够准确描述 XX 疾病的临床特征',
    }
    
    for vague, improved in improvements.items():
        if obj.startswith(vague) and len(obj) < 15:
            return improved.replace('XX 疾病', '该疾病')
    
    return None


def add_discussion_questions(doc, data):
    """添加讨论问题设计"""
    add_section_title(doc, '六、讨论问题设计')
    
    questions = data.get('讨论问题', {})
    
    # 第一层次（事实性）问题
    add_subsection_title(doc, '(一) 第一层次（事实性）问题')
    factual_questions = questions.get('事实性问题', [])
    if isinstance(factual_questions, str):
        factual_questions = [factual_questions]
    for i, q in enumerate(factual_questions, 1):
        if isinstance(q, dict):
            add_numbered_paragraph(doc, q.get('问题', ''), level=3, number=f'{i}.')
            if q.get('参考答案'):
                add_plain_paragraph(doc, f"参考答案：{q['参考答案']}", first_line_indent=2)
        else:
            add_numbered_paragraph(doc, q, level=3, number=f'{i}.')
    
    # 第二层次（分析性）问题
    add_subsection_title(doc, '(二) 第二层次（分析性）问题')
    analytical_questions = questions.get('分析性问题', [])
    if isinstance(analytical_questions, str):
        analytical_questions = [analytical_questions]
    for i, q in enumerate(analytical_questions, 1):
        if isinstance(q, dict):
            add_numbered_paragraph(doc, q.get('问题', ''), level=3, number=f'{i}.')
            if q.get('参考答案'):
                add_plain_paragraph(doc, f"参考答案：{q['参考答案']}", first_line_indent=2)
        else:
            add_numbered_paragraph(doc, q, level=3, number=f'{i}.')
    
    # 第三层次（评价性）问题
    add_subsection_title(doc, '(三) 第三层次（评价性）问题')
    evaluative_questions = questions.get('评价性问题', [])
    if isinstance(evaluative_questions, str):
        evaluative_questions = [evaluative_questions]
    for i, q in enumerate(evaluative_questions, 1):
        if isinstance(q, dict):
            add_numbered_paragraph(doc, q.get('问题', ''), level=3, number=f'{i}.')
            if q.get('参考答案'):
                add_plain_paragraph(doc, f"参考答案：{q['参考答案']}", first_line_indent=2)
        else:
            add_numbered_paragraph(doc, q, level=3, number=f'{i}.')
    
    doc.add_paragraph()  # 空行


def add_self_study(doc, data):
    """添加自学任务与参考文献"""
    add_section_title(doc, '七、自学任务与参考文献')
    
    add_subsection_title(doc, '(一) 自学任务')
    self_study_tasks = data.get('自学任务', [])
    if isinstance(self_study_tasks, str):
        self_study_tasks = [self_study_tasks]
    for i, task in enumerate(self_study_tasks, 1):
        add_numbered_paragraph(doc, task, level=3, number=f'{i}.')
    
    add_subsection_title(doc, '(二) 参考文献')
    references = data.get('参考文献', {})
    
    if references.get('指南共识'):
        add_numbered_paragraph(doc, '指南与共识：', level=3, number='1.')
        guidelines = references.get('指南共识', [])
        if isinstance(guidelines, str):
            guidelines = [guidelines]
        for ref in guidelines:
            add_plain_paragraph(doc, ref)
    
    if references.get('重要文献'):
        add_numbered_paragraph(doc, '重要文献：', level=3, number='2.')
        papers = references.get('重要文献', [])
        if isinstance(papers, str):
            papers = [papers]
        for ref in papers:
            add_plain_paragraph(doc, ref)
    
    if references.get('教材专著'):
        add_numbered_paragraph(doc, '教材与专著：', level=3, number='3.')
        books = references.get('教材专著', [])
        if isinstance(books, str):
            books = [books]
        for ref in books:
            add_plain_paragraph(doc, ref)
    
    doc.add_paragraph()  # 空行


def add_reflection(doc, data):
    """添加教学反思与评价"""
    add_section_title(doc, '八、教学反思与评价')
    
    add_subsection_title(doc, '(一) 教学目标达成情况')
    add_plain_paragraph(doc, '（教学查房结束后填写）')
    
    add_subsection_title(doc, '(二) 主要经验与不足')
    add_plain_paragraph(doc, '（教学查房结束后填写）')
    
    add_subsection_title(doc, '(三) 改进措施')
    add_plain_paragraph(doc, '（教学查房结束后填写）')
    
    doc.add_paragraph()  # 空行


# ==================== 内容扩展与网络搜索 ====================

def expand_content_with_search(case_data: Dict, disease: str) -> Dict:
    """
    通过网络搜索扩展内容，补充讨论问题和最新研究
    
    Args:
        case_data: 病例数据
        disease: 疾病名称
    
    Returns:
        扩展后的数据字典
    """
    # 这里预留网络搜索接口
    # 实际使用时可以集成搜索引擎 API 或医学数据库 API
    
    expanded = {
        '讨论问题': {
            '事实性问题': [],
            '分析性问题': [],
            '评价性问题': []
        },
        '参考文献': {
            '指南共识': [],
            '重要文献': [],
            '教材专著': []
        }
    }
    
    # 根据疾病类型生成基础讨论问题模板
    # 实际应用中可以通过网络搜索获取最新研究
    
    return expanded


# ==================== 主生成函数 ====================

def generate_teaching_rounds_plan(jiao_an_data: Dict, output_path: str) -> str:
    """
    生成教学查房教案 docx 文件
    
    Args:
        jiao_an_data: 教案数据字典
        output_path: 输出文件路径
    
    Returns:
        输出文件路径
    """
    doc = Document()
    
    # 设置页面格式
    set_page_format(doc)
    
    # 1. 封面页
    add_cover_page(doc, jiao_an_data.get('基本信息', {}))
    
    # 2. 教学基本信息
    add_basic_info(doc, jiao_an_data)
    
    # 3. 教学目标
    add_teaching_objectives(doc, jiao_an_data)
    
    # 4. 病例摘要
    add_case_summary(doc, jiao_an_data)
    
    # 5. 教学流程与时间安排
    add_teaching_flow(doc, jiao_an_data)
    
    # 6. 教学重点与难点
    add_key_points(doc, jiao_an_data)
    
    # 7. 讨论问题设计
    add_discussion_questions(doc, jiao_an_data)
    
    # 8. 自学任务与参考文献
    add_self_study(doc, jiao_an_data)
    
    # 9. 教学反思与评价
    add_reflection(doc, jiao_an_data)
    
    # 保存文件
    doc.save(output_path)
    
    # 统计字数
    word_count = sum(len(paragraph.text) for paragraph in doc.paragraphs)
    print(f'教案生成成功！')
    print(f'文件路径：{output_path}')
    print(f'总字数：约{word_count}字')
    
    return output_path


if __name__ == '__main__':
    # 测试示例数据
    test_data = {
        '基本信息': {
            '教学查房标题': '肺炎教学查房',
            '专业基地': '内科专业基地',
            '科室': '呼吸与危重症医学科',
            '指导医师': '张三，主任医师',
            '职称': '主任医师',
            '教学日期': '2026 年 3 月 5 日',
            '参与住院医师': '李四（二年级）、王五（一年级）等 8 人',
            '教学时长': '90 分钟'
        },
        '教学目标': {
            '知识目标': [
                '掌握肺炎的临床表现、诊断标准和治疗原则',
                '熟悉社区获得性肺炎与医院获得性肺炎的异同点',
                '了解重症肺炎的评估标准及最新的抗感染治疗指南'
            ],
            '能力目标': [
                '能够系统采集和汇报肺炎患者的病史',
                '能够正确进行胸部体格检查并识别肺部阳性体征',
                '能够根据临床资料和辅助检查结果进行分析，提出诊断和鉴别诊断意见',
                '能够参与制定抗感染治疗方案并合理调整',
                '能够向患者进行疾病健康宣教和用药指导'
            ],
            '素养目标': [
                '体现以患者为中心的医疗服务理念，尊重患者隐私，注重人文关怀',
                '培养循证医学思维和批判性分析能力',
                '增强团队协作意识和医患沟通能力',
                '树立依法执业和医疗安全意识'
            ]
        },
        '病例摘要': {
            '患者基本信息': '男性，65 岁，退休教师',
            '主诉': '咳嗽、咳痰、发热 5 天',
            '现病史': '患者 5 天前受凉后出现咳嗽，咳白色粘痰，伴发热，体温最高 38.5℃，无畏寒、寒战。近 2 天出现胸闷、气促，活动后加重。自行口服"感冒药"效果不佳，遂来我院就诊。病程中患者食欲减退，睡眠欠佳，大小便正常，体重无明显变化。',
            '既往史': '高血压病史 10 年，规律服用降压药，血压控制良好。否认糖尿病、冠心病病史。否认结核、肝炎等传染病史。否认手术、输血史。对"青霉素"过敏。',
            '个人史': '吸烟 30 年，每日 20 支，未戒烟。偶有饮酒。退休教师，长期居住在市区。',
            '家族史': '父亲因"脑卒中"去世，母亲健在。否认家族性遗传性疾病史。',
            '体格检查': 'T 38.2℃，P 95 次/分，R 22 次/分，BP 125/80mmHg。精神尚可，自动体位。口唇轻度发绀，咽充血，扁桃体不大。胸廓对称，呼吸运动度减弱，右下肺叩诊浊音，听诊呼吸音减低，可闻及湿性啰音。心界不大，心律齐，各瓣膜区未闻及杂音。腹平软，肝脾肋下未及。双下肢无水肿。',
            '辅助检查': '（2024-03-15）血常规：WBC 12.5×10⁹/L，N 85%，L 12%；血生化：ALT 35U/L，AST 40U/L，BUN 6.5mmol/L，Cr 85μmol/L；CRP 85mg/L；血气分析：pH 7.42，PaO₂ 65mmHg，PaCO₂ 32mmHg；胸部 CT：右下肺大片状实变影，可见支气管充气征；痰培养：待回报。',
            '初步诊断': '1.社区获得性肺炎（右下肺）；2.高血压病 2 级（极高危）',
            '诊疗经过': '入院后给予吸氧、抗感染（头孢曲松 + 阿奇霉素）、祛痰、退热等对症支持治疗。治疗 3 天后体温恢复正常，咳嗽、气促症状明显好转。复查血常规：WBC 8.5×10⁹/L，N 70%。继续巩固治疗 5 天后出院。'
        },
        '教学重点': [
            '社区获得性肺炎的诊断标准（临床症状 + 影像学 + 实验室检查）',
            'CAP 的抗感染治疗原则及药物选择',
            '重症肺炎的早期识别和评估',
            '病情评估后的治疗方案调整'
        ],
        '教学难点': [
            '肺炎的鉴别诊断思路（与肺结核、肺癌、肺栓塞等疾病的鉴别）',
            '耐药菌肺炎的抗生素选择策略',
            '老年肺炎的个体化治疗决策',
            '抗感染治疗的疗程和停药指征'
        ],
        '讨论问题': {
            '事实性问题': [
                {'问题': '该患者的主要临床表现有哪些？', '参考答案': '咳嗽、咳白色粘痰、发热（最高 38.5℃）、胸闷、气促；右下肺湿性啰音；WBC 升高，CRP 升高；胸部 CT 示右下肺大片状实变影。'},
                {'问题': '该患者的胸部 CT 有什么特点？', '参考答案': '右下肺大片状实变影，可见支气管充气征，这是大叶性肺炎的典型表现。'},
                {'问题': '该患者的血常规和 CRP 有什么异常？', '参考答案': 'WBC 12.5×10⁹/L（升高），N 85%（中性粒细胞比例升高），CRP 85mg/L（明显升高），提示细菌感染。'}
            ],
            '分析性问题': [
                {'问题': '该患者肺炎诊断成立的依据是什么？', '参考答案': '①临床症状：咳嗽、咳痰、发热；②体征：右下肺湿性啰音；③实验室检查：WBC、CRP 升高；④影像学：胸部 CT 示右下肺大片状实变影。符合 CAP 诊断标准。'},
                {'问题': '该患者需要与哪些疾病进行鉴别诊断？', '参考答案': '肺结核、肺癌、肺栓塞、急性支气管炎、慢性阻塞性肺疾病急性加重等。'},
                {'问题': '该患者的 CURB-65 评分是多少？属于哪个严重程度等级？', '参考答案': 'CURB-65 评分：年龄≥65 岁（1 分），尿素氮>7mmol/L（0 分），呼吸频率≥30 次/分（0 分），血压（0 分），意识障碍（0 分）。总分 1 分，属于低危组，可考虑门诊治疗。'}
            ],
            '评价性问题': [
                {'问题': '如果该患者治疗效果不佳，需要考虑哪些可能原因？', '参考答案': '①病原体耐药；②非典型病原体感染；③并发症（如脓胸）；④诊断错误；⑤宿主因素（如免疫功能低下）。'},
                {'问题': '该患者出院后需要哪些健康宣教和随访建议？', '参考答案': '①戒烟指导；②疫苗接种（流感疫苗、肺炎球菌疫苗）；③定期复查胸部影像；④出现发热、咳嗽加重及时就诊；⑤加强营养和锻炼，提高免疫力。'},
                {'问题': '根据最新的 CAP 指南，该患者的抗感染治疗方案应如何优化？', '参考答案': '参考 2023 版 CAP 指南，对于门诊低危患者，首选阿莫西林或多西环素；对于有基础疾病者，可选用呼吸喹诺酮类或β-内酰胺类 + 大环内酯类联合治疗。'}
            ]
        },
        '自学任务': [
            '阅读《社区获得性肺炎诊断和治疗指南（2023 版）》，了解 CAP 的最新诊疗推荐',
            '复习肺炎的影像学表现，掌握不同类型肺炎的影像学特点',
            '查阅文献，了解耐药肺炎链球菌（DRSP）的流行现状及治疗策略',
            '思考：如果该患者合并肾功能不全，抗感染药物应如何调整？'
        ],
        '参考文献': {
            '指南共识': [
                '中华医学会呼吸病学分会。社区获得性肺炎诊断和治疗指南（2023 版）.中华结核和呼吸杂志，2023, 46(12): 1-24',
                'Metlay JP, Waterer GW, Long AC, et al. Diagnosis and Treatment of Adults with Community-acquired Pneumonia. Am J Respir Crit Care Med, 2019, 200(7): e45-e67'
            ],
            '重要文献': [
                'Lee JH, Kim J, Kim K, et al. A Prospective Observational Study of Community-acquired Pneumonia in Korean Adults. J Korean Med Sci, 2022, 37(8): e62',
                '张伟，等。社区获得性肺炎病原学分布及耐药性分析。中华医院感染学杂志，2024, 34(2): 245-250'
            ],
            '教材专著': [
                '王辰。呼吸病学。第 3 版。北京：人民卫生出版社，2022',
                '葛均波，徐永健。内科学。第 9 版。北京：人民卫生出版社，2018'
            ]
        }
    }
    
    # 生成教案
    output_file = r'D:\ww\onedrive\workspace\测试教学查房教案.docx'
    generate_teaching_rounds_plan(test_data, output_file)
