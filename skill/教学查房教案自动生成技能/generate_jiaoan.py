#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
住院医师规范化培训教学查房教案自动生成

使用方法：
1. 准备病例资料（JSON 格式或 Python 字典）
2. 调用 generate_from_case 函数
3. 生成符合格式要求的 DOCX 文件（8000-12000 字）

默认配置（不可更改）：
- 培训基地名称：十堰市人民医院普外科基地
- 专业基地/科室名称：甲乳血管外科一病区
- 指导医师姓名及职称：王巍（主治医师）
- 参加讨论的住院医师人数：在科住培医师和基地住培医师
- 教学时长：60 分钟

示例：
    from generate_jiaoan import generate_from_case

    case_data = {
        '患者基本信息': '男性，65 岁',
        '主诉': '咳嗽、发热 5 天',
        '现病史': '...',
        ...
    }

    output_path = generate_from_case(case_data, '肺炎教学查房教案.docx')
"""

import sys
import os

# 添加脚本目录到路径
current_dir = os.path.dirname(os.path.abspath(__file__))
scripts_dir = os.path.join(current_dir, 'scripts')
sys.path.insert(0, scripts_dir)

from generate_teaching_rounds_plan import generate_teaching_rounds_plan


# ==================== 默认配置常量（不可更改） ====================

# 培训基地名称
DEFAULT_TRAINING_BASE = "十堰市人民医院普外科基地"

# 专业基地/科室名称
DEFAULT_DEPARTMENT = "甲乳血管外科一病区"

# 指导医师姓名及职称
DEFAULT_INSTRUCTOR = "王巍（主治医师）"

# 参加讨论的住院医师人数
DEFAULT_RESIDENTS = "在科住培医师和基地住培医师"

# 教学时长（分钟）
DEFAULT_TEACHING_DURATION = "60 分钟"


def get_default_basic_info():
    """获取默认的基本信息配置"""
    return {
        '专业基地': DEFAULT_TRAINING_BASE,
        '科室': DEFAULT_DEPARTMENT,
        '指导医师': DEFAULT_INSTRUCTOR,
        '参与住院医师': DEFAULT_RESIDENTS,
        '教学时长': DEFAULT_TEACHING_DURATION
    }


def generate_from_case(case_data: dict, output_path: str, disease_name: str = None) -> str:
    """
    根据病例资料生成教学查房教案

    Args:
        case_data: 病例资料字典，包含：
            - 教学查房标题（可选，默认从疾病名称生成）
            - 病例摘要：患者基本信息、主诉、现病史、体格检查、辅助检查、诊断、治疗经过
            - 其他可选字段：教学目标、教学重点、教学难点、讨论问题等
            注意：基本信息将使用默认配置（培训基地、科室、指导医师等）
        output_path: 输出 DOCX 文件路径
        disease_name: 疾病名称（可选，用于自动补充内容）

    Returns:
        str: 生成的文件路径
    """
    # 获取默认基本信息
    default_info = get_default_basic_info()
    
    # 构建完整教案数据
    jiao_an_data = {
        '基本信息': {
            '教学查房标题': case_data.get('教学查房标题', f'{disease_name or "疾病"}教学查房'),
            '专业基地': default_info['专业基地'],
            '科室': default_info['科室'],
            '指导医师': default_info['指导医师'],
            '职称': '主治医师',
            '教学日期': case_data.get('教学日期', '2026 年 X 月 X 日'),
            '参与住院医师': default_info['参与住院医师'],
            '教学时长': default_info['教学时长']
        },
        '教学目标': case_data.get('教学目标', {
            '知识目标': [
                f'掌握{disease_name or "该疾病"}的临床表现、诊断标准和治疗原则',
                f'熟悉{disease_name or "该疾病"}的鉴别诊断要点',
                f'了解{disease_name or "该疾病"}的最新诊疗进展'
            ],
            '能力目标': [
                '能够系统采集和汇报患者的病史',
                '能够正确进行体格检查并识别阳性体征',
                '能够根据临床资料提出诊断和鉴别诊断意见',
                '能够参与制定个体化诊疗方案'
            ],
            '素养目标': [
                '体现以患者为中心的医疗服务理念',
                '培养循证医学思维和批判性分析能力',
                '增强团队协作意识和医患沟通能力'
            ]
        }),
        '病例摘要': case_data.get('病例摘要', {
            '患者基本信息': case_data.get('患者基本信息', ''),
            '主诉': case_data.get('主诉', ''),
            '现病史': case_data.get('现病史', ''),
            '既往史': case_data.get('既往史', ''),
            '个人史': case_data.get('个人史', ''),
            '家族史': case_data.get('家族史', ''),
            '体格检查': case_data.get('体格检查', ''),
            '辅助检查': case_data.get('辅助检查', ''),
            '初步诊断': case_data.get('初步诊断', ''),
            '诊疗经过': case_data.get('诊疗经过', '')
        }),
        '教学重点': case_data.get('教学重点', [
            f'{disease_name or "该疾病"}的诊断要点',
            f'{disease_name or "该疾病"}的治疗原则',
            '鉴别诊断思路'
        ]),
        '教学难点': case_data.get('教学难点', [
            f'{disease_name or "该疾病"}的鉴别诊断',
            '复杂情况下的治疗决策',
            '最新指南的理解和应用'
        ]),
        '讨论问题': case_data.get('讨论问题', {
            '事实性问题': [
                {'问题': '该患者的主要临床表现有哪些？', '参考答案': '根据病例资料总结'},
                {'问题': '该患者的辅助检查有什么异常？', '参考答案': '根据病例资料总结'}
            ],
            '分析性问题': [
                {'问题': f'该患者诊断为{disease_name or "该疾病"}的依据是什么？', '参考答案': '从症状、体征、辅助检查等方面分析'},
                {'问题': '需要与哪些疾病进行鉴别诊断？', '参考答案': '列出主要鉴别诊断'}
            ],
            '评价性问题': [
                {'问题': '如果治疗效果不佳，需要考虑哪些可能原因？', '参考答案': '从诊断、治疗、宿主因素等方面分析'},
                {'问题': '根据最新指南，治疗方案应如何优化？', '参考答案': '参考最新指南推荐'}
            ]
        }),
        '自学任务': case_data.get('自学任务', [
            f'阅读《{disease_name or "相关"}疾病诊断和治疗指南》，了解最新诊疗推荐',
            '复习相关疾病的影像学表现',
            '查阅文献，了解该领域的最新研究进展'
        ]),
        '参考文献': case_data.get('参考文献', {
            '指南共识': [
                f'中华医学会相关专业分会。{disease_name or "相关"}疾病诊断和治疗指南。中华医学杂志，近年'
            ],
            '重要文献': [
                '相关重要研究论文'
            ],
            '教材专著': [
                '人卫版相关专业教材'
            ]
        })
    }
    
    # 生成教案
    return generate_teaching_rounds_plan(jiao_an_data, output_path)


if __name__ == '__main__':
    # 简单测试
    test_case = {
        '基本信息': {
            '教学查房标题': '肺炎教学查房',
            '专业基地': '内科专业基地',
            '科室': '呼吸与危重症医学科',
            '指导医师': '张三，主任医师',
            '教学日期': '2026 年 3 月 5 日',
            '参与住院医师': '李四等 8 人',
            '教学时长': '90 分钟'
        },
        '病例摘要': {
            '患者基本信息': '男性，65 岁，退休教师',
            '主诉': '咳嗽、咳痰、发热 5 天',
            '现病史': '患者 5 天前受凉后出现咳嗽，咳白色粘痰，伴发热...',
            '体格检查': 'T 38.2℃，右下肺湿性啰音',
            '辅助检查': 'WBC 12.5×10⁹/L，胸部 CT 示右下肺大片状实变影',
            '初步诊断': '社区获得性肺炎',
            '诊疗经过': '给予抗感染、对症支持治疗，好转出院'
        }
    }
    
    output = generate_from_case(test_case, '测试输出.docx', '肺炎')
    print(f'教案生成成功：{output}')
