#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
教学查房教案技能打包脚本
将技能打包成 .skill 文件（zip 格式）
"""

import os
import sys
import zipfile
from pathlib import Path


def validate_skill(skill_dir):
    """
    验证技能目录结构
    
    Args:
        skill_dir: 技能目录路径
    
    Returns:
        (bool, list): 验证是否通过，错误信息列表
    """
    errors = []
    
    # 检查 SKILL.md 是否存在
    skill_md = os.path.join(skill_dir, 'SKILL.md')
    if not os.path.exists(skill_md):
        errors.append("缺少 SKILL.md 文件")
        return False, errors
    
    # 检查 SKILL.md 前格式
    with open(skill_md, 'r', encoding='utf-8') as f:
        content = f.read()
    
    if not content.strip().startswith('---'):
        errors.append("SKILL.md 缺少 YAML 前格式（---）")
    
    if 'name:' not in content:
        errors.append("SKILL.md 缺少 name 字段")
    
    if 'description:' not in content:
        errors.append("SKILL.md 缺少 description 字段")
    
    # 检查 scripts 目录
    scripts_dir = os.path.join(skill_dir, 'scripts')
    if os.path.exists(scripts_dir):
        py_files = [f for f in os.listdir(scripts_dir) if f.endswith('.py')]
        if not py_files:
            errors.append("scripts 目录中没有 Python 文件")
    
    # 检查 references 目录（可选）
    references_dir = os.path.join(skill_dir, 'references')
    if os.path.exists(references_dir):
        md_files = [f for f in os.listdir(references_dir) if f.endswith('.md')]
        if not md_files:
            print("提示：references 目录中没有 Markdown 文件（可选）")
    
    return len(errors) == 0, errors


def package_skill(skill_dir, output_dir=None):
    """
    打包技能为 .skill 文件
    
    Args:
        skill_dir: 技能目录路径
        output_dir: 输出目录（可选，默认为技能目录的父目录）
    
    Returns:
        str: 生成的 .skill 文件路径
    """
    skill_dir = os.path.abspath(skill_dir)
    skill_name = os.path.basename(skill_dir)
    
    # 验证技能
    print(f'正在验证技能：{skill_name}')
    is_valid, errors = validate_skill(skill_dir)
    
    if not is_valid:
        print('验证失败，错误如下：')
        for error in errors:
            print(f'  - {error}')
        return None
    
    print('验证通过')
    
    # 确定输出目录
    if output_dir is None:
        output_dir = os.path.dirname(skill_dir)
    else:
        output_dir = os.path.abspath(output_dir)
        os.makedirs(output_dir, exist_ok=True)
    
    # 生成 .skill 文件路径
    skill_file = os.path.join(output_dir, f'{skill_name}.skill')
    
    # 创建 zip 文件
    print(f'正在打包技能到：{skill_file}')
    
    with zipfile.ZipFile(skill_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(skill_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, os.path.dirname(skill_dir))
                zipf.write(file_path, arcname)
    
    # 显示文件信息
    file_size = os.path.getsize(skill_file)
    print(f'打包完成！文件大小：{file_size / 1024:.2f} KB')
    
    return skill_file


if __name__ == '__main__':
    # 默认打包当前目录的 teaching-rounds-jiaoan 技能
    if len(sys.argv) > 1:
        skill_dir = sys.argv[1]
    else:
        # 获取技能根目录（scripts 的父目录）
        skill_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    result = package_skill(skill_dir, output_dir)
    
    if result:
        print(f'\n技能打包成功：{result}')
    else:
        print('\n技能打包失败')
        sys.exit(1)
