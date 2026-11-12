#!/usr/bin/env python3
"""Behavioral regressions for requirement sheets and body-only targets."""
from pathlib import Path
import tempfile
from docx import Document
from docx.shared import Pt
import thesis_format_from_sample as fmt


def main():
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp)
        sample = Document()
        sample.add_paragraph('摘 要')
        sample.add_paragraph('内容字体为四号宋体，1.5倍行距。摘要约300—500字。')
        sample.add_paragraph('一、示例')
        sample.add_paragraph('内容为小四号宋体。文字的行间距为1.5倍行距，段间距为0。')
        table = sample.add_table(rows=3, cols=2)
        for row, text in zip(table.rows, ['学生姓名：', '学号：', '指导教师：']):
            row.cells[0].text = text
        path = root / 'sample.docx'
        sample.save(path)
        profile = fmt.update_profile_from_observations(fmt.default_profile(), path)
        assert profile['roles']['body']['size_pt'] == 12
        assert profile['roles']['body']['line_spacing'] == 1.5
        assert not any(role.startswith('table_') for role in profile['enabled_roles'])
        target = Document()
        title = target.add_paragraph('某企业营销策略研究')
        title.runs[0].font.size = Pt(20)
        target.add_paragraph('一、绪论')
        target.add_paragraph('（一）研究背景')
        target.add_paragraph('本段用于验证正文在没有摘要时仍能正常应用格式。')
        before = [p.text for p in target.paragraphs]
        source, output = root / 'target.docx', root / 'output.docx'
        target.save(source)
        assert fmt.target_format_start_index(target) == 1
        count = fmt.apply_profile(source, output, profile, True, 'role', 'test', 100)
        result = Document(output)
        assert count > 0
        assert [p.text for p in result.paragraphs] == before
        assert result.paragraphs[0].runs[0].font.size.pt == 20
        assert result.paragraphs[3].style.name == 'ThesisBody'
        assert result.paragraphs[3].runs[0].font.size.pt == 12
        unknown = Document()
        unknown.add_paragraph('没有可识别结构的内容')
        try:
            fmt.target_format_start_index(unknown)
        except ValueError:
            pass
        else:
            raise AssertionError('unrecognized start must fail')
    conflicting = Document()
    conflicting.add_paragraph('正文为小四号宋体')
    conflicting.add_paragraph('正文为四号宋体')
    profile = fmt.default_profile()
    fmt.apply_explicit_body_requirements(profile, conflicting)
    assert any(issue['code'] == 'conflicting_body_requirements' for issue in fmt.lint_profile(profile))
    print('Requirements and body-only regressions passed')


if __name__ == '__main__':
    main()
