# 周敏晖各论课申报材料

上海商学院一流本科课程（各论课）申报相关文档与生成脚本。

## 文件清单

| 文件 | 说明 |
|------|------|
| `周敏晖-各论课申报书-从组织路线到组织再造.docx` | 完整申报书 |
| `周敏晖-各论课教学大纲.docx` | 课程教学大纲 |
| `周敏晖-各论课申报PPT.pptx` | 申报答辩 PPT（17页） |
| `郝宇青-学术合作与研究贡献总结.docx` | 合作学者郝宇青学术概况与研究总结 |
| `generate_course_application.py` | 申报书生成脚本（可重新运行） |
| `generate_ppt_and_syllabus.py` | PPT 与大纲生成脚本 |
| `generate_haoyuqing_summary.py` | 郝宇青总结 Word 生成脚本 |

## 重新生成

```bash
python3 generate_course_application.py
python3 generate_ppt_and_syllabus.py
python3 generate_haoyuqing_summary.py
```

## 环境依赖

```bash
pip install -r requirements.txt
```

## 仓库说明

本仓库独立于 `gcc` 项目，专门存放周敏晖各论课申报材料。

### 推送到 GitHub

1. 在 GitHub 创建空仓库：https://github.com/new?name=minhuizhou
2. 不要勾选 “Add a README file”
3. 在本目录运行：

```bash
./push_to_github.sh
```

如需 Cloud Agent 继续维护此仓库，请在 Cursor Integrations 中为 `minhuizhou` 授权 GitHub App。
