# SDU Comprehensive Evaluation Skill

这是一个用于整理、审核山东大学低空科学与工程学院本科生综合评价材料的 Codex skill 导出包。

## 内容

- `skill/sdu-comprehensive-evaluation/`：Codex skill 本体、规则摘要和检查脚本。
- `project-files/综测模板/`：本 skill 依赖的通知、评分办法、竞赛分类附件、Excel 模板和示例材料。

## 使用方式

1. 将 `skill/sdu-comprehensive-evaluation` 放入 Codex 可发现的 skills 目录，或在项目中保留该目录。
2. 将 `project-files/综测模板` 放到待整理项目工作区。
3. 对学生材料运行检查脚本：

```bash
python3 skill/sdu-comprehensive-evaluation/scripts/inspect_submission.py <学生材料文件夹>
```

## 隐私说明

本仓库不包含学生个人整理结果、学号、个人证明图片或本机临时文件。请提交贡献时也不要上传真实学生个人材料。

## 贡献

欢迎 fork 后通过 Pull Request 改进规则摘要、检查脚本和流程说明。公开仓库默认允许 fork；外部贡献请通过 PR 提交。
