# 网申本地填表工具

这是一个本地运行的 Word 网申表格填充工具。AI 只处理公开职业信息和岗位适配内容；姓名、证件号码、联系方式、住址与家庭信息仅由本机脚本替换。

## 安装

```powershell
python -m pip install -r requirements.txt
```

## 使用顺序

1. 运行 `prepare` 后，脚本会自动把文档所需变量追加到本机 `profile.ini`；同名变量的已有值永不覆盖。
2. 运行 `vault` 生成加密的 `personal.vault`。
3. 对自己的 Word 模板运行 `prepare`，生成 AI 输入版；再将它与岗位描述交给 AI。要求 AI 填写公开内容和黄色岗位字段，并保留所有 `{{...}}` 占位符。
4. 对 AI 返回的文件执行 `check`，确认黄色岗位字段没有遗漏。
5. 使用 `fill --vault` 在本机生成完整信息表。

完整命令、AI 提示词和安全说明见 [填写指引.md](填写指引.md)。

## 安全边界

`profile.ini`、`.vault`、所有 Word 表格、AI 返回的文件和最终完整信息表可能含敏感资料，已由 `.gitignore` 排除。提交前仍请运行 `git status`，确认没有敏感文件出现在待提交列表。
