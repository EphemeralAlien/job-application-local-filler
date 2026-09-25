# Security policy

请勿将以下文件或内容提交到仓库、发送给 AI 或上传到网盘：

- `profile.ini` 与任何 `.vault` 文件；
- 身份证、护照、银行账户、验证码、密码；
- 含真实联系方式、地址、家庭成员资料的完整网申表。

发现敏感信息误入暂存区时，先执行 `git restore --staged <文件>`，确认文件已经在 `.gitignore` 中，再继续提交。
