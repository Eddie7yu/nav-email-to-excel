# IMAP 服务商与核验方法

当用户不知道 IMAP 服务器或认证方式时，部署 AI 必须先阅读本页。公司邮箱后缀不等于
邮箱服务商：即使地址使用 `user@example.invalid` 这类企业自有域名，背后也可能由
腾讯、阿里、Google、Microsoft 或自建系统托管，不能仅凭地址推断服务器。

本页信息最后核验于 2026-07-20。服务商地址和安全策略可能变化，每次部署仍须以用户
现有客户端、邮箱管理员或服务商当前官方文档为准。

## 核验顺序

1. 优先读取用户现有 Outlook、Foxmail、手机邮箱中的“收件服务器/IMAP”设置；不得
   读取或展示已保存密码。
2. 其次让用户从邮箱后台或单位管理员处确认 IMAP 主机、端口、是否开启第三方客户端，
   以及应使用授权码、应用专用密码还是 OAuth。
3. 再查服务商当前官方文档。若 AI 不能联网或官方资料无法确认，应停下来请用户或 IT
   提供设置，不能编造答案。
4. DNS MX 记录只能辅助识别可能的托管服务商。MX 是收信投递地址，不是 IMAP 地址；
   不得把 MX 主机直接填入配置，也不得未经确认就猜 `imap.<公司域名>`。
5. 当前运行时使用 IMAP over SSL，通常是 993 端口。只支持密码、授权码或应用专用
   密码登录；仅支持 OAuth2/Modern Auth 的账号不能部署。
6. 地址确认后，先在本机进行只读连接和预览验证。授权码只能由用户通过 `secret set`
   隐藏输入，不得进入聊天、命令行参数、文档或 Git。

## 常见服务商速查

下表只提供部署候选值，不构成对某个企业邮箱后端的认定。

| 服务商 | 常见 IMAP SSL 主机 | 端口 | 认证与当前支持状态 |
| --- | --- | ---: | --- |
| QQ 邮箱 | `imap.qq.com` | 993 | 开启 IMAP 后使用授权码；支持 |
| 腾讯企业邮 | `imap.exmail.qq.com` | 993 | 需管理员允许第三方客户端，并以单位当前要求的客户端密码或授权码为准；满足密码式认证时支持 |
| 网易 163 | `imap.163.com` | 993 | 使用授权码；运行时会在登录后自动发送 IMAP ID；支持 |
| 网易 126 | `imap.126.com` | 993 | 使用授权码；运行时会在登录后自动发送 IMAP ID；支持 |
| 网易 yeah.net | `imap.yeah.net` | 993 | 使用授权码；运行时会在登录后自动发送 IMAP ID；支持 |
| 网易企业邮 | 以管理员提供的网易 IMAP 主机为准 | 993 | 对以 `163.com`、`126.com` 或 `yeah.net` 结尾的主机自动发送 IMAP ID；其他定制主机须先核验；条件支持 |
| 阿里企业邮箱（大陆通用） | `imap.qiye.aliyun.com` | 993 | 需管理员开启 IMAP，通常使用第三方客户端安全密码；支持 |
| 阿里企业邮箱（香港节点） | `imaphk.qiye.aliyun.com` | 993 | 还可能使用旧版或企业自定义主机；以管理员和官方配置为准；支持密码式认证 |
| Gmail / Google Workspace | `imap.gmail.com` | 993 | Google 优先要求 OAuth；仅当账号允许应用专用密码时条件支持，普通登录密码不可用 |
| Outlook.com / Microsoft 365 | `outlook.office365.com` | 993 | Microsoft 当前要求 OAuth2/Modern Auth；本运行时暂不支持 |
| 自建或其他企业邮箱 | 由单位管理员提供 | 通常 993 | 只有在支持 IMAP SSL 和密码式认证时才支持；不得依据邮箱后缀猜测 |

## 官方核验入口

- [QQ 邮箱帮助中心](https://service.mail.qq.com/)：搜索“POP/IMAP 和 Exchange 服务的设置方法”。
- [腾讯企业邮箱](https://exmail.qq.com/)：企业账号还须向本单位管理员确认客户端访问策略。
- [网易邮箱帮助中心](https://help.mail.163.com/)：分别核对账号所属的 163、126、yeah.net
  或企业邮箱说明。
- [阿里邮箱 IMAP/POP/SMTP 地址与端口](https://help.aliyun.com/zh/document_detail/36576.html)：
  官方同时列出大陆通用、香港、旧版和自定义域名场景。
- [Google Workspace 第三方客户端设置](https://support.google.com/a/answer/9003945?hl=zh-Hans)
  和[应用专用密码条件](https://support.google.com/accounts/answer/185833?hl=zh-Hans)。
- [Microsoft Outlook.com IMAP 设置](https://support.microsoft.com/en-US/Outlook/pop-imap-and-smtp-settings-for-outlook-com)：
  核对服务器地址和 OAuth2/Modern Auth 要求。

如果服务商官方资料与本表冲突，应采用官方当前资料，并把差异作为公开 Skill 的文档
问题报告；不要为了完成部署而继续使用未经证实的旧地址或认证方式。
