<p align="center"><img src="https://massgrave.dev/img/logo_small.png" alt="MAS Logo"></p>

<div align="right">
  <a href="README-zh_TW.md">🇹🇼 繁體中文</a> | <strong>🇨🇳 简体中文</strong> | <a href="README.md">🇺🇸 English</a>
</div>

<h1 align="center">Microsoft  Activation  Scripts (MAS)</h1>

<p align="center">开源的 Windows 和 Office 激活器，支持 HWID、Ohook、TSforge、KMS38 和在线 KMS 激活方法，以及高级故障排除功能。</p>

<hr>
  
## 如何激活 Windows / Office / 扩展更新 (ESU)？

### 方法 1 - PowerShell ❤️

1. **打开 PowerShell**  
   点击**开始菜单**，输入 `PowerShell`，然后打开它。

2. **复制并粘贴下面的代码，然后按回车键。**  
   - 适用于 **Windows 8, 10, 11**: 📌
     ```
     irm https://get.activated.win | iex
     ```
   - 适用于 **Windows 7** 及更高版本:
     ```
     iex ((New-Object Net.WebClient).DownloadString('https://get.activated.win'))
     ```

<details>

<summary>脚本无法启动❓点击这里查看信息。</summary>

---

- 如果上述方法被阻止（被 ISP/DNS 阻止），请尝试这个（需要**更新的 Windows 10 或 11**）：
  ```
  iex (curl.exe -s --doh-url https://1.1.1.1/dns-query https://get.activated.win | Out-String)
  ```
- 如果失败或您使用的是较旧版本的 Windows，请使用下面列出的方法 2。

---

</details>

3. 激活菜单将出现。**选择绿色高亮的选项**来激活 Windows 或 Office。

4. **完成！**

---

### 方法 2 - 传统方法（Windows Vista 及更高版本）

<details>
  <summary>点击这里查看</summary>
  
1.   使用以下链接之一下载文件：  
`https://github.com/massgravel/Microsoft-Activation-Scripts/archive/refs/heads/master.zip`  
或  
`https://git.activated.win/massgrave/Microsoft-Activation-Scripts/archive/master.zip`
2.   右键点击下载的 zip 文件并解压缩。
3.   在解压缩的文件夹中，找到名为 `All-In-One-Version` 的文件夹。
4.   运行名为 `MAS_AIO.cmd` 的文件。
5.   您将看到激活选项。按照屏幕上的说明操作。
6.   就是这样。

</details>

---

> [!TIP]
> - 一些 ISP/DNS 会阻止对我们域名的访问。您可以通过在浏览器中启用 [DNS-over-HTTPS (DoH)](https://developers.cloudflare.com/1.1.1.1/encryption/dns-over-https/encrypted-dns-browsers/) 来绕过这个问题。  
> - **遇到问题**❓访问我们的[故障排除页面](https://massgrave.dev/troubleshoot)或在 [GitHub](https://github.com/massgravel/Microsoft-Activation-Scripts/issues) 上提出问题。

---

- 要激活其他产品，如 **macOS 版 Office、Visual Studio、RDS CAL 和 Windows XP**，请查看[这里](https://massgrave.dev/unsupported_products_activation)。
- 要以无人值守模式运行脚本，请查看[这里](https://massgrave.dev/command_line_switches)。

---

> [!NOTE]
>
> - PowerShell 中的 IRM 命令从指定的 URL 下载脚本，IEX 命令执行它。
> - 在执行命令之前，请务必仔细检查 URL，如果手动下载文件，请验证来源。
> - 要谨慎，因为有些人会通过在 IRM 命令中使用不同的 URL 来传播伪装成 MAS 的恶意软件。

---

```
最新版本: 3.5
发布日期: 2025年8月10日
```

### [故障排除 / 帮助](https://massgrave.dev/troubleshoot)
### [下载原版 Windows 和 Office](https://massgrave.dev/genuine-installation-media)
### 主页 - [https://massgrave.dev/](https://massgrave.dev/)

<div align="center">
  
[![1.1]][1]
[![1.2]][2]
[![1.3]][3]

</div>

<div align="center">
  
[![1.4]][4]
[![1.5]][5]
[![1.6]][6]
[![1.7]][7]

</div>

[1.1]: https://massgrave.dev/img/logo_github.png (GitHub)
[1.2]: https://massgrave.dev/img/logo_azuredevops.png (AzureDevOps)
[1.3]: https://massgrave.dev/img/logo_gitea.png (自托管 Git)

[1.4]: https://massgrave.dev/img/logo_discord.png (无需注册即可与我们聊天)
[1.5]: https://massgrave.dev/img/logo_reddit.png (Reddit)
[1.6]: https://massgrave.dev/img/logo_bluesky.png (Bluesky)
[1.7]: https://massgrave.dev/img/logo_x.png (Twitter)

[1]: https://github.com/massgravel/Microsoft-Activation-Scripts
[2]: https://dev.azure.com/massgrave/_git/Microsoft-Activation-Scripts
[3]: https://git.activated.win/massgrave/Microsoft-Activation-Scripts
[4]: https://discord.gg/j2yFsV5ZVC
[5]: https://www.reddit.com/r/MAS_Activator
[6]: https://bsky.app/profile/massgrave.dev
[7]: https://twitter.com/massgravel

---

<p align="center">用爱制作 ❤️</p>