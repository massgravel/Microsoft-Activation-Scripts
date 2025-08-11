<p align="center"><img src="https://massgrave.dev/img/logo_small.png" alt="MAS Logo"></p>

<div align="right">
  <strong>🇹🇼 繁體中文</strong> | <a href="README-zh_CN.md">🇨🇳 简体中文</a> | <a href="README.md">🇺🇸 English</a>
</div>

<h1 align="center">Microsoft Activation Scripts (MAS)</h1>

<p align="center">開源的 Windows 和 Office 啟用器，支援 HWID、Ohook、TSforge、KMS38 和線上 KMS 啟用方法，以及進階故障排除功能。</p>

<hr>
  
## 如何啟用 Windows / Office / 延伸更新 (ESU)？

### 方法 1 - PowerShell ❤️

1. **開啟 PowerShell**  
   點擊**開始功能表**，輸入 `PowerShell`，然後開啟它。

2. **複製並貼上下面的程式碼，然後按 Enter 鍵。**  
   - 適用於 **Windows 8, 10, 11**: 📌
     ```
     irm https://get.activated.win | iex
     ```
   - 適用於 **Windows 7** 及更高版本:
     ```
     iex ((New-Object Net.WebClient).DownloadString('https://get.activated.win'))
     ```

<details>

<summary>腳本無法啟動❓點擊這裡查看資訊。</summary>

---

- 如果上述方法被封鎖（被 ISP/DNS 封鎖），請嘗試這個（需要**更新的 Windows 10 或 11**）：
  ```
  iex (curl.exe -s --doh-url https://1.1.1.1/dns-query https://get.activated.win | Out-String)
  ```
- 如果失敗或您使用的是較舊版本的 Windows，請使用下面列出的方法 2。

---

</details>

3. 啟用功能表將出現。**選擇綠色高亮的選項**來啟用 Windows 或 Office。

4. **完成！**

---

### 方法 2 - 傳統方法（Windows Vista 及更高版本）

<details>
  <summary>點擊這裡檢視</summary>
  
1.   使用以下連結之一下載檔案：  
`https://github.com/massgravel/Microsoft-Activation-Scripts/archive/refs/heads/master.zip`  
或  
`https://git.activated.win/massgrave/Microsoft-Activation-Scripts/archive/master.zip`
2.   右鍵點擊下載的 zip 檔案並解壓縮。
3.   在解壓縮的資料夾中，找到名為 `All-In-One-Version` 的資料夾。
4.   執行名為 `MAS_AIO.cmd` 的檔案。
5.   您將看到啟用選項。按照螢幕上的說明操作。
6.   就是這樣。

</details>

---

> [!TIP]
> - 一些 ISP/DNS 會封鎖對我們網域的存取。您可以透過在瀏覽器中啟用 [DNS-over-HTTPS (DoH)](https://developers.cloudflare.com/1.1.1.1/encryption/dns-over-https/encrypted-dns-browsers/) 來繞過這個問題。  
> - **遇到問題**❓造訪我們的[故障排除頁面](https://massgrave.dev/troubleshoot)或在 [GitHub](https://github.com/massgravel/Microsoft-Activation-Scripts/issues) 上提出問題。

---

- 要啟用其他產品，如 **macOS 版 Office、Visual Studio、RDS CAL 和 Windows XP**，請檢視[這裡](https://massgrave.dev/unsupported_products_activation)。
- 要以無人值守模式執行腳本，請檢視[這裡](https://massgrave.dev/command_line_switches)。

---

> [!NOTE]
>
> - PowerShell 中的 IRM 命令從指定的 URL 下載腳本，IEX 命令執行它。
> - 在執行命令之前，請務必仔細檢查 URL，如果手動下載檔案，請驗證來源。
> - 要謹慎，因為有些人會透過在 IRM 命令中使用不同的 URL 來傳播偽裝成 MAS 的惡意軟體。

---

```
最新版本: 3.5
發布日期: 2025年8月10日
```

### [故障排除 / 說明](https://massgrave.dev/troubleshoot)
### [下載原版 Windows 和 Office](https://massgrave.dev/genuine-installation-media)
### 首頁 - [https://massgrave.dev/](https://massgrave.dev/)

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

[1.4]: https://massgrave.dev/img/logo_discord.png (無需註冊即可與我們聊天)
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

<p align="center">用愛製作 ❤️</p>