<p align="center"><img src="https://massgrave.dev/img/logo_small.png" alt="MAS Logo"></p>

<h1 align="center">Microsoft  Activation  Scripts (MAS)</h1>

<p align="center">Open-source Windows and Office activator featuring HWID, Ohook, TSforge, KMS38, and Online KMS activation methods, along with advanced troubleshooting.</p>

<p align="center">
  <img src="https://img.shields.io/github/downloads/massgravel/Microsoft-Activation-Scripts/total?style=for-the-badge&color=green" alt="Downloads">
  <img src="https://img.shields.io/github/last-commit/massgravel/Microsoft-Activation-Scripts?style=for-the-badge&color=blue" alt="Last Commit">
  <img src="https://img.shields.io/github/license/massgravel/Microsoft-Activation-Scripts?style=for-the-badge&color=orange" alt="License">
</p>

<hr>

## ✨ Key Features

- **🚀 Multiple Activation Methods**: HWID, Ohook, TSforge, KMS38, and Online KMS
- **🛡️ Safe & Open Source**: Transparent code, no hidden malware or backdoors
- **💻 Wide Compatibility**: Supports Windows 7/8/8.1/10/11 and Server editions
- **📱 Office Support**: Activates Office 2010-2021 and Office 365
- **🔧 Advanced Troubleshooting**: Built-in diagnostic and repair tools
- **⚡ Easy to Use**: Simple GUI and PowerShell one-liner execution
- **🔄 Automatic Updates**: Always uses the latest activation methods

<hr>
  
## How to Activate Windows / Office / Extended Updates (ESU)?

### Method 1 - PowerShell ❤️

1. **Open PowerShell**  
   Click the **Start Menu**, type `PowerShell`, then open it.

2. **Copy and paste the code below, then press enter.**  
   - For **Windows 8, 10, 11**: 📌
     ```
     irm https://get.activated.win | iex
     ```
   - For **Windows 7** and later:
     ```
     iex ((New-Object Net.WebClient).DownloadString('https://get.activated.win'))
     ```

<details>

<summary>Script not launching❓Click here for info.</summary>

---

- If the above is blocked (by ISP/DNS), try this (needs **updated Windows 10 or 11**):
  ```
  iex (curl.exe -s --doh-url https://1.1.1.1/dns-query https://get.activated.win | Out-String)
  ```
- If that fails or you have an older Windows, use the below-listed Method 2.

---

</details>

3. The activation menu will appear. **Choose the green-highlighted options** to activate Windows or Office.

4. **Done!**

---

### Method 2 - Traditional (Windows Vista and later)

<details>
  <summary>Click here to view</summary>
  
1.   Download the file using one of the links below:  
`https://github.com/massgravel/Microsoft-Activation-Scripts/archive/refs/heads/master.zip`  
or  
`https://git.activated.win/massgrave/Microsoft-Activation-Scripts/archive/master.zip`
2.   Right-click on the downloaded zip file and extract it.
3.   In the extracted folder, navigate to `MAS` → `All-In-One-Version-KL`
4.   Run the file named `MAS_AIO.cmd`.
5.   You will see the activation options. Follow the on-screen instructions.
6.   That's all.

</details>

---

> [!TIP]
> - Some ISPs/DNS block access to our domains. You can bypass this by enabling [DNS-over-HTTPS (DoH)](https://developers.cloudflare.com/1.1.1.1/encryption/dns-over-https/encrypted-dns-browsers/) in your browser.  
> - **Having trouble**❓Visit our [troubleshooting page](https://massgrave.dev/troubleshoot) or raise an issue on [GitHub](https://github.com/massgravel/Microsoft-Activation-Scripts/issues).

---

- To activate additional products such as **Office for macOS, Visual Studio, RDS CALs, and Windows XP**, check [here](https://massgrave.dev/unsupported_products_activation).
- To run the scripts in unattended mode, check [here](https://massgrave.dev/command_line_switches).

---

> [!NOTE]
>
> - The IRM command in PowerShell downloads a script from a specified URL, and the IEX command executes it.
> - Always double-check the URL before executing the command and verify the source if manually downloading files.
> - Be cautious, as some spread malware disguised as MAS by using different URLs in the IRM command.

---

## ⚠️ Important Safety Information

- **Official Sources Only**: Always download MAS from official repositories to avoid malware
- **Verify Commands**: Double-check PowerShell commands before execution
- **Antivirus Warnings**: Some antivirus software may flag activation tools - this is normal for this type of software
- **Backup Recommended**: Create a system backup before making activation changes
- **Educational Purpose**: This tool is for educational and testing purposes

> [!CAUTION]
> Beware of fake MAS versions distributed with malware. Only use official sources listed in this repository.

---

## 🔧 Development & Contributing

Interested in contributing to MAS? We welcome improvements that enhance reliability and user experience!

**Quick Guidelines:**
- Use modern batch scripting practices (e.g., `timeout` instead of `ping` for delays)
- Test changes on multiple Windows versions (10, 11, Server)
- Maintain backward compatibility
- Follow existing code formatting and conventions

**Report Issues:** [GitHub Issues](https://github.com/massgravel/Microsoft-Activation-Scripts/issues) | **Get Help:** [Discord](https://discord.gg/j2yFsV5ZVC)

---

## 📋 Supported Activation Methods

| Method | Supported Products | Duration | Notes |
|--------|-------------------|----------|--------|
| **HWID** | Windows 10-11 | Permanent | Hardware-based activation |
| **Ohook** | Office | Permanent | Office hook method |
| **TSforge** | Windows / ESU / Office | Permanent | Enhanced activation |
| **KMS38** | Windows 10-11-Server | Until 2038 | Extended KMS activation |
| **Online KMS** | Windows / Office | 180 Days | Renewable with task |

For detailed compatibility, visit: [https://massgrave.dev/chart](https://massgrave.dev/chart)

---

## ❓ Frequently Asked Questions

<details>
<summary><b>Is MAS safe to use?</b></summary>

Yes, MAS is completely open-source and safe. The code is transparent and can be reviewed by anyone. Some antivirus programs may flag it due to the nature of activation tools, but this is a false positive.
</details>

<details>
<summary><b>Will this harm my computer?</b></summary>

No, MAS only modifies Windows licensing components and doesn't install any malware or unwanted software. It's used by millions of people worldwide.
</details>

<details>
<summary><b>Can I get banned or in trouble for using this?</b></summary>

MAS is for educational and testing purposes. The legal implications vary by jurisdiction and use case. Please review your local laws and Microsoft's terms of service.
</details>

<details>
<summary><b>Which activation method should I use?</b></summary>

- **HWID**: Best for Windows 10/11 (permanent activation)
- **KMS38**: Good for Windows 10/11/Server (activates until 2038)
- **Ohook**: Perfect for Office (permanent activation)
- **Online KMS**: Universal method for Windows/Office (180-day renewable)
</details>

<details>
<summary><b>My antivirus is blocking MAS, what should I do?</b></summary>

This is normal for activation tools. You can temporarily disable your antivirus or add MAS to the exclusion list. Always download from official sources to ensure safety.
</details>

---

```
Latest Version: 3.7
Release date: 11-Sep-2025
```

### [Troubleshooting / Help](https://massgrave.dev/troubleshoot)
### [Download Original Windows & Office](https://massgrave.dev/genuine-installation-media)
### Homepage - [https://massgrave.dev/](https://massgrave.dev/)

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
[1.3]: https://massgrave.dev/img/logo_gitea.png (Self-hosted Git)

[1.4]: https://massgrave.dev/img/logo_discord.png (Chat with us without signup)
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

<p align="center">Made with Love ❤️</p>
