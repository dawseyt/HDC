## 2024-05-18 - Cryptographically Secure Random Number Generation
**Vulnerability:** Weak random number generation (`Get-Random`) used for creating new Active Directory user passwords.
**Learning:** `Get-Random` in PowerShell 5.1 relies on the non-cryptographically secure `System.Random` class, meaning that the generated passwords could theoretically be predicted by a malicious actor who observes the PRNG's outputs.
**Prevention:** Use `System.Security.Cryptography.RandomNumberGenerator::Create()` for generating strong, unpredictable random numbers for security-critical tasks like password generation, tokens, or encryption keys.
## 2024-05-18 - Command Injection via AppX Repair/Uninstall String Interpolation
**Vulnerability:** Untrusted input (`$targetId`) from remote systems was insufficiently sanitized (`$targetId.Replace("'", "''")`) and then injected directly into a PowerShell command string executed by a child `powershell.exe` process, allowing for arbitrary command execution.
**Learning:** Manual string replacement for sanitizing input for PowerShell execution is flawed. Edge-case escape sequence injections can bypass such measures. The string was interpolated into `$psCmd = "Remove-AppxPackage -Package '$safeTargetId' -AllUsers"`.
**Prevention:** Treat untrusted input strictly as data. Base64 encode the user-controlled variable separately, then construct a wrapper command that decodes it and pass the entire wrapper command to `powershell.exe` using `-EncodedCommand`.
