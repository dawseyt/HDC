## 2024-05-18 - Cryptographically Secure Random Number Generation
**Vulnerability:** Weak random number generation (`Get-Random`) used for creating new Active Directory user passwords.
**Learning:** `Get-Random` in PowerShell 5.1 relies on the non-cryptographically secure `System.Random` class, meaning that the generated passwords could theoretically be predicted by a malicious actor who observes the PRNG's outputs.
**Prevention:** Use `System.Security.Cryptography.RandomNumberGenerator::Create()` for generating strong, unpredictable random numbers for security-critical tasks like password generation, tokens, or encryption keys.
## 2024-05-18 - Prevent Command Injection with EncodedCommand
**Vulnerability:** Constructing unescaped remote PowerShell execution commands using string formatting via `Start-Process powershell.exe -ArgumentList "-Command", "Remove-AppxPackage -Package '$targetId'"`.
**Learning:** Even with basic single-quote escaping like `$targetId.Replace("'", "''")`, injecting strings into `-Command` exposes the target machine to command injection if the variable contains unescaped characters like `$()`, `"`, or `;`.
**Prevention:** Eliminate code-as-string execution vulnerabilities by completely decoupling inputs from the parsed command string. To securely pass external parameters into a separate `powershell.exe` process, package the parameter data alongside the static command using `System.Text.Encoding` and `[Convert]::ToBase64String`. Supply this as an `-EncodedCommand` rather than directly stringifying `-Command` arguments.
