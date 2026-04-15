## 2024-05-18 - Cryptographically Secure Random Number Generation
**Vulnerability:** Weak random number generation (`Get-Random`) used for creating new Active Directory user passwords.
**Learning:** `Get-Random` in PowerShell 5.1 relies on the non-cryptographically secure `System.Random` class, meaning that the generated passwords could theoretically be predicted by a malicious actor who observes the PRNG's outputs.
**Prevention:** Use `System.Security.Cryptography.RandomNumberGenerator::Create()` for generating strong, unpredictable random numbers for security-critical tasks like password generation, tokens, or encryption keys.
## 2024-05-15 - [Prevent Command Injection in PowerShell Child Processes]
**Vulnerability:** Found a command injection vulnerability where user-controlled strings (like AppX package names) were sanitized via manual string replacement (`.Replace("'", "''")`) and passed directly to `Start-Process powershell.exe -Command "$psCmd"`.
**Learning:** Manual string replacement is vulnerable to edge cases, escape sequence injections, and unexpected parsing logic.
**Prevention:** Avoid `Start-Process powershell.exe -Command` for executing dynamically generated scripts with user input. Instead, treat untrusted input as data by dynamically generating the PowerShell code, Base64 encoding it using UTF-16LE, and passing it via `-EncodedCommand`.
