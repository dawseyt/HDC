## 2024-05-18 - Cryptographically Secure Random Number Generation
**Vulnerability:** Weak random number generation (`Get-Random`) used for creating new Active Directory user passwords.
**Learning:** `Get-Random` in PowerShell 5.1 relies on the non-cryptographically secure `System.Random` class, meaning that the generated passwords could theoretically be predicted by a malicious actor who observes the PRNG's outputs.
**Prevention:** Use `System.Security.Cryptography.RandomNumberGenerator::Create()` for generating strong, unpredictable random numbers for security-critical tasks like password generation, tokens, or encryption keys.
## 2024-05-18 - Command Injection via Start-Process powershell.exe -Command
**Vulnerability:** Untrusted user input was poorly sanitized using `.Replace("'", "''")` and directly concatenated into the `-Command` string argument for `powershell.exe`. This is vulnerable to edge-case escape sequences and command injection.
**Learning:** Manual string replacement is never sufficient for preventing command injection in shell execution. PowerShell's parsing engine handles complex escape sequences, making string concatenation fundamentally unsafe.
**Prevention:** Treat untrusted input strictly as data by Base64 encoding it in the parent script, inserting the encoded payload into the command string, and decoding it within the child process script before execution via `-EncodedCommand`.
