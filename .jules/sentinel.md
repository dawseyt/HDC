## 2024-05-18 - Cryptographically Secure Random Number Generation
**Vulnerability:** Weak random number generation (`Get-Random`) used for creating new Active Directory user passwords.
**Learning:** `Get-Random` in PowerShell 5.1 relies on the non-cryptographically secure `System.Random` class, meaning that the generated passwords could theoretically be predicted by a malicious actor who observes the PRNG's outputs.
**Prevention:** Use `System.Security.Cryptography.RandomNumberGenerator::Create()` for generating strong, unpredictable random numbers for security-critical tasks like password generation, tokens, or encryption keys.
## 2024-05-24 - Secure Start-Process Parameter Passing
**Vulnerability:** Command injection when unsanitized user input (AppX identifiers) was formatted using string replacement `.Replace("'", "''")` and passed via `-Command` to `Start-Process powershell.exe`.
**Learning:** Manual escaping of user-controlled variables in PowerShell is prone to edge-case bypasses and command injection, especially when spawned as child processes using `-Command`.
**Prevention:** Treat untrusted input as data by passing it via `-EncodedCommand`. First Base64 encode the user input, write a script block that decodes the input back into a string natively within the payload, and finally Base64 encode the entire payload (UTF-16LE) to pass safely to `Start-Process powershell.exe`.
