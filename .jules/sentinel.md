## 2024-05-18 - Cryptographically Secure Random Number Generation
**Vulnerability:** Weak random number generation (`Get-Random`) used for creating new Active Directory user passwords.
**Learning:** `Get-Random` in PowerShell 5.1 relies on the non-cryptographically secure `System.Random` class, meaning that the generated passwords could theoretically be predicted by a malicious actor who observes the PRNG's outputs.
**Prevention:** Use `System.Security.Cryptography.RandomNumberGenerator::Create()` for generating strong, unpredictable random numbers for security-critical tasks like password generation, tokens, or encryption keys.

## 2026-04-17 - Command Injection in Start-Process PowerShell Calls
**Vulnerability:** Command injection vulnerability when passing user-controlled variables (like `$ComputerName` or `$targetId`) into a PowerShell string intended to be executed via `Start-Process powershell.exe -Command`.
**Learning:** Naive escaping using string replacement (e.g., `.Replace("'", "''")`) is insufficient against sophisticated injection payloads. When user input is directly concatenated into a `-Command` argument, a properly crafted input string can still break out of the string boundary or utilize other PowerShell features to execute arbitrary code.
**Prevention:** Always treat untrusted input strictly as data. Base64 encode the user-controlled string in the parent scope, pass it to a Base64 encoded wrapper script using `-EncodedCommand` (which decodes it securely back into a string variable before utilizing it), eliminating any risk of the data being interpreted as executable code.
