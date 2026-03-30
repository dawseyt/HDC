## 2024-05-18 - Cryptographically Secure Random Number Generation
**Vulnerability:** Weak random number generation (`Get-Random`) used for creating new Active Directory user passwords.
**Learning:** `Get-Random` in PowerShell 5.1 relies on the non-cryptographically secure `System.Random` class, meaning that the generated passwords could theoretically be predicted by a malicious actor who observes the PRNG's outputs.
**Prevention:** Use `System.Security.Cryptography.RandomNumberGenerator::Create()` for generating strong, unpredictable random numbers for security-critical tasks like password generation, tokens, or encryption keys.
