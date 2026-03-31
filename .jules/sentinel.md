## 2024-05-18 - Cryptographically Secure Random Number Generation
**Vulnerability:** Weak random number generation (`Get-Random`) used for creating new Active Directory user passwords.
**Learning:** `Get-Random` in PowerShell 5.1 relies on the non-cryptographically secure `System.Random` class, meaning that the generated passwords could theoretically be predicted by a malicious actor who observes the PRNG's outputs.
**Prevention:** Use `System.Security.Cryptography.RandomNumberGenerator::Create()` for generating strong, unpredictable random numbers for security-critical tasks like password generation, tokens, or encryption keys.

## 2024-05-18 - Hardcoded Freshservice API Key
**Vulnerability:** A hardcoded API key for Freshservice (`"FreshserviceAPIKey": "QBZwMMFeWzmRVG8FGN5"`) was committed in the default configuration template `HDCompanionCfg.json`.
**Learning:** Hardcoded credentials exposed in configuration templates or codebase files are severe security vulnerabilities as they provide unauthorized access immediately upon clone or copy. Even if the codebase implements a secure alternative, template files can easily become the source of truth for new deployments if mismanaged.
**Prevention:** Never commit API keys, passwords, tokens, or other secrets to any file tracked by version control. Utilize proper secure credential storage systems (like Windows Credential Manager as handled in `Set-FSApiKey`) or environment variables, and verify configuration templates do not contain sensitive data before pushing.
