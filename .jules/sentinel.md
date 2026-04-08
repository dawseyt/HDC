## 2024-05-18 - Cryptographically Secure Random Number Generation
**Vulnerability:** Weak random number generation (`Get-Random`) used for creating new Active Directory user passwords.
**Learning:** `Get-Random` in PowerShell 5.1 relies on the non-cryptographically secure `System.Random` class, meaning that the generated passwords could theoretically be predicted by a malicious actor who observes the PRNG's outputs.
**Prevention:** Use `System.Security.Cryptography.RandomNumberGenerator::Create()` for generating strong, unpredictable random numbers for security-critical tasks like password generation, tokens, or encryption keys.

## 2024-05-18 - Cross-Site Scripting (XSS) in Dashboard Generation
**Vulnerability:** Untrusted string data (like Active Directory usernames or log strings) was being directly interpolated into Javascript and assigned to `.innerHTML` in `GenDash.ps1`.
**Learning:** This occurs when server-side scripts (like PowerShell) inject raw data directly into the DOM structure of generated HTML/JS files, meaning malicious strings in the source logs could break out and execute arbitrary scripts in the viewer's browser.
**Prevention:** Whenever assigning user-controlled or dynamically retrieved external text data to `.innerHTML`, it must be processed through an HTML escaping function (e.g., replacing `<`, `>`, `&`, `"`, `'` with their HTML entities) before injection.
