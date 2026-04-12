## 2024-05-18 - Cryptographically Secure Random Number Generation
**Vulnerability:** Weak random number generation (`Get-Random`) used for creating new Active Directory user passwords.
**Learning:** `Get-Random` in PowerShell 5.1 relies on the non-cryptographically secure `System.Random` class, meaning that the generated passwords could theoretically be predicted by a malicious actor who observes the PRNG's outputs.
**Prevention:** Use `System.Security.Cryptography.RandomNumberGenerator::Create()` for generating strong, unpredictable random numbers for security-critical tasks like password generation, tokens, or encryption keys.

## 2024-05-18 - [Prevent XSS in Embedded JavaScript Template Literals]
**Vulnerability:** The HTML dashboard generator (`GenDash.ps1`) embedded user-controlled data (Active Directory usernames, log events) directly into raw HTML template literals via `.innerHTML` assignments.
**Learning:** Assigning unescaped data to `.innerHTML` exposes the application to Cross-Site Scripting (XSS). If log data contains `<script>` or other malicious payloads, it executes within the browser context. Data assigned to `.innerText` is inherently safe, but dynamically constructed HTML tags must have their variables sanitized.
**Prevention:** Implement and apply a custom `escapeHtml` function to sanitize user-controlled variables before they are interpolated into strings destined for `.innerHTML`.
