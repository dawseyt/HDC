## 2024-05-18 - Cryptographically Secure Random Number Generation
**Vulnerability:** Weak random number generation (`Get-Random`) used for creating new Active Directory user passwords.
**Learning:** `Get-Random` in PowerShell 5.1 relies on the non-cryptographically secure `System.Random` class, meaning that the generated passwords could theoretically be predicted by a malicious actor who observes the PRNG's outputs.
**Prevention:** Use `System.Security.Cryptography.RandomNumberGenerator::Create()` for generating strong, unpredictable random numbers for security-critical tasks like password generation, tokens, or encryption keys.

## 2024-06-03 - [Cross-Site Scripting (XSS) in HTML Dashboard Generation]
**Vulnerability:** Embedded JavaScript in `GenDash.ps1` dynamically injected unescaped data from external sources (logs and Active Directory) into the DOM using `.innerHTML`.
**Learning:** If user-controlled data or data from external sources is injected into HTML without proper escaping, malicious users can include `<script>` tags or other HTML attributes (like `onload` or `onerror` event handlers) that the browser will execute, leading to XSS vulnerabilities.
**Prevention:** Implement and use a custom `escapeHtml` function that replaces problematic characters (`&`, `<`, `>`, `"`, `'`) with their respective HTML entities before injecting dynamically retrieved data into `.innerHTML`. Data assigned to `.innerText` or `textContent` is inherently safe and does not need to be escaped.
