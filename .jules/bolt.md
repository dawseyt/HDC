## 2025-02-12 - [PowerShell Get-ChildItem path bug and performance optimization]
**Learning:** `Get-ChildItem -Filter` breaks when `LiteralPath` contains square brackets `[]` (like `[toolkit]`). However, falling back to pipeline filtering using `Where-Object` is extremely slow.
**Action:** Use `[System.Management.Automation.WildcardPattern]::Escape($path)` to escape the brackets, then use `-Path` and `-Filter` for drastically faster native filtering.

## 2025-03-31 - [Optimize DOM rendering in embedded JavaScript]
**Learning:** Embedded JavaScript manipulating the DOM used iterative `.innerHTML +=` inside loops, causing significant O(n^2) DOM reflow bottlenecks.
**Action:** Replace iterative concatenation with array accumulations (`map().join("")`) and assign them to `.innerHTML` in a single operation. Use `document.createElement` for dynamic one-off elements like error banners to further improve safety and performance.

## 2025-04-10 - [Avoid O(n^2) array operations in loops]
**Learning:** Using `+=` to append to arrays inside `foreach` or `for` loops in PowerShell forces the array to be rebuilt on each iteration, causing significant performance degradation (O(n^2) complexity). Multi-line formatting in subexpressions is crucial to avoid creating massive, unreadable one-liners.
**Action:** Replace `+=` with variable assignment from a subexpression loop (`$array = @(foreach(...) { ... })`). Ensure the inner object creation logic spans multiple lines for readability.
