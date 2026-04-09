## 2025-02-12 - [PowerShell Get-ChildItem path bug and performance optimization]
**Learning:** `Get-ChildItem -Filter` breaks when `LiteralPath` contains square brackets `[]` (like `[toolkit]`). However, falling back to pipeline filtering using `Where-Object` is extremely slow.
**Action:** Use `[System.Management.Automation.WildcardPattern]::Escape($path)` to escape the brackets, then use `-Path` and `-Filter` for drastically faster native filtering.

## 2025-03-31 - [Optimize DOM rendering in embedded JavaScript]
**Learning:** Embedded JavaScript manipulating the DOM used iterative `.innerHTML +=` inside loops, causing significant O(n^2) DOM reflow bottlenecks.
**Action:** Replace iterative concatenation with array accumulations (`map().join("")`) and assign them to `.innerHTML` in a single operation. Use `document.createElement` for dynamic one-off elements like error banners to further improve safety and performance.
## 2026-04-09 - Avoid += for Array Concatenation in Loops
**Learning:** In PowerShell, using the `+=` operator to append elements to an array inside loops (e.g., when collecting hardware information or lists of remote processes) forces the entire array to be re-allocated and copied on every iteration. This leads to an $O(n^2)$ time complexity which causes severe performance degradation for large datasets.
**Action:** Instead of `+=`, use array subexpressions (e.g., `$array = @(foreach ($item in $collection) { ... })`) or pipeline assignments which pre-allocate the data and ensure an $O(n)$ time complexity, maintaining fast and efficient performance.
