## 2025-02-12 - [PowerShell Get-ChildItem path bug and performance optimization]
**Learning:** `Get-ChildItem -Filter` breaks when `LiteralPath` contains square brackets `[]` (like `[toolkit]`). However, falling back to pipeline filtering using `Where-Object` is extremely slow.
**Action:** Use `[System.Management.Automation.WildcardPattern]::Escape($path)` to escape the brackets, then use `-Path` and `-Filter` for drastically faster native filtering.

## 2025-03-31 - [Optimize DOM rendering in embedded JavaScript]
**Learning:** Embedded JavaScript manipulating the DOM used iterative `.innerHTML +=` inside loops, causing significant O(n^2) DOM reflow bottlenecks.
**Action:** Replace iterative concatenation with array accumulations (`map().join("")`) and assign them to `.innerHTML` in a single operation. Use `document.createElement` for dynamic one-off elements like error banners to further improve safety and performance.

## 2026-04-11 - [Refactoring PowerShell Array Accumulation]
**Learning:** Using the `+=` operator to append to arrays inside PowerShell loops causes O(n^2) performance penalties because arrays are immutable, forcing recreation on each iteration.
**Action:** Use pipeline assignment with an array subexpression (e.g., `$results = @(foreach(...) {...})`) or a `[System.Collections.Generic.List[type]]` to accumulate elements efficiently in O(n) time.
