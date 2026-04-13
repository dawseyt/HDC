## 2025-02-12 - [PowerShell Get-ChildItem path bug and performance optimization]
**Learning:** `Get-ChildItem -Filter` breaks when `LiteralPath` contains square brackets `[]` (like `[toolkit]`). However, falling back to pipeline filtering using `Where-Object` is extremely slow.
**Action:** Use `[System.Management.Automation.WildcardPattern]::Escape($path)` to escape the brackets, then use `-Path` and `-Filter` for drastically faster native filtering.

## 2025-03-31 - [Optimize DOM rendering in embedded JavaScript]
**Learning:** Embedded JavaScript manipulating the DOM used iterative `.innerHTML +=` inside loops, causing significant O(n^2) DOM reflow bottlenecks.
**Action:** Replace iterative concatenation with array accumulations (`map().join("")`) and assign them to `.innerHTML` in a single operation. Use `document.createElement` for dynamic one-off elements like error banners to further improve safety and performance.

## 2025-04-10 - [PowerShell Array Accumulation Performance Pitfall]
**Learning:** Using the `+=` operator to append items to a standard array inside a loop (e.g., `$array += $item`) causes severe O(n^2) performance degradation because PowerShell recreates the entire array in memory on every iteration.
**Action:** Replace inline array additions with array subexpressions. Wrap the loop in `@(...)` and assign the result directly to the variable (e.g., `$array = @(foreach ($item in $collection) { ... })`), which operates in linear O(n) time.
