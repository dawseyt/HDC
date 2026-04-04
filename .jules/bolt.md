## 2025-02-12 - [PowerShell Get-ChildItem path bug and performance optimization]
**Learning:** `Get-ChildItem -Filter` breaks when `LiteralPath` contains square brackets `[]` (like `[toolkit]`). However, falling back to pipeline filtering using `Where-Object` is extremely slow.
**Action:** Use `[System.Management.Automation.WildcardPattern]::Escape($path)` to escape the brackets, then use `-Path` and `-Filter` for drastically faster native filtering.

## 2025-03-31 - [Optimize DOM rendering in embedded JavaScript]
**Learning:** Embedded JavaScript manipulating the DOM used iterative `.innerHTML +=` inside loops, causing significant O(n^2) DOM reflow bottlenecks.
**Action:** Replace iterative concatenation with array accumulations (`map().join("")`) and assign them to `.innerHTML` in a single operation. Use `document.createElement` for dynamic one-off elements like error banners to further improve safety and performance.

## 2025-05-18 - [Optimize PowerShell Array Accumulation]
**Learning:** Using the `+=` operator inside loops to accumulate items in fixed-size arrays results in O(n^2) performance because it forces PowerShell to recreate the entire array on every iteration.
**Action:** Replace `+=` assignments inside loops with native array/pipeline assignments (`$array = @(foreach(...) { ... })`) to bring the operation time complexity down to O(n), vastly reducing execution time for large lists.
