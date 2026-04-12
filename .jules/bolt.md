## 2025-02-12 - [PowerShell Get-ChildItem path bug and performance optimization]
**Learning:** `Get-ChildItem -Filter` breaks when `LiteralPath` contains square brackets `[]` (like `[toolkit]`). However, falling back to pipeline filtering using `Where-Object` is extremely slow.
**Action:** Use `[System.Management.Automation.WildcardPattern]::Escape($path)` to escape the brackets, then use `-Path` and `-Filter` for drastically faster native filtering.

## 2025-03-31 - [Optimize DOM rendering in embedded JavaScript]
**Learning:** Embedded JavaScript manipulating the DOM used iterative `.innerHTML +=` inside loops, causing significant O(n^2) DOM reflow bottlenecks.
**Action:** Replace iterative concatenation with array accumulations (`map().join("")`) and assign them to `.innerHTML` in a single operation. Use `document.createElement` for dynamic one-off elements like error banners to further improve safety and performance.
## 2026-04-12 - [O(N^2) Array Accumulation in PowerShell]
**Learning:** Using `+=` to append to an array inside a loop in PowerShell creates a severe O(N^2) bottleneck because PowerShell arrays are immutable and fixed-size. The entire array must be copied into a new memory location on every iteration.
**Action:** Always wrap the entire loop in an array subexpression: `$array = @(foreach (...) { ... })`. This leverages the pipeline to gather all output optimally in O(N) time without memory reallocation penalties. Note that `+=` is perfectly acceptable for integer counters or string concatenations.
