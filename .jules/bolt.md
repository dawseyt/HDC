## 2025-02-12 - [PowerShell Get-ChildItem path bug and performance optimization]
**Learning:** `Get-ChildItem -Filter` breaks when `LiteralPath` contains square brackets `[]` (like `[toolkit]`). However, falling back to pipeline filtering using `Where-Object` is extremely slow.
**Action:** Use `[System.Management.Automation.WildcardPattern]::Escape($path)` to escape the brackets, then use `-Path` and `-Filter` for drastically faster native filtering.

## 2025-03-31 - [Optimize DOM rendering in embedded JavaScript]
**Learning:** Embedded JavaScript manipulating the DOM used iterative `.innerHTML +=` inside loops, causing significant O(n^2) DOM reflow bottlenecks.
**Action:** Replace iterative concatenation with array accumulations (`map().join("")`) and assign them to `.innerHTML` in a single operation. Use `document.createElement` for dynamic one-off elements like error banners to further improve safety and performance.

## 2025-04-05 - [PowerShell Array Accumulation Performance Anti-pattern]
**Learning:** Using `+=` to append to an array inside a loop in PowerShell (e.g., `$array += $item`) is an O(n^2) operation because PowerShell arrays are fixed size, so the engine completely recreates the array on every iteration. This creates severe performance bottlenecks when fetching large lists of remote data like files, event logs, and processes.
**Action:** Always use pipeline assignments with `@(foreach...)` (e.g., `$array = @(foreach ($item in $items) { $item })`) or a strongly-typed `[System.Collections.Generic.List[type]]` for efficient O(n) array accumulation.
