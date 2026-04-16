## 2025-02-12 - [PowerShell Get-ChildItem path bug and performance optimization]
**Learning:** `Get-ChildItem -Filter` breaks when `LiteralPath` contains square brackets `[]` (like `[toolkit]`). However, falling back to pipeline filtering using `Where-Object` is extremely slow.
**Action:** Use `[System.Management.Automation.WildcardPattern]::Escape($path)` to escape the brackets, then use `-Path` and `-Filter` for drastically faster native filtering.

## 2025-03-31 - [Optimize DOM rendering in embedded JavaScript]
**Learning:** Embedded JavaScript manipulating the DOM used iterative `.innerHTML +=` inside loops, causing significant O(n^2) DOM reflow bottlenecks.
**Action:** Replace iterative concatenation with array accumulations (`map().join("")`) and assign them to `.innerHTML` in a single operation. Use `document.createElement` for dynamic one-off elements like error banners to further improve safety and performance.

## 2025-05-18 - [Optimize PowerShell Array Accumulation in Loops]
**Learning:** Using the `+=` operator to append objects to an array inside a `foreach` loop causes O(n^2) time complexity because PowerShell must recreate the entire array in memory on every iteration. This severely degraded performance in `UI.ComputerActions.psm1` when fetching remote computer details (services, processes, profiles, event logs, etc.).
**Action:** Replace `+=` inside loops with array subexpressions `$var = @(foreach ($x in $y) { [PSCustomObject]@{...} })`. This guarantees array typing while dropping time complexity to O(n).
