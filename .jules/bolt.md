## 2025-02-12 - [PowerShell Get-ChildItem path bug and performance optimization]
**Learning:** `Get-ChildItem -Filter` breaks when `LiteralPath` contains square brackets `[]` (like `[toolkit]`). However, falling back to pipeline filtering using `Where-Object` is extremely slow.
**Action:** Use `[System.Management.Automation.WildcardPattern]::Escape($path)` to escape the brackets, then use `-Path` and `-Filter` for drastically faster native filtering.

## 2025-03-31 - [Optimize DOM rendering in embedded JavaScript]
**Learning:** Embedded JavaScript manipulating the DOM used iterative `.innerHTML +=` inside loops, causing significant O(n^2) DOM reflow bottlenecks.
**Action:** Replace iterative concatenation with array accumulations (`map().join("")`) and assign them to `.innerHTML` in a single operation. Use `document.createElement` for dynamic one-off elements like error banners to further improve safety and performance.

## 2025-04-08 - [Optimize PowerShell Array Accumulation in Loops]
**Learning:** Appending items to fixed-size PowerShell arrays inside loops using `+=` (e.g. `$resProcs += [PSCustomObject]@{...}`) causes massive $O(n^2)$ memory reallocation overhead, especially for large datasets like running processes or system event logs.
**Action:** Refactor iterative array concatenation using array subexpression assignment (`$array = @(foreach(...) { ... })`). This is significantly faster and eliminates the pipeline overhead. However, be careful not to unnecessarily refactor variables that are just integers (like `$filesRemoved += 5`), as simple integer addition is perfectly fine and performant.