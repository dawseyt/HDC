## 2025-02-12 - [PowerShell Get-ChildItem path bug and performance optimization]
**Learning:** `Get-ChildItem -Filter` breaks when `LiteralPath` contains square brackets `[]` (like `[toolkit]`). However, falling back to pipeline filtering using `Where-Object` is extremely slow.
**Action:** Use `[System.Management.Automation.WildcardPattern]::Escape($path)` to escape the brackets, then use `-Path` and `-Filter` for drastically faster native filtering.

## 2025-03-31 - [Optimize DOM rendering in embedded JavaScript]
**Learning:** Embedded JavaScript manipulating the DOM used iterative `.innerHTML +=` inside loops, causing significant O(n^2) DOM reflow bottlenecks.
**Action:** Replace iterative concatenation with array accumulations (`map().join("")`) and assign them to `.innerHTML` in a single operation. Use `document.createElement` for dynamic one-off elements like error banners to further improve safety and performance.

## 2025-04-10 - [PowerShell Pipeline Overhead in Tight Loops]
**Learning:** Using `Where-Object` inside a tight polling loop (e.g., waiting for jobs to complete) incurs significant pipeline overhead on every iteration, drastically slowing down execution and wasting CPU cycles compared to standard flow control.
**Action:** Replace `Where-Object` pipelines inside `while` loops with native `foreach` loops and early `break` statements. This avoids pipeline instantiation overhead and allows for efficient short-circuit evaluation.
