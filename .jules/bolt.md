## 2024-05-18 - [PowerShell Array Appending Optimization]
**Learning:** In PowerShell, using the `+=` operator to append items to an array inside loops causes O(n^2) time complexity because arrays are immutable. The entire array is recreated on each iteration.
**Action:** Always use pipeline assignment (e.g., `$array = foreach (...) { ... }`) or `[System.Collections.Generic.List[type]]` for efficient accumulation when building arrays in a loop.
