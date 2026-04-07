## 2024-05-24 - Initial Setup
**Learning:** This is the initial setup of the palette log.
**Action:** Keep track of critical UX learnings here.

## 2024-05-24 - Tooltips
**Learning:** Adding explicit tooltips to main action buttons helps provide context to the users, allowing them to know exactly what the buttons will do.
**Action:** When creating action buttons without explicitly clear texts, make sure to add `ToolTip` elements to them.

## 2024-05-25 - WPF Screen Reader Accessibility
**Learning:** WPF text boxes and password boxes that rely on separate TextBlocks for visual labels need explicit `AutomationProperties.Name` tags for proper screen reader accessibility, acting as the equivalent of ARIA labels.
**Action:** Ensure form fields in WPF layouts have `AutomationProperties.Name` applied if their label is a separate element.

## 2024-05-26 - Custom Window Close Buttons Accessibility
**Learning:** When implementing custom window close buttons in WPF (e.g., using Segoe UI Symbol icons like `&#xE711;`), explicit `ToolTip="Close"` and `AutomationProperties.Name="Close"` are required to ensure both visual clarity and screen reader accessibility, as the symbol alone conveys no meaning to assistive technologies.
**Action:** Always verify that custom icon-only close controls have explicit tooltip and automation name properties defined.