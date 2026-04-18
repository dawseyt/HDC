## 2024-05-24 - Initial Setup
**Learning:** This is the initial setup of the palette log.
**Action:** Keep track of critical UX learnings here.

## 2024-05-24 - Tooltips
**Learning:** Adding explicit tooltips to main action buttons helps provide context to the users, allowing them to know exactly what the buttons will do.
**Action:** When creating action buttons without explicitly clear texts, make sure to add `ToolTip` elements to them.

## 2024-05-25 - WPF Screen Reader Accessibility
**Learning:** WPF text boxes and password boxes that rely on separate TextBlocks for visual labels need explicit `AutomationProperties.Name` tags for proper screen reader accessibility, acting as the equivalent of ARIA labels.
**Action:** Ensure form fields in WPF layouts have `AutomationProperties.Name` applied if their label is a separate element.
## 2024-05-25 - Form Field Tooltips
**Learning:** Adding contextual `ToolTip`s to form fields, such as those in the Printer Install dialog, is an effective way to guide users on expected input formats without cluttering the UI with additional helper text.
**Action:** When adding or updating form inputs, include a `ToolTip` to explain the expected value if it's not immediately obvious from the label.
