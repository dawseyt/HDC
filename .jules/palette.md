## 2024-05-24 - Initial Setup
**Learning:** This is the initial setup of the palette log.
**Action:** Keep track of critical UX learnings here.

## 2024-05-24 - Tooltips
**Learning:** Adding explicit tooltips to main action buttons helps provide context to the users, allowing them to know exactly what the buttons will do.
**Action:** When creating action buttons without explicitly clear texts, make sure to add `ToolTip` elements to them.

## 2024-05-25 - WPF Screen Reader Accessibility
**Learning:** WPF text boxes and password boxes that rely on separate TextBlocks for visual labels need explicit `AutomationProperties.Name` tags for proper screen reader accessibility, acting as the equivalent of ARIA labels.
**Action:** Ensure form fields in WPF layouts have `AutomationProperties.Name` applied if their label is a separate element.## $(date +%Y-%m-%d) - Dialog Accessibility and Tooltips
**Learning:** Dialogs (like GPResult and PrinterInstall) often lack accessible names for text boxes and comboboxes when their visual labels are standalone `TextBlock` elements. Screen readers won't automatically associate them without `AutomationProperties.Name`.
**Action:** When creating or modifying WPF input fields, always ensure `AutomationProperties.Name` and a descriptive `ToolTip` are explicitly set, even if a nearby `TextBlock` serves as a visual label.
