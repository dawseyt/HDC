## 2024-05-24 - Initial Setup
**Learning:** This is the initial setup of the palette log.
**Action:** Keep track of critical UX learnings here.

## 2024-05-24 - Tooltips
**Learning:** Adding explicit tooltips to main action buttons helps provide context to the users, allowing them to know exactly what the buttons will do.
**Action:** When creating action buttons without explicitly clear texts, make sure to add `ToolTip` elements to them.

## 2024-05-25 - WPF Screen Reader Accessibility
**Learning:** WPF text boxes and password boxes that rely on separate TextBlocks for visual labels need explicit `AutomationProperties.Name` tags for proper screen reader accessibility, acting as the equivalent of ARIA labels.
**Action:** Ensure form fields in WPF layouts have `AutomationProperties.Name` applied if their label is a separate element.

## 2024-04-14 - WPF Screen Reader Accessibility for Inputs
**Learning:** Added 'AutomationProperties.Name' and 'ToolTip' to WPF inputs (like TextBox and DatePicker) and buttons whose visual labels were disconnected 'TextBlock' elements. This drastically improves screen reader support, serving as the XAML equivalent to ARIA labels, while also giving visual users on-hover context.
**Action:** Always add 'AutomationProperties.Name' and 'ToolTip' to interactive elements in WPF (inputs, buttons) when their labels are not directly embedded as content.
