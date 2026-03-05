# Architecture Overview

## Language Runtime
**PowerShell 5.1+ (.NET API Integration)**

We use PowerShell rather than C# or Python because IT Desktop Support users operate universally within PS environments. It prevents needing to ship massive `.NET Desktop Runtime` dependencies.

## GUI Rendering
**WPF (Windows Presentation Foundation) XAML**
The UI layout is coded directly into an `[xml]$xaml` object array. PowerShell `[Windows.Markup.XamlReader]::Load($reader)` dynamically converts this schema into an executable overlay.

Because it mounts directly to the internal .NET hook, we have full access to native UI events like `$btnGenerate.Add_Click` or manipulating the `$Dispatcher` to keep the UI active during heavy threads.

## Workflow Pipeline
1. **Load:** Mount the UI XML structure via PowerShell. Setup `.Text` field bindings.
2. **Execute:** Initiate a Background block to prevent the main thread from freezing.
3. **Capture:** Map `Win32` object layers into isolated arrays.
4. **Compile:** Convert variables into a large HTML string mapped with custom CSS rules locally injected.
5. **Print:** Pipe the `$HtmlBody` file into the local MS Edge engine renderer to bind it as a rigid PDF frame.
