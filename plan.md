# Project Plan: Greenshot Macro Automation ("Middle Path")

## Objective
To develop a declarative, JSON-driven automation engine for Greenshot that allows external scripts to trigger complex workflows (Capture -> Annotate -> Export) without requiring a full COM/Inter-process communication setup.

---

## 1. High-Level Architecture
- **Trigger**: Command-line argument `--macro [path_to_json]`.
- **Implementation**: A new Plugin `Greenshot.Plugin.MacroRunner` or a core module in `Greenshot.exe`.
- **Execution Model**: Sequential execution of steps defined in the JSON schema.
- **Engine**: A "headless" or "background" instance of `ISurface` to perform annotations before passing to destinations.

---

## 2. Development Milestones

### Milestone 1: Foundation & Infrastructure (DONE)
*   [x] **Schema Definition**: Finalize the JSON schema (Capture, Annotate, Export blocks).
*   [x] **CLI Integration**: Add command-line argument parsing in `MainForm.cs` or `GreenshotMain.cs` to detect the `--macro` flag.
*   [x] **Parser Setup**: Integrate `System.Text.Json` or `Newtonsoft.Json` into the build process.
*   [x] **Logging**: Implement a specialized log for the Macro Runner to debug script failures.

### Milestone 2: Automated Capture Engine (DONE)
*   [x] **Silent Capture**: Modify `CaptureHelper` calls to skip interactive selection when coordinates/HWNDs are provided in JSON.
*   [x] **Window Search Logic**: Implement a "finder" within the macro runner to locate windows by title or process name before capture.
*   [x] **Capture Actions**:
    *   `CaptureRegion`
    *   `CaptureWindow`
    *   `CaptureFullscreen`

### Milestone 3: Annotation & Surface Mapper (DONE)
*   [x] **Surface Orchestration**: Create a non-visible `Surface` instance that can host a capture.
*   [x] **Element Converters**: Write mappers for:
    *   **Shapes**: `RectangleContainer`, `EllipseContainer`, `LineContainer`.
    *   **Annotations**: `ArrowContainer`, `TextContainer`.
    *   **Effects**: `HighlightContainer`, `ObfuscateContainer` (Blur/Pixelate).
*   [x] **Property Binding**: Map style properties (Hex Colors, Thickness, Opacity) from JSON to internal `Field` values.

### Milestone 4: Export & Workflow Finishing (DONE)
*   [x] **Destination Routing**: Map JSON export types to existing `IDestination` plugins.
*   [x] **Variable Injection**: Support variables like `{timestamp}`, `{active_window_title}`, or `{env:TMP}` in file paths.
*   [x] **Multi-Export**: Allow a single script to export to multiple places (e.g., File + Clipboard + Imgur).

### Milestone 5: Refinement & Error Handling (DONE)
*   [x] **Validation**: JSON Schema validation before execution.
*   [x] **Transaction Safety**: Ensure that if one step fails (e.g., window not found), the script exits gracefully without leaving Greenshot in a broken state.
*   [x] **Documentation**: Create a handbook of all supported JSON keys and examples.

---

## 3. Technical Requirements & Challenges

| Feature | Technical Challenge | Solution |
| :--- | :--- | :--- |
| **Silent UI** | Greenshot is heavily tied to Windows Forms UI events. | Use `BeginInvoke` to ensure macro steps run on the UI thread without blocking the runner. |
| **Object Lifecycle** | `ISurface` and `Image` objects need proper disposal. | Wrap the Macro execution in a `using` block to prevent GDI+ memory leaks. |
| **Z-Order/Sync** | Ensuring annotations are finished before export starts. | Implement an internal `Task` based sequence or simple synchronous calls for the macro runner. |

---

## 4. Success Criteria
1.  Ability to take a screenshot and save it to a specific folder using a single one-line command.
2.  Ability to draw a red arrow at specific coordinates on a captured image without user interaction.
3.  Error reporting that tells the user exactly which line of their JSON failed.
