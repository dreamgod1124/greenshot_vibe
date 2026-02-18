# Greenshot Macro Handbook

Greenshot macros allow you to automate screenshot tasks using JSON configuration files. You can trigger a macro via the command line:

`Greenshot.exe /macro "C:\path\to\your_macro.json"`

---

## JSON Structure Overview

A macro file consists of a version and a list of workflow steps.

```json
{
  "version": "1.0",
  "workflow": [
    { "step": "capture", ... },
    { "step": "annotate", ... },
    { "step": "export", ... }
  ]
}
```

---

## 1. Capture Step (`capture`)

Triggers a new screenshot or loads an existing image.

| Key | Value | Description |
| :--- | :--- | :--- |
| `type` | `"fullscreen"` \| `"region"` \| `"file"` | The capture mode. |
| `area` | Object | (Optional) Required for `"region"`. Defines the capture area. |
| `path` | String | (Optional) Required for `"file"`. Path to the source image. |
| `autocrop` | Boolean | (Optional) Overrides the global "Auto crop" setting. |
| `autocrop_difference` | Integer | (Optional) Sets the threshold for autocrop (0-255). |

**`area` Object Properties:**
- `X`, `Y`: Top-left coordinates.
- `Width`, `Height`: Dimensions of the capture.

**Example (Region):**
```json
{
  "step": "capture",
  "type": "region",
  "area": { "X": 0, "Y": 0, "Width": 800, "Height": 600 }
}
```

**Example (File):**
```json
{
  "step": "capture",
  "type": "file",
  "path": "C:\\Images\\source.png"
}
```

---

## 2. Annotate Step (`annotate`)

Adds drawing elements to the current capture.

| Key | Value | Description |
| :--- | :--- | :--- |
| `elements` | Array of Objects | List of shapes/annotations to add. |

### Supported Elements

#### Rectangle
- `type`: `"rectangle"`
- `bounds`: `{ "X": int, "Y": int, "W": int, "H": int }`
- `style`: Style object (see below).

#### Arrow
- `type`: `"arrow"`
- `from`: `{ "X": int, "Y": int }` (Start point)
- `to`: `{ "X": int, "Y": int }` (End point / Head)
- `style`: Style object.

#### Text
- `type`: `"text"`
- `position`: `{ "X": int, "Y": int }`
- `content`: `"string"` (The text to display)
- `style`: Style object.

#### Obfuscate (Blur / Pixelate)
- `type`: `"obfuscate"`
- `bounds`: `{ "X": int, "Y": int, "W": int, "H": int }`
- `style`: Style object. 
  *   Use `blur_radius` in style for blurring.
  *   Use `pixel_size` in style for pixelation.

### Style Object Properties
Keys are case-insensitive and can include underscores (e.g., `line_color` or `linecolor`).

| Property | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `line_color` | String | `#FF0000` | Hex code (e.g., `"#FF0000"`) or standard color name. |
| `fill_color` | String | `transparent` | Hex code or `"transparent"`. |
| `line_thickness` | Integer | `3` | Thickness of lines/borders. |
| `shadow` | Boolean | `false` | Enable drop shadow. |
| `font_size` | Float | `12` | Font size (Text only). |
| `font_bold` | Boolean | `false` | Bold text (Text only). |
| `font_italic` | Boolean | `false` | Italic text (Text only). |
| `blur_radius` | Integer | `3` | Radius for blur (Obfuscate only). |
| `pixel_size` | Integer | `5` | Size of pixels (Obfuscate only). |

---

## 3. Export Step (`export`)

Saves or sends the final image to destinations.

| Key | Value | Description |
| :--- | :--- | :--- |
| `destinations` | Array of Objects | List of locations to send the image. |

### Destination Options

#### File
- `type`: `"file"`
- `path`: Full path to save the file. Supports `{timestamp}` placeholder (output format: `yyyyMMdd_HHmmss`).

#### Clipboard
- `type`: `"clipboard"`
- (No additional parameters)

---

## Full Example Script

```json
{
  "version": "1.0",
  "workflow": [
    {
      "step": "capture",
      "type": "fullscreen"
    },
    {
      "step": "annotate",
      "elements": [
        {
          "type": "arrow",
          "from": { "x": 100, "y": 100 },
          "to": { "x": 200, "y": 200 },
          "style": { "line_color": "Green", "line_thickness": 5 }
        },
        {
          "type": "obfuscate",
          "bounds": { "x": 300, "y": 300, "w": 200, "h": 100 },
          "style": { "blur_radius": 10 }
        }
      ]
    },
    {
      "step": "export",
      "destinations": [
        { "type": "file", "path": "C:\\Docs\\Capture_{timestamp}.png" },
        { "type": "clipboard" }
      ]
    }
  ]
}
```
