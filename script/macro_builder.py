import sys
import json
import os
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QListWidget, QComboBox, QFormLayout, QLineEdit, 
                             QSpinBox, QDoubleSpinBox, QCheckBox, QLabel, QTextEdit, 
                             QColorDialog, QFileDialog, QGroupBox, QScrollArea)
from PyQt6.QtCore import Qt

class MacroBuilder(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Greenshot Macro Builder")
        self.setMinimumSize(1000, 700)
        
        self.steps = []
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # --- Left Panel: Steps List ---
        left_panel = QVBoxLayout()
        left_panel.addWidget(QLabel("Workflow Steps:"))
        
        self.steps_list = QListWidget()
        self.steps_list.currentRowChanged.connect(self.load_step_properties)
        left_panel.addWidget(self.steps_list)

        btn_layout = QHBoxLayout()
        add_btn = QPushButton("Add Step")
        add_btn.clicked.connect(self.add_step_dialog)
        remove_btn = QPushButton("Remove Step")
        remove_btn.clicked.connect(self.remove_step)
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(remove_btn)
        left_panel.addLayout(btn_layout)
        
        main_layout.addLayout(left_panel, 1)

        # --- Middle Panel: Properties ---
        self.prop_scroll = QScrollArea()
        self.prop_scroll.setWidgetResizable(True)
        self.prop_container = QWidget()
        self.prop_layout = QVBoxLayout(self.prop_container)
        self.prop_scroll.setWidget(self.prop_container)
        
        main_layout.addWidget(self.prop_scroll, 2)

        # --- Right Panel: JSON Preview ---
        right_panel = QVBoxLayout()
        right_panel.addWidget(QLabel("JSON Preview:"))
        self.json_preview = QTextEdit()
        self.json_preview.setReadOnly(True)
        right_panel.addWidget(self.json_preview)

        action_layout = QHBoxLayout()
        load_btn = QPushButton("Load JSON")
        load_btn.clicked.connect(self.load_json)
        save_btn = QPushButton("Save JSON")
        save_btn.clicked.connect(self.save_json)
        run_btn = QPushButton("Run in Greenshot")
        run_btn.clicked.connect(self.run_macro)
        action_layout.addWidget(load_btn)
        action_layout.addWidget(save_btn)
        action_layout.addWidget(run_btn)
        right_panel.addLayout(action_layout)

        main_layout.addLayout(right_panel, 2)

    def add_step_dialog(self):
        # Default to capture
        step = {"step": "capture", "type": "fullscreen"}
        self.steps.append(step)
        self.steps_list.addItem(f"{len(self.steps)}. {step['step']} ({step['type']})")
        self.steps_list.setCurrentRow(len(self.steps) - 1)
        self.update_preview()

    def remove_step(self):
        row = self.steps_list.currentRow()
        if row >= 0:
            self.steps.pop(row)
            self.steps_list.takeItem(row)
            self.update_preview()
            self.clear_properties()

    def clear_layout(self, layout):
        if layout is not None:
            while layout.count():
                item = layout.takeAt(0)
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()
                else:
                    self.clear_layout(item.layout())

    def clear_properties(self):
        self.clear_layout(self.prop_layout)

    def load_step_properties(self, index):
        if index < 0 or index >= len(self.steps):
            self.clear_properties()
            return

        # Optimization: if selecting same step, don't always clear everything?
        # For now, keep it simple but fix the crashing behavior.
        self.clear_properties()
        step = self.steps[index]
        
        # Step Type Selector
        type_group = QGroupBox("Step Type")
        type_form = QFormLayout()
        step_combo = QComboBox()
        step_combo.addItems(["capture", "annotate", "export"])
        step_combo.setCurrentText(step["step"])
        # Use lambda with default argument to capture current index
        step_combo.currentTextChanged.connect(lambda v, idx=index: self.update_step_base(idx, "step", v))
        type_form.addRow("Step:", step_combo)
        type_group.setLayout(type_form)
        self.prop_layout.addWidget(type_group)

        if step["step"] == "capture":
            self.build_capture_props(index, step)
        elif step["step"] == "annotate":
            self.build_annotate_props(index, step)
        elif step["step"] == "export":
            self.build_export_props(index, step)

        self.prop_layout.addStretch()

    def update_step_base(self, index, key, value):
        if index >= len(self.steps): return
        if self.steps[index].get(key) == value: return
        
        self.steps[index][key] = value
        # Reset specific fields if step type changed
        if key == "step":
            if value == "capture": self.steps[index] = {"step": "capture", "type": "fullscreen"}
            elif value == "annotate": self.steps[index] = {"step": "annotate", "elements": []}
            elif value == "export": self.steps[index] = {"step": "export", "destinations": []}
            # Defer reloading to avoid signal-processing issues
            from PyQt6.QtCore import QTimer
            QTimer.singleShot(0, lambda: self.load_step_properties(index))
        
        self.steps_list.item(index).setText(f"{index+1}. {self.steps[index].get('step')} ({self.steps[index].get('type', '')})")
        self.update_preview()

    def build_capture_props(self, index, step):
        cap_group = QGroupBox("Capture Settings")
        form = QFormLayout()
        
        type_combo = QComboBox()
        type_combo.addItems(["fullscreen", "region", "file"])
        type_combo.setCurrentText(step.get("type", "fullscreen"))
        type_combo.currentTextChanged.connect(lambda v: self.update_capture_type(index, v))
        form.addRow("Mode:", type_combo)
        
        if step.get("type") == "region":
            area = step.get("area", {"X": 0, "Y": 0, "Width": 800, "Height": 600})
            step["area"] = area
            for k in ["X", "Y", "Width", "Height"]:
                sb = QSpinBox()
                sb.setRange(0, 5000)
                sb.setValue(area[k])
                sb.valueChanged.connect(lambda v, key=k: self.update_dict(area, key, v))
                form.addRow(f"{k}:", sb)
        elif step.get("type") == "file":
            row = QHBoxLayout()
            le = QLineEdit(step.get("path", ""))
            le.textChanged.connect(lambda v: self.update_dict(step, "path", v))
            btn = QPushButton("Browse")
            btn.clicked.connect(lambda: self.browse_capture_file(le))
            row.addWidget(le)
            row.addWidget(btn)
            form.addRow("Image Path:", row)

        # Autocrop
        form.addRow(QLabel("<b>Post-Capture:</b>"))
        ac_cb = QCheckBox("Enable Autocrop")
        ac_cb.setChecked(step.get("autocrop", False))
        ac_cb.toggled.connect(lambda v: self.update_autocrop(step, v))
        form.addRow(ac_cb)
        
        diff_sb = QSpinBox()
        diff_sb.setRange(0, 255)
        diff_sb.setValue(step.get("autocrop_difference", 10))
        diff_sb.valueChanged.connect(lambda v: self.update_dict(step, "autocrop_difference", v))
        form.addRow("Autocrop Diff:", diff_sb)

        cap_group.setLayout(form)
        self.prop_layout.addWidget(cap_group)

    def browse_capture_file(self, line_edit):
        path, _ = QFileDialog.getOpenFileName(self, "Select Image File", "", "Images (*.png *.jpg *.jpeg *.bmp)")
        if path:
            line_edit.setText(path)

    def update_capture_type(self, index, val):
        self.steps[index]["type"] = val
        if val != "region": self.steps[index].pop("area", None)
        if val != "file": self.steps[index].pop("path", None)
        self.load_step_properties(index)
        self.update_preview()

    def build_annotate_props(self, index, step):
        ann_group = QGroupBox("Annotations")
        vbox = QVBoxLayout()
        
        elements_list = QListWidget()
        for i, el in enumerate(step.get("elements", [])):
            elements_list.addItem(f"{i+1}. {el['type']}")
        
        # Area for editing the selected element
        self.el_prop_container = QWidget()
        self.el_prop_layout = QVBoxLayout(self.el_prop_container)
        
        elements_list.currentRowChanged.connect(lambda row, idx=index: self.load_element_properties(idx, row))
        
        vbox.addWidget(QLabel("Elements:"))
        vbox.addWidget(elements_list)
        
        btn_layout = QHBoxLayout()
        add_el_combo = QComboBox()
        add_el_combo.addItems(["rectangle", "arrow", "text", "obfuscate"])
        add_btn = QPushButton("Add Element")
        add_btn.clicked.connect(lambda: self.add_element(index, add_el_combo.currentText()))
        
        remove_btn = QPushButton("Remove Selected")
        remove_btn.clicked.connect(lambda: self.remove_element(index, elements_list.currentRow()))
        
        btn_layout.addWidget(add_el_combo)
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(remove_btn)
        vbox.addLayout(btn_layout)
        
        vbox.addWidget(self.el_prop_container)
        
        ann_group.setLayout(vbox)
        self.prop_layout.addWidget(ann_group)

    def load_element_properties(self, step_idx, el_idx):
        self.clear_layout(self.el_prop_layout)
        if el_idx < 0 or el_idx >= len(self.steps[step_idx]["elements"]):
            return
            
        el = self.steps[step_idx]["elements"][el_idx]
        form = QFormLayout()
        
        # --- Bounds/Coordinates ---
        if "bounds" in el:
            for k in ["x", "y", "w", "h"]:
                sb = QSpinBox()
                sb.setRange(-5000, 5000)
                sb.setValue(el["bounds"].get(k, 0))
                sb.valueChanged.connect(lambda v, key=k: self.update_dict(el["bounds"], key, v))
                form.addRow(f"Bounds {k.upper()}:", sb)
        elif "from" in el:
            for pt in ["from", "to"]:
                for k in ["x", "y"]:
                    sb = QSpinBox()
                    sb.setRange(-5000, 5000)
                    sb.setValue(el[pt].get(k, 0))
                    sb.valueChanged.connect(lambda v, p=pt, key=k: self.update_dict(el[p], key, v))
                    form.addRow(f"{pt.capitalize()} {k.upper()}:", sb)
        elif "position" in el:
            for k in ["x", "y"]:
                sb = QSpinBox()
                sb.setRange(-5000, 5000)
                sb.setValue(el["position"].get(k, 0))
                sb.valueChanged.connect(lambda v, key=k: self.update_dict(el["position"], key, v))
                form.addRow(f"Position {k.upper()}:", sb)

        # --- Content ---
        if "content" in el:
            le = QLineEdit(el.get("content", ""))
            le.textChanged.connect(lambda v: self.update_dict(el, "content", v))
            form.addRow("Text Content:", le)

        # --- Styles ---
        style_title = QLabel("<b>Style Settings:</b>")
        form.addRow(style_title)
        
        style = el.get("style", {})
        el["style"] = style # ensure it exists
        
        # Colors
        color_props = [("Line Color", "line_color"), ("Fill Color", "fill_color")]
        for label, key in color_props:
            row = QHBoxLayout()
            val = style.get(key, "#FF0000" if "line" in key else "transparent")
            le = QLineEdit(str(val))
            le.textChanged.connect(lambda v, k=key: self.update_dict(style, k, v))
            btn = QPushButton("Pick")
            btn.clicked.connect(lambda checked, l=le: self.pick_color(l))
            row.addWidget(le)
            row.addWidget(btn)
            form.addRow(f"{label}:", row)

        # Numeric Styles
        thick_sb = QSpinBox()
        thick_sb.setRange(0, 50)
        thick_sb.setValue(style.get("line_thickness", 3))
        thick_sb.valueChanged.connect(lambda v: self.update_dict(style, "line_thickness", v))
        form.addRow("Line Thickness:", thick_sb)

        if el["type"] == "obfuscate":
            br_sb = QSpinBox()
            br_sb.setValue(style.get("blur_radius", 5))
            br_sb.valueChanged.connect(lambda v: self.update_dict(style, "blur_radius", v))
            form.addRow("Blur Radius:", br_sb)
            
            ps_sb = QSpinBox()
            ps_sb.setValue(style.get("pixel_size", 5))
            ps_sb.valueChanged.connect(lambda v: self.update_dict(style, "pixel_size", v))
            form.addRow("Pixel Size:", ps_sb)
        
        if el["type"] == "text":
            fs_sb = QDoubleSpinBox()
            fs_sb.setValue(style.get("font_size", 12))
            fs_sb.valueChanged.connect(lambda v: self.update_dict(style, "font_size", v))
            form.addRow("Font Size:", fs_sb)

        shadow_cb = QCheckBox("Drop Shadow")
        shadow_cb.setChecked(style.get("shadow", False))
        shadow_cb.toggled.connect(lambda v: self.update_dict(style, "shadow", v))
        form.addRow(shadow_cb)

        self.el_prop_layout.addLayout(form)

    def pick_color(self, line_edit):
        color = QColorDialog.getColor()
        if color.isValid():
            line_edit.setText(color.name())

    def remove_element(self, step_idx, el_idx):
        if el_idx >= 0:
            self.steps[step_idx]["elements"].pop(el_idx)
            self.load_step_properties(step_idx)
            self.update_preview()

    def add_element(self, step_idx, el_type):
        el = {"type": el_type, "style": {}}
        if el_type == "rectangle": el["bounds"] = {"x": 100, "y": 100, "w": 200, "h": 150}
        elif el_type == "arrow": 
            el["from"] = {"x": 100, "y": 100}
            el["to"] = {"x": 200, "y": 200}
        elif el_type == "text":
            el["position"] = {"x": 100, "y": 100}
            el["content"] = "New Text"
        elif el_type == "obfuscate":
            el["bounds"] = {"x": 100, "y": 100, "w": 200, "h": 150}
            el["style"]["blur_radius"] = 5
            
        self.steps[step_idx]["elements"].append(el)
        self.load_step_properties(step_idx)
        self.update_preview()

    def build_export_props(self, index, step):
        exp_group = QGroupBox("Export Destinations")
        vbox = QVBoxLayout()
        
        for dest in step.get("destinations", []):
            label = QLabel(f"Dest: {dest['type']}")
            vbox.addWidget(label)
            if dest["type"] == "file":
                vbox.addWidget(QLabel(f"Path: {dest.get('path', '')}"))

        btn_layout = QHBoxLayout()
        add_file_btn = QPushButton("+ File")
        add_file_btn.clicked.connect(lambda: self.add_destination(index, "file"))
        add_clip_btn = QPushButton("+ Clipboard")
        add_clip_btn.clicked.connect(lambda: self.add_destination(index, "clipboard"))
        btn_layout.addWidget(add_file_btn)
        btn_layout.addWidget(add_clip_btn)
        
        vbox.addLayout(btn_layout)
        exp_group.setLayout(vbox)
        self.prop_layout.addWidget(exp_group)

    def add_destination(self, step_idx, dest_type):
        dest = {"type": dest_type}
        if dest_type == "file":
            dest["path"] = QFileDialog.getSaveFileName(self, "Select Export Path", "", "Images (*.png *.jpg)")[0]
            if not dest["path"]: return
        
        self.steps[step_idx]["destinations"].append(dest)
        self.load_step_properties(step_idx)
        self.update_preview()

    def update_autocrop(self, step_dict, enabled):
        step_dict["autocrop"] = enabled
        if not enabled:
            if "autocrop_difference" in step_dict:
                del step_dict["autocrop_difference"]
        else:
            # Set a default if it doesn't exist when enabled
            if "autocrop_difference" not in step_dict:
                step_dict["autocrop_difference"] = 10
        self.update_preview()

    def update_dict(self, d, key, val):
        d[key] = val
        self.update_preview()

    def update_preview(self):
        data = {
            "version": "1.0",
            "workflow": self.steps
        }
        self.json_preview.setText(json.dumps(data, indent=2))

    def load_json(self):
        path, _ = QFileDialog.getOpenFileName(self, "Load Macro JSON", "", "JSON Files (*.json)")
        if path:
            try:
                with open(path, 'r') as f:
                    data = json.load(f)
                    self.steps = data.get("workflow", [])
                    self.steps_list.clear()
                    for i, step in enumerate(self.steps):
                        self.steps_list.addItem(f"{i+1}. {step.get('step')} ({step.get('type', '')})")
                    self.update_preview()
                    self.clear_properties()
            except Exception as e:
                from PyQt6.QtWidgets import QMessageBox
                QMessageBox.critical(self, "Error", f"Failed to load JSON: {str(e)}")

    def save_json(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Macro JSON", "", "JSON Files (*.json)")
        if path:
            with open(path, 'w') as f:
                f.write(self.json_preview.toPlainText())

    def run_macro(self):
        # Save temp file
        temp_path = os.path.join(os.getcwd(), "temp_macro.json")
        with open(temp_path, 'w') as f:
            f.write(self.json_preview.toPlainText())
        
        # Try to find Greenshot
        gs_path = r"D:\Program Files\Greenshot\Greenshot.exe" # Common path
        if not os.path.exists(gs_path):
            gs_path, _ = QFileDialog.getOpenFileName(self, "Locate Greenshot.exe", "C:\\", "Executable (*.exe)")
        
        if gs_path:
            import subprocess
            subprocess.Popen([gs_path, "/macro", temp_path])

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MacroBuilder()
    window.show()
    sys.exit(app.exec())
