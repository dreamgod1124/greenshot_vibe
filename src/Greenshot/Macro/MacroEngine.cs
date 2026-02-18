using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;
using log4net;
using Greenshot.Base.Interfaces;
using Greenshot.Base.Interfaces.Drawing;
using Greenshot.Base.Core;
using Greenshot.Base.Core.Enums;
using Greenshot.Base.IniFile;
using Greenshot.Helpers;
using Greenshot.Editor.Drawing;
using Greenshot.Editor.Drawing.Fields;
using Dapplo.Windows.Common.Structs;
using Dapplo.Windows.User32;

namespace Greenshot.Macro
{
    public class MacroEngine
    {
        private static readonly ILog LOG = LogManager.GetLogger(typeof(MacroEngine));
        private static readonly CoreConfiguration CoreConfig = IniConfig.GetIniSection<CoreConfiguration>();
        private ICapture _currentCapture;
        private ISurface _currentSurface;

        public static void Run(string scriptPath)
        {
            try
            {
                if (!File.Exists(scriptPath))
                {
                    LOG.Error($"Macro script not found: {scriptPath}");
                    return;
                }

                LOG.Info($"Loading macro script from: {scriptPath}");
                string json = File.ReadAllText(scriptPath);
                MacroScript script = JsonConvert.DeserializeObject<MacroScript>(json);
                
                if (script == null)
                {
                    LOG.Error("Failed to deserialize macro script");
                    return;
                }

                MacroEngine engine = new MacroEngine();
                engine.Execute(script);
            }
            catch (Exception ex)
            {
                LOG.Error("Error while running macro script", ex);
            }
        }

        private void Execute(MacroScript script)
        {
            LOG.Info($"Executing workflow with {script.Workflow?.Count ?? 0} steps");
            if (script.Workflow == null) return;

            foreach (var step in script.Workflow)
            {
                try
                {
                    LOG.Info($"Step: {step.Step} ({step.Type})");
                    switch (step.Step.ToLower())
                    {
                        case "capture":
                            ExecuteCapture(step);
                            break;
                        case "annotate":
                            ExecuteAnnotate(step);
                            break;
                        case "export":
                            ExecuteExport(step);
                            break;
                        default:
                            LOG.Warn($"Unsupported step: {step.Step}");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    LOG.Error($"Error in step {step.Step}", ex);
                    break; // stop on error for safety
                }
            }
        }

        private void ExecuteCapture(MacroStep step)
        {
            LOG.Info($"Capturing with type: {step.Type}");
            
            _currentCapture = new Capture();
            _currentCapture.CaptureDetails.DateTime = DateTime.Now;
            
            NativeRect captureBounds = NativeRect.Empty;

            switch (step.Type?.ToLower())
            {
                case "region":
                    if (step.Area != null) {
                        captureBounds = new NativeRect(step.Area.X, step.Area.Y, step.Area.Width, step.Area.Height);
                    } else {
                        throw new Exception("Capture type 'region' requires 'area' property.");
                    }
                    break;
                case "fullscreen":
                    captureBounds = DisplayInfo.ScreenBounds;
                    break;
                case "file":
                    if (string.IsNullOrEmpty(step.Path) || !File.Exists(step.Path))
                    {
                        throw new Exception($"Capture file not found: {step.Path}");
                    }
                    using (Image loadedImage = Image.FromFile(step.Path))
                    {
                        // Clone the image to avoid locking the file
                        _currentCapture = new Capture(new Bitmap(loadedImage));
                        _currentCapture.CaptureDetails.DateTime = DateTime.Now;
                        _currentCapture.CaptureDetails.Title = Path.GetFileName(step.Path);
                        _currentSurface = new Surface(_currentCapture);
                        LOG.Info($"Loaded image from file: {step.Path}");
                    }
                    return; // skip the rest of the method as we already have the capture
                default:
                    throw new Exception($"Unsupported capture type: {step.Type}");
            }
            
            if (!captureBounds.IsEmpty)
            {
                _currentCapture = WindowCapture.CaptureRectangle(_currentCapture, captureBounds);
            }
            
            if (_currentCapture?.Image != null)
            {
                 // Handle AutoCrop
                 bool autoCrop = step.AutoCrop ?? CoreConfig.AutoCrop;
                 if (autoCrop)
                 {
                     int difference = step.AutoCropDifference ?? CoreConfig.AutoCropDifference;
                     var cropRect = ImageHelper.FindAutoCropRectangle(_currentCapture.Image, difference);
                     if (!cropRect.IsEmpty)
                     {
                         _currentCapture.Crop(cropRect);
                         LOG.Info("AutoCrop applied.");
                     }
                 }

                 _currentSurface = new Surface(_currentCapture);
                 LOG.Info("Capture created and surface initialized.");
            }
            else
            {
                throw new Exception("Capture failed: No image acquired.");
            }
        }

        private void ExecuteAnnotate(MacroStep step)
        {
            if (_currentSurface == null) throw new Exception("No active surface to annotate. Run 'capture' first.");
            if (step.Elements == null) return;

            foreach (var element in step.Elements)
            {
                IDrawableContainer container = null;
                switch (element.Type.ToLower())
                {
                    case "rectangle":
                        container = new RectangleContainer(_currentSurface);
                        if (element.Bounds != null)
                        {
                            container.Left = element.Bounds.X;
                            container.Top = element.Bounds.Y;
                            container.Width = element.Bounds.W;
                            container.Height = element.Bounds.H;
                        }
                        break;
                    case "arrow":
                        container = new ArrowContainer(_currentSurface);
                        if (element.From != null && element.To != null)
                        {
                            // In Greenshot LineContainer, Left/Top is start, Width/Height is delta
                            container.Left = element.From.X;
                            container.Top = element.From.Y;
                            container.Width = element.To.X - element.From.X;
                            container.Height = element.To.Y - element.From.Y;
                        }
                        break;
                    case "text":
                        var textContainer = new TextContainer(_currentSurface);
                        if (element.Position != null)
                        {
                            textContainer.Left = element.Position.X;
                            textContainer.Top = element.Position.Y;
                        }
                        textContainer.Text = element.Content ?? "";
                        container = textContainer;
                        break;
                    case "obfuscate":
                        container = new ObfuscateContainer(_currentSurface);
                        if (element.Bounds != null)
                        {
                            container.Left = element.Bounds.X;
                            container.Top = element.Bounds.Y;
                            container.Width = element.Bounds.W;
                            container.Height = element.Bounds.H;
                        }
                        break;
                    default:
                        LOG.Warn($"Unsupported element type: {element.Type}");
                        continue;
                }

                if (container != null)
                {
                    ApplyStyle(container, element.Style);
                    // For text, we might need FitToText
                    if (container is TextContainer tc) tc.FitToText();
                    
                    _currentSurface.AddElement(container);
                }
            }
        }

        private void ApplyStyle(IDrawableContainer container, MacroStyle style)
        {
            if (style == null) return;
            var dict = style.ToDictionary();
            
            // Handle ObfuscateContainer preset switching
            if (container is ObfuscateContainer obfuscate)
            {
                if (dict.ContainsKey("blur_radius") || dict.ContainsKey("blurradius"))
                {
                    obfuscate.SetFieldValue(FieldType.PREPARED_FILTER_OBFUSCATE, FilterContainer.PreparedFilter.BLUR);
                }
                else if (dict.ContainsKey("pixel_size") || dict.ContainsKey("pixelsize"))
                {
                    obfuscate.SetFieldValue(FieldType.PREPARED_FILTER_OBFUSCATE, FilterContainer.PreparedFilter.PIXELIZE);
                }
            }

            foreach (var pair in dict)
            {
                string key = pair.Key.ToUpper().Replace("_", "");
                IFieldType fieldType = FieldType.Values.FirstOrDefault(f => f.Name.Replace("_", "") == key);
                
                if (fieldType != null)
                {
                    // Use AbstractFieldHolderWithChildren if possible to avoid interface shadowing issues in Greenshot
                    // where HasField/GetField are 'new' instead of 'override'
                    AbstractFieldHolderWithChildren holder = container as AbstractFieldHolderWithChildren;
                    
                    if (holder != null)
                    {
                        if (!holder.HasField(fieldType))
                        {
                            LOG.Warn($"Field '{fieldType.Name}' is not supported by {container.GetType().Name} or its children.");
                            continue;
                        }

                        object value = pair.Value;
                        if (fieldType == FieldType.LINE_COLOR || fieldType == FieldType.FILL_COLOR || fieldType == FieldType.HIGHLIGHT_COLOR)
                        {
                            value = ParseColor(value);
                        }
                        else if (fieldType == FieldType.LINE_THICKNESS || fieldType == FieldType.PIXEL_SIZE || fieldType == FieldType.MAGNIFICATION_FACTOR || fieldType == FieldType.BLUR_RADIUS)
                        {
                            value = Convert.ToInt32(value);
                        }
                        else if (fieldType == FieldType.SHADOW || fieldType == FieldType.FONT_BOLD || fieldType == FieldType.FONT_ITALIC)
                        {
                            value = Convert.ToBoolean(value);
                        }

                        try
                        {
                            holder.SetFieldValue(fieldType, value);
                        }
                        catch (Exception ex)
                        {
                            LOG.Error($"Failed to set field {fieldType.Name} on {container.GetType().Name}", ex);
                        }
                    }
                }
                else
                {
                    LOG.Warn($"Unknown style property: {pair.Key}");
                }
            }
        }

        private Color ParseColor(object value)
        {
            if (value == null) return Color.Transparent;
            string s = value.ToString();
            if (s.ToLower() == "transparent") return Color.Transparent;
            try
            {
                return ColorTranslator.FromHtml(s);
            }
            catch
            {
                return Color.Red;
            }
        }

        private void ExecuteExport(MacroStep step)
        {
            if (_currentSurface == null) throw new Exception("No active surface to export.");
            if (step.Destinations == null) return;

            foreach (var destination in step.Destinations)
            {
                LOG.Info($"Exporting to: {destination.Type}");
                switch (destination.Type.ToLower())
                {
                    case "file":
                        string path = destination.Path.Replace("{timestamp}", DateTime.Now.ToString("yyyyMMdd_HHmmss"));
                        // Create directory if not exists
                        string dir = Path.GetDirectoryName(path);
                        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);

                        using (Image image = _currentSurface.GetImageForExport())
                        {
                            image.Save(path);
                        }
                        LOG.Info($"Saved to {path}");
                        break;
                    case "clipboard":
                        using (Image image = _currentSurface.GetImageForExport())
                        {
                            Clipboard.SetImage(image);
                        }
                        LOG.Info("Copied to clipboard");
                        break;
                    default:
                        LOG.Warn($"Unsupported destination type: {destination.Type}");
                        break;
                }
            }
        }
    }
}
