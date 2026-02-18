using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace Greenshot.Macro
{
    public class MacroScript
    {
        [JsonProperty("version")]
        public string Version { get; set; }

        [JsonProperty("workflow")]
        public List<MacroStep> Workflow { get; set; }
    }

    public class MacroStep
    {
        [JsonProperty("step")]
        public string Step { get; set; } // capture, annotate, export

        [JsonProperty("type", NullValueHandling = NullValueHandling.Ignore)]
        public string Type { get; set; }
        
        // Capture properties
        [JsonProperty("area", NullValueHandling = NullValueHandling.Ignore)]
        public MacroArea Area { get; set; }

        [JsonProperty("path", NullValueHandling = NullValueHandling.Ignore)]
        public string Path { get; set; }

        [JsonProperty("options", NullValueHandling = NullValueHandling.Ignore)]
        public MacroOptions Options { get; set; }

        [JsonProperty("autocrop", NullValueHandling = NullValueHandling.Ignore)]
        public bool? AutoCrop { get; set; }

        [JsonProperty("autocrop_difference", NullValueHandling = NullValueHandling.Ignore)]
        public int? AutoCropDifference { get; set; }

        // Annotate properties
        [JsonProperty("elements", NullValueHandling = NullValueHandling.Ignore)]
        public List<MacroElement> Elements { get; set; }

        // Export properties
        [JsonProperty("destinations", NullValueHandling = NullValueHandling.Ignore)]
        public List<MacroDestination> Destinations { get; set; }
    }

    public class MacroArea
    {
        [JsonProperty("x")]
        public int X { get; set; }

        [JsonProperty("y")]
        public int Y { get; set; }

        [JsonProperty("width")]
        public int Width { get; set; }

        [JsonProperty("height")]
        public int Height { get; set; }
    }

    public class MacroOptions
    {
        [JsonProperty("show_cursor")]
        public bool ShowCursor { get; set; }

        [JsonProperty("delay_ms")]
        public int DelayMs { get; set; }
    }

    public class MacroElement
    {
        [JsonProperty("type")]
        public string Type { get; set; } // rectangle, arrow, text, obfuscate

        [JsonProperty("bounds", NullValueHandling = NullValueHandling.Ignore)]
        public MacroBounds Bounds { get; set; }

        [JsonProperty("from", NullValueHandling = NullValueHandling.Ignore)]
        public MacroPoint From { get; set; }

        [JsonProperty("to", NullValueHandling = NullValueHandling.Ignore)]
        public MacroPoint To { get; set; }

        [JsonProperty("position", NullValueHandling = NullValueHandling.Ignore)]
        public MacroPoint Position { get; set; }

        [JsonProperty("content", NullValueHandling = NullValueHandling.Ignore)]
        public string Content { get; set; }

        [JsonProperty("method", NullValueHandling = NullValueHandling.Ignore)]
        public string Method { get; set; } // for obfuscate

        [JsonProperty("pixel_size", NullValueHandling = NullValueHandling.Ignore)]
        public int? PixelSize { get; set; }

        [JsonProperty("style", NullValueHandling = NullValueHandling.Ignore)]
        public MacroStyle Style { get; set; }
    }

    public class MacroStyle
    {
        [JsonProperty("line_color", NullValueHandling = NullValueHandling.Ignore)]
        public string LineColor { get; set; }

        [JsonProperty("fill_color", NullValueHandling = NullValueHandling.Ignore)]
        public string FillColor { get; set; }

        [JsonProperty("line_thickness", NullValueHandling = NullValueHandling.Ignore)]
        public int? LineThickness { get; set; }

        [JsonProperty("shadow", NullValueHandling = NullValueHandling.Ignore)]
        public bool? Shadow { get; set; }

        [JsonProperty("font_size", NullValueHandling = NullValueHandling.Ignore)]
        public double? FontSize { get; set; }

        [JsonProperty("blur_radius", NullValueHandling = NullValueHandling.Ignore)]
        public int? BlurRadius { get; set; }

        [JsonProperty("pixel_size", NullValueHandling = NullValueHandling.Ignore)]
        public int? PixelSize { get; set; }

        public Dictionary<string, object> ToDictionary()
        {
            var dict = new Dictionary<string, object>();
            if (LineColor != null) dict["line_color"] = LineColor;
            if (FillColor != null) dict["fill_color"] = FillColor;
            if (LineThickness.HasValue) dict["line_thickness"] = LineThickness.Value;
            if (Shadow.HasValue) dict["shadow"] = Shadow.Value;
            if (FontSize.HasValue) dict["font_size"] = FontSize.Value;
            if (BlurRadius.HasValue) dict["blur_radius"] = BlurRadius.Value;
            if (PixelSize.HasValue) dict["pixel_size"] = PixelSize.Value;
            return dict;
        }
    }
    
    public class MacroBounds
    {
        [JsonProperty("x")]
        public int X { get; set; }

        [JsonProperty("y")]
        public int Y { get; set; }

        [JsonProperty("w")]
        public int W { get; set; }

        [JsonProperty("h")]
        public int H { get; set; }
    }

    public class MacroPoint
    {
        [JsonProperty("x")]
        public int X { get; set; }

        [JsonProperty("y")]
        public int Y { get; set; }
    }

    public class MacroDestination
    {
        [JsonProperty("type")]
        public string Type { get; set; } // file, clipboard

        [JsonProperty("path", NullValueHandling = NullValueHandling.Ignore)]
        public string Path { get; set; }

        [JsonProperty("overwrite", NullValueHandling = NullValueHandling.Ignore)]
        public bool? Overwrite { get; set; }
    }
}
