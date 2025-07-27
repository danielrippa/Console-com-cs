using System;
using System.Runtime.InteropServices;
using static Win32.Kernel32;

namespace Console {

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.AutoDispatch)]
  [Guid("17982804-7649-47E9-9787-805F29744287")]
  [ProgId("Console.TextAttributes")]

  public class TextAttributes {

    public short Value;
    public IntPtr Handle { get; set; }

    public TextAttributes(short value = 0) {
      Value = value;
    }

    public TextColor Ink {
      get => GetTextColor(Value);
      set {
        Value = ApplyTextColor(Value, value);
        ApplyToConsoleScreenBufferInfo();
      }
    }

    public TextColor Paper {
      get => GetTextColor(Value);
      set {
        Value = ApplyTextColor(Value, value);
        ApplyToConsoleScreenBufferInfo();
      }
    }

    public TextBorders Borders {
      get => GetTextBorders(Value);
      set {
        Value = ApplyTextBorders(Value, value);
        ApplyToConsoleScreenBufferInfo();
      }
    }

    public bool? Inverted {
      get => GetInverse(Value);
      set {
        Value = ApplyInverse(Value, value ?? false);
        ApplyToConsoleScreenBufferInfo();
      }
    }

    public void Apply(TextAttributes attributes) {
      if (attributes.Ink != null) Ink = attributes.Ink;
      if (attributes.Paper != null) Paper = attributes.Paper;
      if (attributes.Borders != null) Borders = attributes.Borders;

      if (attributes.Inverted.HasValue) Inverted = attributes.Inverted.Value;
    }

    private void ApplyToConsoleScreenBufferInfo() {
      if (Handle != IntPtr.Zero) {
        SetConsoleTextAttribute(Handle, (ushort)Value);
      }
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [Guid("A1928701-43A6-4C67-8286-1E8A3023A8EB")]
    [ProgId("WinCon.TextBorders")]
    public class TextBorders {

      public bool? Top { get; set; }
      public bool? Left { get; set; }
      public bool? Bottom { get; set; }
      public bool? Right { get; set; }

      public TextBorders(bool? top = null, bool? left = null, bool? bottom = null, bool? right = null) {
        Top = top;
        Left = left;
        Bottom = bottom;
        Right = right;
      }

      public static readonly TextBorders None = new TextBorders();
      public static readonly TextBorders TopBorder = new TextBorders(top: true);
      public static readonly TextBorders Sides = new TextBorders(left: true, right: true);
      public static readonly TextBorders Underlined = new TextBorders(bottom: true);

      public static readonly TextBorders NW = new TextBorders(top: true, left: true);
      public static readonly TextBorders NE = new TextBorders(top: true, right: true);
      public static readonly TextBorders SE = new TextBorders(bottom: true, right: true);
      public static readonly TextBorders SW = new TextBorders(bottom: true, left: true);

      public static readonly TextBorders All = new TextBorders(top: true, left: true, bottom: true, right: true);

    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [Guid("65F7D943-FBB2-4942-B4E8-BD355CEDE334")]
    [ProgId("WinCon.TextColor")]
    public class TextColor {

      public bool? Red { get; set; }
      public bool? Green { get; set; }
      public bool? Blue { get; set; }
      public bool? Intensity { get; set; }

      public TextColor(bool? red = null, bool? green = null, bool? blue = null, bool? intensity = null) {
        Red = red;
        Green = green;
        Blue = blue;
        Intensity = intensity;
      }

    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [Guid("D4E1F3A2-5B8B-4F8A-9B6E-8C6B3A8D4F3A")]
    [ProgId("WinCon.TextColors")]
    public class TextColors {

      public TextColor Black { get; } = new TextColor();
      public TextColor Red { get; } = new TextColor(red: true);
      public TextColor Green { get; } = new TextColor(green: true);
      public TextColor Navy { get; } = new TextColor(blue: true);

      public TextColor Celestial { get; } = new TextColor(green: true, blue: true);
      public TextColor Curry { get; } = new TextColor(red: true, green: true);
      public TextColor Purple { get; } = new TextColor(red: true, blue: true);
      public TextColor Gray { get; } = new TextColor(red: true, green: true, blue: true);

      public TextColor Silver { get; } = new TextColor(intensity: true);
      public TextColor Vermillion { get; } = new TextColor(red: true, intensity: true);
      public TextColor Emerald { get; } = new TextColor(green: true, intensity: true);
      public TextColor Blue { get; } = new TextColor(blue: true, intensity: true);
      public TextColor Turquoise { get; } = new TextColor(green: true, blue: true, intensity: true);
      public TextColor Yellow { get; } = new TextColor(red: true, green: true, intensity: true);
      public TextColor Fuchsia { get; } = new TextColor(red: true, blue: true, intensity: true);
      public TextColor White { get; } = new TextColor(red: true, green: true, blue: true, intensity: true);

    }

    private static short SetBit(short value, int bit, bool enabled) {
      short mask = (short)(1 << bit);
      return enabled ? (short)(value | mask) : (short)(value & ~mask);
    }

    private static short ApplyTextColor(short value, TextColor color, int offset = 0) {
      if (color.Intensity.HasValue) value = SetBit(value, 3 + offset, color.Intensity.Value);

      if (color.Red.HasValue) value = SetBit(value, 2 + offset, color.Red.Value);
      if (color.Green.HasValue) value = SetBit(value, 1 + offset, color.Green.Value);
      if (color.Blue.HasValue) value = SetBit(value, 0 + offset, color.Blue.Value);

      return value;
    }

    private static short ApplyTextBorders(short value, TextBorders borders) {
      if (borders.Top.HasValue) value = SetBit(value, 10, borders.Top.Value);
      if (borders.Left.HasValue) value = SetBit(value, 11, borders.Left.Value);
      if (borders.Bottom.HasValue) value = SetBit(value, 12, borders.Bottom.Value);
      if (borders.Right.HasValue) value = SetBit(value, 15, borders.Right.Value);

      return value;
    }

    private static short ApplyInverse(short value, bool inverse) {
      return SetBit(value, 14, inverse);
    }

    private static TextColor GetTextColor(short value, int offset = 0) {
      return new TextColor {
        Intensity = (value & (1 << (3 + offset))) != 0,

        Red = (value & (1 << (2 + offset))) != 0,
        Green = (value & (1 << (1 + offset))) != 0,
        Blue = (value & (1 << (0 + offset))) != 0
      };
    }

    private static TextBorders GetTextBorders(short value) {
      return new TextBorders {
        Top = (value & (1 << 10)) != 0,
        Left = (value & (1 << 11)) != 0,
        Right = (value & (1 << 12)) != 0,
        Bottom = (value & (1 << 15)) != 0
      };
    }

    private bool? GetInverse(short value) {
      return (value & (1 << 14)) != 0;
    }

  }

}