using System;
using System.Runtime.InteropServices;

using static Win32.Kernel32;

namespace Console {
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [Guid("C1D2E3F4-A5B6-7C8D-9E0F-1A2B3C4D5E6F")]
    [ProgId("Console.Palette")]
    public class Palette {

      public IntPtr Handle = GetStdOutHandle();

      private bool IsValidColorIndex(int index) {
        return index >= 0 && index <= 15;
      }

      private uint GetColorValue(int index) {
        return GetBufferInfoEx(Handle).ColorTable[index];
      }

      public bool SetColorRGB(int index, short r, short g, short b) {
        if (!IsValidColorIndex(index)) {
          return false;
        } else {
          var info = GetBufferInfoEx(Handle);
          info.ColorTable[index] = (uint)(((byte)r << 16) | ((byte)g << 8) | (byte)b);
          return SetBufferInfoEx(Handle, info);
        }
      }

      public string GetColorCsv(int index) {
        if (!IsValidColorIndex(index)) {
          return "";
        } else {
          uint color = GetBufferInfoEx(Handle).ColorTable[index];

          short r = (short)((color >> 16) & 0xFF);
          short g = (short)((color >> 8) & 0xFF);
          short b = (short)(color & 0xFF);

          return $"{r},{g},{b}";
        }
      }

      public string GetPaletteAsString() {
        var info = GetBufferInfoEx(Handle);
        var colors = new string[16];
        for (int i = 0; i < 16; i++) {
          uint color = info.ColorTable[i];
          short r = (short)((color >> 16) & 0xFF);
          short g = (short)((color >> 8) & 0xFF);
          short b = (short)(color & 0xFF);
          colors[i] = $"{r},{g},{b}";
        }
        return string.Join(";", colors);
      }

      public bool SetPaletteFromString(string paletteString) {
        if (string.IsNullOrEmpty(paletteString)) return false;
        var parts = paletteString.Split(';');
        if (parts.Length != 16) return false;

        var info = GetBufferInfoEx(Handle);
        for (int i = 0; i < 16; i++) {
          var rgb = parts[i].Split(',');
          if (rgb.Length != 3) return false;
          if (!short.TryParse(rgb[0], out short r)) return false;
          if (!short.TryParse(rgb[1], out short g)) return false;
          if (!short.TryParse(rgb[2], out short b)) return false;
          info.ColorTable[i] = (uint)(((byte)r << 16) | ((byte)g << 8) | (byte)b);
        }

        return SetBufferInfoEx(Handle, info);
      }

    }

}
