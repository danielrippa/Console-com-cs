using System;
using System.Runtime.InteropServices;
using System.Text;

using static Win32.Kernel32;

namespace Console {

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.AutoDispatch)]
  [Guid("F8A2C56D-3E91-4B74-9D73-C78A5D763A15")]
  [ProgId("Console.Window")]

  public class Window {

    private const int MAX_TITLE_LENGTH = 255;
    private IntPtr handle = GetStdOutHandle();

    public string Title {
      get => GetTitle();
      set => SetTitle(value);
    }

    private string GetTitle() {
      StringBuilder titleBuffer = new StringBuilder(MAX_TITLE_LENGTH);
      GetConsoleTitle(titleBuffer, (uint)titleBuffer.Capacity);
      return titleBuffer.ToString();
    }

    private bool SetTitle(string title) {
      if (string.IsNullOrEmpty(title)) {
        return false;
      }
      return SetConsoleTitle(title);
    }

    public int Width {
      get {
        var info = GetBufferInfo(handle);
        return info.srWindow.Right - info.srWindow.Left + 1;
      }
      set {
        var info = GetBufferInfo(handle);
        Resize(value, info.srWindow.Bottom - info.srWindow.Top + 1);
      }
    }

    public int Height {
      get {
        var info = GetBufferInfo(handle);
        return info.srWindow.Bottom - info.srWindow.Top + 1;
      }
      set {
        var info = GetBufferInfo(handle);
        Resize(info.srWindow.Right - info.srWindow.Left + 1, value);
      }
    }

    public bool Resize(int width, int height) {

      if (width <= 0 || height <= 0) {
        return false;
      }

      try {
        IntPtr handle = GetStdOutHandle();
        var info = GetBufferInfo(handle);

        width = Math.Min(width, info.dwMaximumWindowSize.X);
        height = Math.Min(height, info.dwMaximumWindowSize.Y);

        var bufferSize = new COORD {
          X = (short)Math.Max(width, info.dwSize.X),
          Y = (short)Math.Max(height, info.dwSize.Y)
        };

        if (!SetConsoleScreenBufferSize(handle, bufferSize)) {
          return false;
        }

        var rect = new SMALL_RECT {
          Left = 0,
          Top = 0,
          Right = (short)(width - 1),
          Bottom = (short)(height - 1)
        };

        if (!SetConsoleWindowInfo(handle, true, ref rect)) {
          return false;
        }

        return true;
      } catch (Exception) {
        return false;
      }

    }

  }

}