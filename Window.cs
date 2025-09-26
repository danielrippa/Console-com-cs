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

    public Window() {
    }

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
        IntPtr handle = GetStdOutHandle();
        var info = GetBufferInfo(handle);
        return info.srWindow.Right - info.srWindow.Left + 1;
      }
      set {
        IntPtr handle = GetStdOutHandle();
        var info = GetBufferInfo(handle);
        Resize(info.srWindow.Bottom - info.srWindow.Top + 1, value);
      }
    }

    public int Height {
      get {
        IntPtr handle = GetStdOutHandle();
        var info = GetBufferInfo(handle);
        return info.srWindow.Bottom - info.srWindow.Top + 1;
      }
      set {
        IntPtr handle = GetStdOutHandle();
        var info = GetBufferInfo(handle);
        Resize(value, info.srWindow.Right - info.srWindow.Left + 1);
      }
    }

    public bool Resize(int height, int width) {
      if (width <= 0 || height <= 0) { return false; }

      try {
        IntPtr handle = GetStdOutHandle();
        var info = GetBufferInfo(handle);

        if (width > info.dwMaximumWindowSize.X || height > info.dwMaximumWindowSize.Y) { return false; }

        short currentWidth = (short)(info.srWindow.Right - info.srWindow.Left + 1);
        short currentHeight = (short)(info.srWindow.Bottom - info.srWindow.Top + 1);

        var newSize = new COORD((short)width, (short)height);
        var newWindowRect = new SMALL_RECT { Left = 0, Top = 0, Right = (short)(width - 1), Bottom = (short)(height - 1) };

        if (width > currentWidth || height > currentHeight) {
          if (!SetConsoleScreenBufferSize(handle, newSize)) { return false; }
          if (!SetConsoleWindowInfo(handle, true, ref newWindowRect)) {
            SetConsoleScreenBufferSize(handle, info.dwSize);
            return false;
          }
        } else {
          if (!SetConsoleWindowInfo(handle, true, ref newWindowRect)) { return false; }
          if (!SetConsoleScreenBufferSize(handle, newSize)) {
            SetConsoleWindowInfo(handle, true, ref info.srWindow);
            SetConsoleScreenBufferSize(handle, info.dwSize);
            return false;
          }
        }

        return true;
      } catch (Exception) { return false; }
    }

    public bool ResizeScreenBufferToWindow(long screenBufferHandle) {
      IntPtr handle = new IntPtr(screenBufferHandle);

      int windowWidth = this.Width;
      int windowHeight = this.Height;

      if (windowWidth <= 0 || windowHeight <= 0) { return false; }

      COORD newBufferSize = new COORD((short)windowWidth, (short)windowHeight);
      return SetConsoleScreenBufferSize(handle, newBufferSize);
    }

    public int CodePage {
      get => (int)GetConsoleOutputCP();
      set => SetConsoleOutputCP((uint)value);
    }

  }

}
