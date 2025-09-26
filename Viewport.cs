using System;
using System.Runtime.InteropServices;
using static Win32.Kernel32;

namespace Console {

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.AutoDispatch)]
  [Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890")]
  [ProgId("Console.Viewport")]

  public class Viewport {

    private IntPtr handle;

    public long Handle {
      get => handle.ToInt64();
      set => handle = new IntPtr(value);
    }

    public Viewport() {
      handle = GetStdOutHandle();
    }

    private SMALL_RECT Window {
      get => GetBufferInfo(handle).srWindow;
    }

    public int Left {
      get => Window.Left;
      set => MoveTo(Top, value);
    }

    public int Top {
      get => Window.Top;
      set => MoveTo(value, Left);
    }

    public int Right {
      get => Window.Right;
    }

    public int Bottom {
      get => Window.Bottom;
    }

    public int Width {
      get => Right - Left + 1;
    }

    public int Height {
      get => Bottom - Top + 1;
    }

    public int MaxWidth {
      get => GetBufferInfo(handle).dwMaximumWindowSize.X;
    }
    
    public int MaxHeight {
      get => GetBufferInfo(handle).dwMaximumWindowSize.Y;
    }

    public bool MoveTo(int row, int column) {
      try {
        var info = GetBufferInfo(handle);

        column = Math.Max(0, column);
        row = Math.Max(0, row);

        int maxLeft = Math.Max(0, info.dwSize.X - Width);
        int maxTop = Math.Max(0, info.dwSize.Y - Height);

        column = Math.Min(column, maxLeft);
        row = Math.Min(row, maxTop);

        SMALL_RECT rect = new SMALL_RECT {
          Left = (short)column,
          Top = (short)row,
          Right = (short)(column + Width - 1),
          Bottom = (short)(row + Height - 1)
        };

        return SetConsoleWindowInfo(handle, true, ref rect);
      } catch (Exception) {
        return false;
      }
    }

    public bool ResizeToBuffer() {
      var info = GetBufferInfo(handle);
      SMALL_RECT rect = new SMALL_RECT {
        Left = 0,
        Top = 0,
        Right = (short)(info.dwSize.X - 1),
        Bottom = (short)(info.dwSize.Y - 1)
      };
      return SetConsoleWindowInfo(handle, true, ref rect);
    }

    public bool Resize(int height, int width) {
      if (width <= 0 || height <= 0) {
        return false;
      }

      try {
        var info = GetBufferInfo(handle);

        width = Math.Min(width, info.dwMaximumWindowSize.X);
        height = Math.Min(height, info.dwMaximumWindowSize.Y);

        var rect = new SMALL_RECT {
          Left = info.srWindow.Left,
          Top = info.srWindow.Top,
          Right = (short)(info.srWindow.Left + width - 1),
          Bottom = (short)(info.srWindow.Top + height - 1)
        };

        return SetConsoleWindowInfo(handle, true, ref rect);
      } catch (Exception) {
        return false;
      }
    }

  }

}
