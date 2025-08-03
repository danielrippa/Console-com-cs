using System;
using System.Runtime.InteropServices;
using static Win32.Kernel32;

namespace Console {

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.AutoDispatch)]
  [Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890")]
  [ProgId("Console.Viewport")]

  public class Viewport {

    public IntPtr Handle {
      get => GetStdOutHandle();
    }

    public int Left {
      get => GetBufferInfo(Handle).srWindow.Left;
      set => MoveTo(Top, value);
    }

    public int Top {
      get => GetBufferInfo(Handle).srWindow.Top;
      set => MoveTo(value, Left);
    }

    public int Right {
      get => GetBufferInfo(Handle).srWindow.Right;
    }
    
    public int Bottom {
      get => GetBufferInfo(Handle).srWindow.Bottom;
    }

    public int Width {
      get => GetBufferInfo(Handle).srWindow.Right - GetBufferInfo(Handle).srWindow.Left + 1;
    }

    public int Height {
      get
      {
        var bufferInfo = GetBufferInfo(Handle);
        return bufferInfo.srWindow.Bottom - bufferInfo.srWindow.Top + 1;
      }
    }

    public int MaxWidth {
      get => GetBufferInfo(Handle).dwMaximumWindowSize.X;
    }
    
    public int MaxHeight {
      get => GetBufferInfo(Handle).dwMaximumWindowSize.Y;
    }

    public bool MoveTo(int row, int column) {
      try {
        var info = GetBufferInfo(Handle);
        
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

        return SetConsoleWindowInfo(Handle, true, ref rect);
      } catch (Exception) {
        return false;
      }
    }

    public bool ResizeToBuffer() {
      var info = GetBufferInfo(Handle);
      SMALL_RECT rect = new SMALL_RECT {
        Left = 0,
        Top = 0,
        Right = (short)(info.dwSize.X - 1),
        Bottom = (short)(info.dwSize.Y - 1)
      };
      return SetConsoleWindowInfo(Handle, true, ref rect);
    }
    
    public int GetLargestWidth() {
      return GetBufferInfo(Handle).dwMaximumWindowSize.X;
    }

    public int GetLargestHeight() {
      return GetBufferInfo(Handle).dwMaximumWindowSize.Y;
    }

  }

}