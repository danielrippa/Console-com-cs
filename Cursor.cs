using System;
using System.Runtime.InteropServices;
using static Win32.Kernel32;

namespace Console {

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.AutoDispatch)]
  [Guid("AE72DC0B-E2EC-42B4-B95E-0A1171897413")]
  [ProgId("Console.Cursor")]

  public class Cursor {

    private IntPtr handle;

    public long Handle {
      get => handle.ToInt64();
      set => handle = new IntPtr(value);
    }

    public Cursor() {
      handle = GetStdOutHandle();
    }

    public void Goto(int row, int column) {
      var position = new COORD { X = (short)column, Y = (short)row };
      SetConsoleCursorPosition(handle, position);
    }

    private COORD GetPosition() {
      var info = GetBufferInfo(handle);
      return info.dwCursorPosition;
    }

    public int Row {
      get => GetPosition().Y;
      set => Goto(value, Column);
    }

    public int Column {
      get => GetPosition().X;
      set => Goto(Row, value);
    }

    private CONSOLE_CURSOR_INFO GetCursorInfo() {
      GetConsoleCursorInfo(handle, out var info);
      return info;
    }

    private void SetCursorInfo(ref CONSOLE_CURSOR_INFO info) {
      SetConsoleCursorInfo(handle, ref info);
    }

    public int Size {
      get => (int)GetCursorInfo().dwSize;
      set {
        var info = GetCursorInfo();
        info.dwSize = (uint)value;
        SetCursorInfo(ref info);
      }

    }

    public bool Visible {
      get => GetCursorInfo().bVisible;
      set {
        var info = GetCursorInfo();
        info.bVisible = value;
        SetCursorInfo(ref info);
      }
    }

  }

}