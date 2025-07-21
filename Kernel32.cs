using System;
using System.Runtime.InteropServices;

namespace Console {

  internal static class Kernel32 {

    private const string Dll = "kernel32.dll";

    internal const int STD_OUTPUT_HANDLE = -11;
    internal const int STD_INPUT_HANDLE = -10;

    internal static readonly IntPtr INVALID_HANDLE_VALUE = new IntPtr(-1);

    [StructLayout(LayoutKind.Sequential)]
    internal struct COORD {
      public short X;
      public short Y;

      public COORD(short x, short y) {
        this.X = x;
        this.Y = y;
      }
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct SMALL_RECT {
      public short Left;
      public short Top;
      public short Right;
      public short Bottom;

      public SMALL_RECT(short Left, short Top, short Right, short Bottom) {
          this.Left = Left;
          this.Top = Top;
          this.Right = Right;
          this.Bottom = Bottom;
      }
    }

    [DllImport(Dll)]
    internal static extern IntPtr GetStdHandle(int nStdHandle);

    [DllImport(Dll, SetLastError = true)]
    internal static extern bool WriteConsole(
      IntPtr hConsoleOutput,
      string lpBuffer,
      uint nNumberOfCharsToWrite,
      out uint lpNumberOfCharsWritten,
      IntPtr lpReserved
    );

    [DllImport(Dll)]
    internal static extern bool WriteConsoleOutputCharacter(
      IntPtr hConsoleOutput,
      string lpCharacter,
      uint nLength,
      COORD dwWriteCoord,
      out uint lpNumberOfCharsWritten
    );

    [DllImport(Dll, SetLastError = true)]
    public static extern bool WriteConsoleOutputAttribute(
        IntPtr hConsoleOutput,
        ushort[] lpAttribute,
        uint nLength,
        COORD dwWriteCoord,
        out uint lpNumberOfAttrsWritten
    );

    [StructLayout(LayoutKind.Sequential)]
    internal struct CONSOLE_SCREEN_BUFFER_INFO {
      public COORD dwSize;
      public COORD dwCursorPosition;
      public short wAttributes;
      public SMALL_RECT srWindow;
      public COORD dwMaximumWindowSize;
    }

    [DllImport(Dll)]
    private static extern bool GetConsoleScreenBufferInfo(IntPtr hConsoleOutput, out CONSOLE_SCREEN_BUFFER_INFO lpConsoleScreenBufferInfo);

    internal static CONSOLE_SCREEN_BUFFER_INFO GetScreenBufferInfo(IntPtr handle) {
      GetConsoleScreenBufferInfo(handle, out var info);
      return info;
    }

  [StructLayout(LayoutKind.Explicit)]
    public struct CHAR_INFO_UNION
    {
        [FieldOffset(0)]
        public char UnicodeChar;
        [FieldOffset(0)]
        public byte AsciiChar;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct CHAR_INFO{
      public CHAR_INFO_UNION Char;
      public ushort Attributes;
    }

    [DllImport(Dll)]
    internal static extern bool ReadConsoleOutput(
      IntPtr hConsoleOutput,
      [Out] CHAR_INFO[] lpBuffer,
      COORD dwBufferSize,
      COORD dwBufferCoord,
      ref SMALL_RECT lpReadRegion
    );

    [DllImport(Dll, SetLastError = true)]
    internal static extern bool WriteConsoleOutput(
      IntPtr hConsoleOutput,
      [In] CHAR_INFO[] lpBuffer,
      COORD dwBufferSize,
      COORD dwBufferCoord,
      ref SMALL_RECT lpWriteRegion
    );

  }

}