using System;
using System.Runtime.InteropServices;
using System.Text;

namespace Win32 {

  internal static class Kernel32 {

    private const string Dll = "kernel32.dll";

    internal const int STD_OUTPUT_HANDLE = -11;
    internal const int STD_INPUT_HANDLE = -10;

    internal static readonly IntPtr INVALID_HANDLE_VALUE = new IntPtr(-1);

    internal static IntPtr GetStdOutHandle() {
      return GetStdHandle(STD_OUTPUT_HANDLE);
    }

    internal const uint GENERIC_READ = 0x80000000;
    internal const uint GENERIC_WRITE = 0x40000000;
    internal const uint CONSOLE_TEXTMODE_BUFFER = 1;

    [DllImport(Dll, SetLastError = true)]
    internal static extern IntPtr CreateConsoleScreenBuffer(
      uint dwDesiredAccess,
      uint dwShareMode,
      IntPtr lpSecurityAttributes,
      uint dwFlags,
      IntPtr lpScreenBufferData);

    [DllImport(Dll, SetLastError = true)]
    internal static extern bool SetConsoleActiveScreenBuffer(IntPtr hConsoleOutput);

    [DllImport(Dll, SetLastError = true)]
    internal static extern bool CloseHandle(IntPtr hObject);

    [DllImport(Dll)]
    internal static extern IntPtr GetStdHandle(int nStdHandle);

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

    internal static CONSOLE_SCREEN_BUFFER_INFO GetBufferInfo(IntPtr handle) {
      GetConsoleScreenBufferInfo(handle, out var info);
      return info;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct CONSOLE_SCREEN_BUFFER_INFOEX
    {
        public uint cbSize;
        public COORD dwSize;
        public COORD dwCursorPosition;
        public short wAttributes;
        public SMALL_RECT srWindow;
        public COORD dwMaximumWindowSize;
        public short wPopupAttributes;
        public int bFullscreenSupported; // BOOL is 4 bytes
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 16)]
        public uint[] ColorTable;
    }

    [DllImport(Dll, SetLastError = true)]
    internal static extern bool GetConsoleScreenBufferInfoEx(IntPtr hConsoleOutput, ref CONSOLE_SCREEN_BUFFER_INFOEX lpConsoleScreenBufferInfoEx);

    [DllImport(Dll, SetLastError = true)]
    internal static extern bool SetConsoleScreenBufferInfoEx(IntPtr hConsoleOutput, ref CONSOLE_SCREEN_BUFFER_INFOEX lpConsoleScreenBufferInfoEx);

    internal static CONSOLE_SCREEN_BUFFER_INFOEX GetBufferInfoEx(IntPtr handle) {

      var infoEx = new CONSOLE_SCREEN_BUFFER_INFOEX();
      infoEx.cbSize = (uint)Marshal.SizeOf(typeof(CONSOLE_SCREEN_BUFFER_INFOEX));
      infoEx.ColorTable = new uint[16];

      if (!GetConsoleScreenBufferInfoEx(handle, ref infoEx))
      {
          int error = Marshal.GetLastWin32Error();
          throw new System.ComponentModel.Win32Exception(error, $"Failed to get console screen buffer info extended. Error: {error}");
      }

      // --- Always correct buffer size to fit window ---
      short winWidth = (short)(infoEx.srWindow.Right - infoEx.srWindow.Left + 1);
      short winHeight = (short)(infoEx.srWindow.Bottom - infoEx.srWindow.Top + 1);
      if (infoEx.dwSize.X < winWidth) infoEx.dwSize.X = winWidth;
      if (infoEx.dwSize.Y < winHeight) infoEx.dwSize.Y = winHeight;
      // ------------------------------------------------

      return infoEx;

    }

    internal static bool SetBufferInfoEx(IntPtr handle, CONSOLE_SCREEN_BUFFER_INFOEX infoEx) {
      infoEx.srWindow.Right += 1;
      infoEx.srWindow.Bottom += 1;
      infoEx.cbSize = (uint)Marshal.SizeOf(typeof(CONSOLE_SCREEN_BUFFER_INFOEX));
      return SetConsoleScreenBufferInfoEx(handle, ref infoEx);
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct CHAR_INFO_UNION {
      [FieldOffset(0)]
      public char UnicodeChar;
      [FieldOffset(0)]
      public byte AsciiChar;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct CHAR_INFO {
      public CHAR_INFO_UNION Char;
      public ushort Attributes;
    }

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

    [DllImport(Dll, SetLastError = true)]

    internal static extern bool FillConsoleOutputCharacter(
      IntPtr hConsoleOutput,
      char cCharacter,
      uint nLength,
      COORD dwWriteCoord,
      out uint lpNumberOfCharsWritten
    );

    [DllImport(Dll, SetLastError = true)]
    internal static extern bool FillConsoleOutputAttribute(
      IntPtr hConsoleOutput,
      ushort wAttribute,
      uint nLength,
      COORD dwWriteCoord,
      out uint lpNumberOfAttrsWritten
    );

    [DllImport(Dll, CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern uint GetConsoleTitle(StringBuilder lpConsoleTitle, uint nSize);

    [DllImport(Dll, CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern bool SetConsoleTitle(string lpConsoleTitle);

    [DllImport(Dll, SetLastError = true)]
    internal static extern bool SetConsoleScreenBufferSize(IntPtr hConsoleOutput, COORD dwSize);

    [DllImport(Dll, SetLastError = true)]
    internal static extern bool SetConsoleWindowInfo(IntPtr hConsoleOutput, bool bAbsolute, ref SMALL_RECT lpConsoleWindow);

  }

}