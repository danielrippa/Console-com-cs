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

      short winWidth = (short)(infoEx.srWindow.Right - infoEx.srWindow.Left + 1);
      short winHeight = (short)(infoEx.srWindow.Bottom - infoEx.srWindow.Top + 1);
      if (infoEx.dwSize.X < winWidth) infoEx.dwSize.X = winWidth;
      if (infoEx.dwSize.Y < winHeight) infoEx.dwSize.Y = winHeight;

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

    [DllImport(Dll, SetLastError = true)]
    internal static extern IntPtr GetConsoleWindow();

    internal const uint ENABLE_PROCESSED_INPUT = 0x0001;
    internal const uint ENABLE_LINE_INPUT = 0x0002;
    internal const uint ENABLE_ECHO_INPUT = 0x0004;
    internal const uint ENABLE_WINDOW_INPUT = 0x0008;
    internal const uint ENABLE_MOUSE_INPUT = 0x0010;
    internal const uint ENABLE_INSERT_MODE = 0x0020;
    internal const uint ENABLE_QUICK_EDIT_MODE = 0x0040;
    internal const uint ENABLE_EXTENDED_FLAGS = 0x0080;

    // Output Mode Flags
    internal const uint ENABLE_PROCESSED_OUTPUT = 0x0001;
    internal const uint ENABLE_WRAP_AT_EOL_OUTPUT = 0x0002;
    internal const uint ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004;

    internal const uint KEY_EVENT = 0x0001;
    internal const uint MOUSE_EVENT = 0x0002;
    internal const uint WINDOW_BUFFER_SIZE_EVENT = 0x0004;
    internal const uint FOCUS_EVENT = 0x0010;

    internal const uint MOUSE_MOVED = 0x0001;
    internal const uint DOUBLE_CLICK = 0x0002;
    internal const uint MOUSE_WHEELED = 0x0004;
    internal const uint MOUSE_HWHEELED = 0x0008;

    internal const uint CAPSLOCK_ON = 0x0080;
    internal const uint ENHANCED_KEY = 0x0100;
    internal const uint LEFT_ALT_PRESSED = 0x0002;
    internal const uint LEFT_CTRL_PRESSED = 0x0008;
    internal const uint NUMLOCK_ON = 0x0020;
    internal const uint RIGHT_ALT_PRESSED = 0x0001;
    internal const uint RIGHT_CTRL_PRESSED = 0x0004;
    internal const uint SCROLLLOCK_ON = 0x0040;
    internal const uint SHIFT_PRESSED = 0x0010;

    [Flags]
    public enum ControlKeyState : uint
    {
        RIGHT_ALT_PRESSED = 0x0001,
        LEFT_ALT_PRESSED = 0x0002,
        RIGHT_CTRL_PRESSED = 0x0004,
        LEFT_CTRL_PRESSED = 0x0008,
        SHIFT_PRESSED = 0x0010,
        NUMLOCK_ON = 0x0020,
        SCROLLLOCK_ON = 0x0040,
        CAPSLOCK_ON = 0x0080,
        ENHANCED_KEY = 0x0100
    }

    internal const uint FROM_LEFT_1ST_BUTTON_PRESSED = 0x0001;
    internal const uint FROM_LEFT_2ND_BUTTON_PRESSED = 0x0004;
    internal const uint FROM_LEFT_3RD_BUTTON_PRESSED = 0x0008;
    internal const uint FROM_LEFT_4TH_BUTTON_PRESSED = 0x0010;
    internal const uint RIGHTMOST_BUTTON_PRESSED = 0x0002;

    [DllImport(Dll)]
    internal static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);

    [DllImport(Dll)]
    internal static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);

    [DllImport(Dll)]
    internal static extern bool GetNumberOfConsoleInputEvents(IntPtr hConsoleInput, ref uint lpcNumberOfEvents);

    [DllImport(Dll)]
    internal static extern bool PeekConsoleInput(IntPtr hConsoleInput, ref INPUT_RECORD lpBuffer, uint nLength, ref uint lpNumberOfEventsRead);

    [DllImport(Dll)]
    internal static extern bool ReadConsoleInput(IntPtr hConsoleInput, ref INPUT_RECORD lpBuffer, uint nLength, ref uint lpNumberOfEventsRead);

    [DllImport(Dll)]
    internal static extern uint GetConsoleCP();

    [DllImport(Dll)]
    internal static extern bool SetConsoleCP(uint wCodePageID);

    [DllImport(Dll)]
    internal static extern uint GetConsoleOutputCP();

    [DllImport(Dll)]
    internal static extern bool SetConsoleOutputCP(uint wCodePageID);

    // Corrected INPUT_RECORD definition
    [StructLayout(LayoutKind.Explicit)]
    internal struct INPUT_RECORD
    {
        [FieldOffset(0)]
        public ushort EventType; // WORD (2 bytes)

        [FieldOffset(4)] // Union starts at offset 4 due to padding for 4-byte alignment
        public EventUnion Event;

        [StructLayout(LayoutKind.Explicit)]
        internal struct EventUnion
        {
            [FieldOffset(0)]
            public KEY_EVENT_RECORD KeyEvent;
            [FieldOffset(0)]
            public MOUSE_EVENT_RECORD MouseEvent;
            [FieldOffset(0)]
            public WINDOW_BUFFER_SIZE_RECORD WindowBufferSizeEvent;
            // MENU_EVENT_RECORD is not present in the original code, so I won't add it.
            [FieldOffset(0)]
            public FOCUS_EVENT_RECORD FocusEvent;
        }
    }

        internal struct KEY_EVENT_RECORD {
        public int bKeyDown; // BOOL (4 bytes)
        public ushort wRepeatCount; // WORD (2 bytes)
        public ushort wVirtualKeyCode; // WORD (2 bytes)
        public ushort wVirtualScanCode; // WORD (2 bytes)
        public char UnicodeChar; // Direct UnicodeChar
        public uint dwControlKeyState; // DWORD (4 bytes)
    }

    

    [StructLayout(LayoutKind.Sequential)]
    internal struct MOUSE_EVENT_RECORD {
        public COORD dwMousePosition;
        public uint dwButtonState;
        public uint dwControlKeyState;
        public uint dwEventFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct WINDOW_BUFFER_SIZE_RECORD {
        public COORD dwSize;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct FOCUS_EVENT_RECORD {
        public int bSetFocus; // BOOL (4 bytes)
    }

    internal static ushort HiWord(int dword) {
        return (ushort)((dword >> 16) & 0xFFFF);
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct CONSOLE_CURSOR_INFO {
        public uint dwSize;
        public bool bVisible;
    }

    [DllImport(Dll, SetLastError = true)]
    internal static extern bool SetConsoleCursorPosition(IntPtr hConsoleOutput, COORD dwCursorPosition);

    [DllImport(Dll, SetLastError = true)]
    internal static extern bool GetConsoleCursorInfo(IntPtr hConsoleOutput, out CONSOLE_CURSOR_INFO lpConsoleCursorInfo);

    [DllImport(Dll, SetLastError = true)]
    internal static extern bool SetConsoleCursorInfo(IntPtr hConsoleOutput, ref CONSOLE_CURSOR_INFO lpConsoleCursorInfo);

    internal static CONSOLE_CURSOR_INFO GetCursorInfo(IntPtr handle) {
        GetConsoleCursorInfo(handle, out var info);
        return info;
    }

    internal static CONSOLE_SCREEN_BUFFER_INFO GetScreenBufferInfo(IntPtr handle) {
        GetConsoleScreenBufferInfo(handle, out var info);
        return info;
    }

    [DllImport(Dll, SetLastError = true)]
    internal static extern bool SetConsoleTextAttribute(IntPtr hConsoleOutput, ushort wAttributes);

  }

}
