using System;
using System.Runtime.InteropServices;
using System.Linq;

using static Win32.Kernel32;

namespace Console {

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.AutoDispatch)]
  [Guid("F4DD1051-8D4A-4432-91FB-C6B52F06254E")]
  [ProgId("Console.ScreenBuffer")]

  public class ScreenBuffer {

    private IntPtr handle;

    public ScreenBuffer() {
      handle = GetStdHandle(STD_OUTPUT_HANDLE);
    }

    public static IntPtr CreateScreenBufferHandle() {
      return CreateConsoleScreenBuffer(
          GENERIC_READ | GENERIC_WRITE,
          FILE_SHARE_READ | FILE_SHARE_WRITE,
          IntPtr.Zero,
          CONSOLE_TEXTMODE_BUFFER,
          IntPtr.Zero
      );
    }

    public bool Activate() {
      return SetConsoleActiveScreenBuffer(handle);
    }

    public bool IsActive() {
      IntPtr activeHandle = GetStdHandle(STD_OUTPUT_HANDLE);
      return handle == activeHandle;
    }

    public bool Close() {
      if (handle != IntPtr.Zero && handle != INVALID_HANDLE_VALUE) {
        bool result = CloseHandle(handle);
        handle = IntPtr.Zero;
        return result;
      }
      return true;
    }

    public bool Write(string text) {
      if (text == null) return false;
      uint charsWritten;
      return WriteConsole(handle, text, (uint)text.Length, out charsWritten, IntPtr.Zero);
    }

    public bool SetCharsAt(int row, int column, string text) {
      uint charsWritten;
      COORD writeCoord = new COORD((short)column, (short)row);
      return WriteConsoleOutputCharacter(handle, text, (uint)text.Length, writeCoord, out charsWritten);
    }

    public bool SetAttrsAt(int row, int column, string attrsCsv) {

      if (column < 0 || row < 0|| column >= Width || row >= Height) {
        return false;
      }

      int attrsFitOnRow = Width - column;

      if (string.IsNullOrEmpty(attrsCsv)) {
        return false;
      }

      ushort[] parsedAttrs;
      try {

        parsedAttrs = attrsCsv.Split(',').Select(s => ushort.Parse(s.Trim())).ToArray();

      } catch (FormatException) {
        return false;
      } catch (OverflowException) {
        return false;
      }

      uint attrsToWrite = (uint)Math.Min(parsedAttrs.Length, attrsFitOnRow);

      if (attrsToWrite <= 0) {
        return true;
      }

      COORD dwWriteCoord = new COORD((short)column, (short)row);

      uint lpNumberOfAttrsWritten;

      return WriteConsoleOutputAttribute(
        handle,
        parsedAttrs,
        attrsToWrite,
        dwWriteCoord,
        out lpNumberOfAttrsWritten
      );

    }

    public bool SetAttrAt(int row, int column, int attrValue, int length) {
      if (column < 0 || row < 0 || column >= Width || row >= Height) {
        return false;
      }

      int attrsFitOnRow = Width - column;
      uint attrsToWrite = (uint)Math.Min(length, attrsFitOnRow);

      if (attrsToWrite <= 0) {
        return true;
      }

      ushort[] attrs = new ushort[attrsToWrite];
      for (int i = 0; i < attrsToWrite; i++) {
        attrs[i] = (ushort)attrValue;
      }

      COORD dwWriteCoord = new COORD((short)column, (short)row);
      uint lpNumberOfAttrsWritten;

      return WriteConsoleOutputAttribute(
        handle,
        attrs,
        attrsToWrite,
        dwWriteCoord,
        out lpNumberOfAttrsWritten
      );
    }

    public bool WriteTextWithAttrs(int row, int column, string text, int attrValue) {
      if (text == null) return false;

      bool charsResult = SetCharsAt(row, column, text);
      bool attrsResult = SetAttrAt(row, column, attrValue, text.Length);

      return charsResult && attrsResult;
    }

    public int Width {
      get => GetBufferInfo(handle).dwSize.Y;
    }

    public int Height {
      get => GetBufferInfo(handle).dwSize.X;
    }

    public int TextAttributes {
      get => GetBufferInfo(handle).wAttributes;
      set => SetConsoleTextAttribute(handle, (ushort)value);
    }

    public long Handle {
      get => handle.ToInt64();
      set => handle = new IntPtr(value);
    }

    private bool GetMode(uint bit) {
      return ((uint)ModeState & bit) != 0;
    }

    private void SetMode(uint bit, bool value) {
      uint currentMode = (uint)ModeState;
      uint newMode = value ? (currentMode | bit) : (currentMode & ~bit);
      ModeState = (int)newMode;
    }

    public int ModeState {
      get {
        uint currentState;
        GetConsoleMode(handle, out currentState);
        return (int)currentState;
      }
      set {
        SetConsoleMode(handle, (uint)value);
      }
    }

    public bool ProcessedOutputEnabled {
      get { return GetMode(ENABLE_PROCESSED_OUTPUT); }
      set { SetMode(ENABLE_PROCESSED_OUTPUT, value); }
    }

    public bool WrapAtEOLOutputEnabled {
      get { return GetMode(ENABLE_WRAP_AT_EOL_OUTPUT); }
      set { SetMode(ENABLE_WRAP_AT_EOL_OUTPUT, value); }
    }

    public bool VirtualTerminalProcessingEnabled {
      get { return GetMode(ENABLE_VIRTUAL_TERMINAL_PROCESSING); }
      set { SetMode(ENABLE_VIRTUAL_TERMINAL_PROCESSING, value); }
    }

    public bool NewlineAutoReturnEnabled {
      get { return !GetMode(DISABLE_NEWLINE_AUTO_RETURN); }
      set { SetMode(DISABLE_NEWLINE_AUTO_RETURN, !value); }
    }

    private CHAR_INFO[] pasteAreaBuffer;
    private int pasteAreaWidth;
    private int pasteAreaHeight;

    public bool CopyArea(int row, int col, int height, int width, long sourceScreenBufferHandle) {

      if (width <= 0 || height <= 0) {
        return false;
      }

      IntPtr screenBufferHandle = (IntPtr)sourceScreenBufferHandle;

      pasteAreaWidth = width;
      pasteAreaHeight = height;

      if (pasteAreaBuffer == null || pasteAreaBuffer.Length != (width * height)) {
        pasteAreaBuffer = new CHAR_INFO[width * height];
      }

      COORD bufferSize = new COORD((short)width, (short)height);
      COORD bufferCoord = new COORD(0, 0);

      SMALL_RECT readRegion = new SMALL_RECT((short)col, (short)row, (short)(col + width - 1), (short)(row + height - 1));

      return ReadConsoleOutput(
        screenBufferHandle,
        pasteAreaBuffer,
        bufferSize,
        bufferCoord,
        ref readRegion
      );

    }

    public bool PasteAreaAt(int row, int col, long targetScreenBufferHandle) {

      if (pasteAreaBuffer == null || pasteAreaBuffer.Length == 0 || pasteAreaWidth == 0 || pasteAreaHeight == 0) {
        return false;
      }

      IntPtr screenBufferHandle = (IntPtr)targetScreenBufferHandle;

      COORD bufferSize = new COORD((short)pasteAreaWidth, (short)pasteAreaHeight);
      COORD bufferCoord = new COORD(0, 0);

      SMALL_RECT writeRegion = new SMALL_RECT(
        (short)col,
        (short)row,
        (short)(col + pasteAreaWidth - 1),
        (short)(row + pasteAreaHeight - 1)
      );

      return WriteConsoleOutput(
        screenBufferHandle,
        pasteAreaBuffer,
        bufferSize,
        bufferCoord,
        ref writeRegion
      );

    }

    public bool SetPasteAreaContent(int width, int height, string chars, string attrsCsv, string fillChar, int defaultAttrs) {

      if (width <= 0 || height <= 0) {
        return false;
      }

      int expectedTotalLength = width * height;

      if (expectedTotalLength == 0) {

        pasteAreaBuffer = null;
        pasteAreaWidth = 0;
        pasteAreaHeight = 0;

        return true;
      }

      char[] parsedChars = chars.ToCharArray();

      ushort[] parsedAttrs;
      try {

        parsedAttrs = attrsCsv.Split(',').Select(s => ushort.Parse(s.Trim())).ToArray();

      } catch (FormatException) {
        return false;
      } catch (OverflowException) {
        return false;
      }

      char actualFillChar = string.IsNullOrEmpty(fillChar) ? ' ' : fillChar[0];

      ushort actualDefaultAttrs;
      if (defaultAttrs != -1) {
        actualDefaultAttrs = (ushort)defaultAttrs;
      } else {
        actualDefaultAttrs = (ushort)GetBufferInfo(handle).wAttributes;
      }

      pasteAreaBuffer = new CHAR_INFO[expectedTotalLength];
      pasteAreaWidth = width;
      pasteAreaHeight = height;

      for (int i = 0; i < expectedTotalLength; i++) {

        if (i < parsedChars.Length) {
          pasteAreaBuffer[i].Char.UnicodeChar = parsedChars[i];
        } else {
          pasteAreaBuffer[i].Char.UnicodeChar = actualFillChar;
        }

        if (i < parsedAttrs.Length) {
          pasteAreaBuffer[i].Attributes = parsedAttrs[i];
        } else {
          pasteAreaBuffer[i].Attributes = actualDefaultAttrs;
        }

      }

      return true;
    }

    public bool SetPasteAreaChars(int row, int col, string charData) {

      if (pasteAreaBuffer == null || pasteAreaBuffer.Length == 0 || pasteAreaWidth == 0 || pasteAreaHeight == 0) {
        return false;
      }

      if (charData == null) {
        return false;
      }

      int actualRow = row == -1 ? 0 : row;
      int actualCol = col == -1 ? 0 : col;

      if (actualRow < 0 || actualRow >= pasteAreaHeight || actualCol < 0 || actualCol >= pasteAreaWidth) {
        return false;
      }

      int startIndex = (actualRow * pasteAreaWidth) + actualCol;
      int charsFitOnRow = pasteAreaWidth - actualCol;

      int charsToWrite = Math.Min(charData.Length, charsFitOnRow);
      for (int i = 0; i < charsToWrite; i++) {
        pasteAreaBuffer[startIndex + i].Char.UnicodeChar = charData[i];
      }

      return true;
    }

    public bool SetArea(int row, int col, int height, int width, string character, int attributes) {

      if (width < 1 || height < 1 || col < 0 || row < 0) return false;

      var writeRegion = new SMALL_RECT {
        Left = (short)col,
        Top = (short)row,
        Right = (short)(col + width - 1),
        Bottom = (short)(row + height - 1)
      };

      var bufferSize = new COORD { X = (short)width, Y = (short)height };
      var bufferCoord = new COORD { X = 0, Y = 0 };
      var buffer = new CHAR_INFO[width * height];

      char charToUse = string.IsNullOrEmpty(character) ? '|' : character[0];

      if (attributes != -1) {
        var fillCell = new CHAR_INFO {
          Char = new CHAR_INFO_UNION { UnicodeChar = charToUse },
          Attributes = (ushort)attributes
        };

        for (int i = 0; i < buffer.Length; i++) {
          buffer[i] = fillCell;
        }

        return WriteConsoleOutput(handle, buffer, bufferSize, bufferCoord, ref writeRegion);
      } else {

        if (!ReadConsoleOutput(handle, buffer, bufferSize, bufferCoord, ref writeRegion)) {
          return false;
        }

        for (int i = 0; i < buffer.Length; i++) {
          buffer[i].Char.UnicodeChar = charToUse;
        }

        return WriteConsoleOutput(handle, buffer, bufferSize, bufferCoord, ref writeRegion);
      }
    }

    public bool SetAreaAttribute(int row, int col, int height, int width, int attributes) {
      if (width < 1 || height < 1 || col < 0 || row < 0) return false;

      var readRegion = new SMALL_RECT {
        Left = (short)col,
        Top = (short)row,
        Right = (short)(col + width - 1),
        Bottom = (short)(row + height - 1)
      };

      var bufferSize = new COORD { X = (short)width, Y = (short)height };
      var bufferCoord = new COORD { X = 0, Y = 0 };
      var buffer = new CHAR_INFO[width * height];

      if (!ReadConsoleOutput(handle, buffer, bufferSize, bufferCoord, ref readRegion)) {
        return false;
      }

      for (int i = 0; i < buffer.Length; i++) {
        buffer[i].Attributes = (ushort)attributes;
      }

      return WriteConsoleOutput(handle, buffer, bufferSize, bufferCoord, ref readRegion);
    }

  }

}
