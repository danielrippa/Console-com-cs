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

    private IntPtr handle = GetStdHandle(STD_OUTPUT_HANDLE);

    public bool AssignHandle() {

      IntPtr newHandle = CreateConsoleScreenBuffer(
        GENERIC_READ | GENERIC_WRITE,
        0,
        IntPtr.Zero,
        CONSOLE_TEXTMODE_BUFFER,
        IntPtr.Zero
      );

      if (newHandle == INVALID_HANDLE_VALUE) {
        return false;
      }

      var info = GetBufferInfoEx(GetStdHandle(STD_OUTPUT_HANDLE));

      SetBufferInfoEx(newHandle, info);

      if (!SetConsoleActiveScreenBuffer(newHandle)) {
        CloseHandle(newHandle);
        return false;
      }

      handle = newHandle;
      return true;
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

    public bool SetCharsAt(short row, short column, string text) {
      uint charsWritten;
      COORD writeCoord = new COORD(column, row);
      return WriteConsoleOutputCharacter(handle, text, (uint)text.Length, writeCoord, out charsWritten);
    }

    public bool SetAttrsAt(short row, short column, string attrsCsv) {

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

      COORD dwWriteCoord = new COORD(column, row);

      uint lpNumberOfAttrsWritten;

      return WriteConsoleOutputAttribute(
        handle,
        parsedAttrs,
        attrsToWrite,
        dwWriteCoord,
        out lpNumberOfAttrsWritten
      );

    }

    public short Width {
      get => GetBufferInfo(handle).dwSize.Y;
    }

    public short Height {
      get => GetBufferInfo(handle).dwSize.X;
    }

    private CHAR_INFO[] pasteAreaBuffer;
    private short pasteAreaWidth;
    private short pasteAreaHeight;

    public bool CopyArea(short row, short col, short height, short width, [Optional, DefaultParameterValue(0L)] long sourceScreenBufferHandle) {

      if (width <= 0 || height <= 0) {
        return false;
      }

      IntPtr screenBufferHandle = handle;
      if (sourceScreenBufferHandle != 0L && (IntPtr)sourceScreenBufferHandle != INVALID_HANDLE_VALUE) {
        screenBufferHandle = (IntPtr)sourceScreenBufferHandle;
      }

      pasteAreaWidth = width;
      pasteAreaHeight = height;

      if (pasteAreaBuffer == null || pasteAreaBuffer.Length != (width * height)) {
        pasteAreaBuffer = new CHAR_INFO[width * height];
      }

      COORD bufferSize = new COORD(width, height);
      COORD bufferCoord = new COORD(0, 0);

      SMALL_RECT readRegion = new SMALL_RECT(col, row, (short)(col + width - 1), (short)(row + height - 1));

      return ReadConsoleOutput(
        screenBufferHandle,
        pasteAreaBuffer,
        bufferSize,
        bufferCoord,
        ref readRegion
      );

    }

    public bool PasteAreaAt(short row, short col, [Optional, DefaultParameterValue(0L)] long targetScreenBufferHandle) {

      if (pasteAreaBuffer == null || pasteAreaBuffer.Length == 0 || pasteAreaWidth == 0 || pasteAreaHeight == 0) {
        return false;
      }

      IntPtr screenBufferHandle = handle;
      if (targetScreenBufferHandle != 0L && (IntPtr)targetScreenBufferHandle != INVALID_HANDLE_VALUE) {
        screenBufferHandle = (IntPtr)targetScreenBufferHandle;
      }

      COORD bufferSize = new COORD(pasteAreaWidth, pasteAreaHeight);
      COORD bufferCoord = new COORD(0, 0);

      SMALL_RECT writeRegion = new SMALL_RECT(
        col,
        row,
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

    public bool SetPasteAreaContent(short width, short height, string chars, string attrsCsv, char fillChar = ' ', ushort? defaultAttrs = null) {

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

      ushort actualDefaultAttrs;
      if (defaultAttrs.HasValue) {
        actualDefaultAttrs = defaultAttrs.Value;
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
          pasteAreaBuffer[i].Char.UnicodeChar = fillChar;
        }

        if (i < parsedAttrs.Length) {
          pasteAreaBuffer[i].Attributes = parsedAttrs[i];
        } else {
          pasteAreaBuffer[i].Attributes = actualDefaultAttrs;
        }

      }

      return true;
    }

    public bool SetPasteAreaChars(string charData, [Optional, DefaultParameterValue((short)0)] short row, [Optional, DefaultParameterValue((short)0)] short col) {

      if (pasteAreaBuffer == null || pasteAreaBuffer.Length == 0 || pasteAreaWidth == 0 || pasteAreaHeight == 0) {
        return false;
      }

      if (charData == null) {
        return false;
      }

      if (row < 0 || row >= pasteAreaHeight || col < 0 || col >= pasteAreaWidth) {
        return false;
      }

      int startIndex = (row * pasteAreaWidth) + col;
      int charsFitOnRow = pasteAreaWidth - col;

      int charsToWrite = Math.Min(charData.Length, charsFitOnRow);
      for (int i = 0; i < charsToWrite; i++) {
        pasteAreaBuffer[startIndex + i].Char.UnicodeChar = charData[i];
      }

      return true;
    }

  }

}