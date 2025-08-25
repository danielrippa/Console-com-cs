using System;
using System.Runtime.InteropServices;
using static Win32.Kernel32;

namespace Console {

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.AutoDispatch)]
  [Guid("17982804-7649-47E9-9787-805F29744287")]
  [ProgId("Console.TextAttributes")]
  public class TextAttributes {

    public short Value;

    public TextAttributes() {
      Value = 0;
    }

    public short SetInk(bool red = false, bool green = false, bool blue = false, bool intensity = false) {
      Value = ApplyTextColor(Value, red, green, blue, intensity, 0);
      return Value;
    }

    public short SetPaper(bool red = false, bool green = false, bool blue = false, bool intensity = false) {
      Value = ApplyTextColor(Value, red, green, blue, intensity, 4);
      return Value;
    }

    public short SetBorders(bool top = false, bool left = false, bool bottom = false, bool right = false) {
      Value = ApplyTextBorders(Value, top, left, bottom, right);
      return Value;
    }

    public short SetInverted(bool inverted = true) {
      Value = SetBit(Value, 14, inverted);
      return Value;
    }

    public short SetAll(bool inkRed = false, bool inkGreen = false, bool inkBlue = false, bool inkIntensity = false,
                                bool paperRed = false, bool paperGreen = false, bool paperBlue = false, bool paperIntensity = false,
                                bool borderTop = false, bool borderLeft = false, bool borderBottom = false, bool borderRight = false,
                                bool inverted = false) {
      Value = ApplyTextColor(Value, inkRed, inkGreen, inkBlue, inkIntensity, 0);
      Value = ApplyTextColor(Value, paperRed, paperGreen, paperBlue, paperIntensity, 4);
      Value = ApplyTextBorders(Value, borderTop, borderLeft, borderBottom, borderRight);
      Value = SetBit(Value, 14, inverted);
      return Value;
    }

    public short EnableInkBits(bool red = false, bool green = false, bool blue = false, bool intensity = false) {
      if (red) Value = SetBit(Value, 2 + 0, true);
      if (green) Value = SetBit(Value, 1 + 0, true);
      if (blue) Value = SetBit(Value, 0 + 0, true);
      if (intensity) Value = SetBit(Value, 3 + 0, true);
      return Value;
    }

    public short DisableInkBits(bool red = false, bool green = false, bool blue = false, bool intensity = false) {
      if (red) Value = SetBit(Value, 2 + 0, false);
      if (green) Value = SetBit(Value, 1 + 0, false);
      if (blue) Value = SetBit(Value, 0 + 0, false);
      if (intensity) Value = SetBit(Value, 3 + 0, false);
      return Value;
    }

    public short EnablePaperBits(bool red = false, bool green = false, bool blue = false, bool intensity = false) {
      if (red) Value = SetBit(Value, 2 + 4, true);
      if (green) Value = SetBit(Value, 1 + 4, true);
      if (blue) Value = SetBit(Value, 0 + 4, true);
      if (intensity) Value = SetBit(Value, 3 + 4, true);
      return Value;
    }

    public short DisablePaperBits(bool red = false, bool green = false, bool blue = false, bool intensity = false) {
      if (red) Value = SetBit(Value, 2 + 4, false);
      if (green) Value = SetBit(Value, 1 + 4, false);
      if (blue) Value = SetBit(Value, 0 + 4, false);
      if (intensity) Value = SetBit(Value, 3 + 4, false);
      return Value;
    }

    public short EnableBorderBits(bool top = false, bool left = false, bool bottom = false, bool right = false) {
      if (top) Value = SetBit(Value, 10, true);
      if (left) Value = SetBit(Value, 11, true);
      if (bottom) Value = SetBit(Value, 12, true);
      if (right) Value = SetBit(Value, 15, true);
      return Value;
    }

    public short DisableBorderBits(bool top = false, bool left = false, bool bottom = false, bool right = false) {
      if (top) Value = SetBit(Value, 10, false);
      if (left) Value = SetBit(Value, 11, false);
      if (bottom) Value = SetBit(Value, 12, false);
      if (right) Value = SetBit(Value, 15, false);
      return Value;
    }

    public short EnableInvertedBit(bool inverted = true) {
      if (inverted) Value = SetBit(Value, 14, true);
      return Value;
    }

    public short DisableInvertedBit(bool inverted = true) {
      if (inverted) Value = SetBit(Value, 14, false);
      return Value;
    }

    private static short SetBit(short value, int bit, bool enabled) {
      short mask = (short)(1 << bit);
      return enabled ? (short)(value | mask) : (short)(value & ~mask);
    }

    private short ApplyTextColor(short value, bool red, bool green, bool blue, bool intensity, int offset) {
      value = SetBit(value, 3 + offset, intensity);
      value = SetBit(value, 2 + offset, red);
      value = SetBit(value, 1 + offset, green);
      value = SetBit(value, 0 + offset, blue);
      return value;
    }

    private short ApplyTextBorders(short value, bool top, bool left, bool bottom, bool right) {
      value = SetBit(value, 10, top);
      value = SetBit(value, 11, left);
      value = SetBit(value, 12, bottom);
      value = SetBit(value, 15, right);
      return value;
    }

    public override string ToString() {
      return Value.ToString();
    }

  }

}