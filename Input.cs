using System;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;

using Win32;
using static Win32.Kernel32;

namespace Console {

   [ComVisible(true)]
   [ClassInterface(ClassInterfaceType.AutoDispatch)]
   [Guid("E4B35606-B68E-4B54-A438-E2DD1B139022")]
   [ProgId("Console.Input")]

    public class Input {

        private IntPtr inputHandle;

        public Input() {
          inputHandle = GetStdHandle(STD_INPUT_HANDLE);
        }

        public string GetInputEvent() {
          var inputEventType = GetNextInputEventType();

          object eventDetails;
          switch (inputEventType) {
            case InputEventType.WindowEvent:
              eventDetails = GetWindowEventDetails();
              break;
            case InputEventType.KeyEvent:
              eventDetails = GetKeyEventDetails();
              break;
            case InputEventType.MouseEvent:
              eventDetails = GetMouseEventDetails();
              break;
            default:
              eventDetails = new { type = "None" };
              break;
          }

          var serializer = new JavaScriptSerializer();
          return serializer.Serialize(eventDetails);
        }

        private bool GetConsoleMode(uint bit) {
          Kernel32.GetConsoleMode(inputHandle, out uint mode);
          return (mode & bit) != 0;
        }

        private void SetConsoleMode(uint bit, bool value) {
          Kernel32.GetConsoleMode(inputHandle, out uint mode);
          mode = value ? (mode | bit) : (mode & ~bit);
          Kernel32.SetConsoleMode(inputHandle, mode);
        }

        public bool EchoInputEnabled {
          get { return GetConsoleMode(ENABLE_ECHO_INPUT); }
          set { SetConsoleMode(ENABLE_ECHO_INPUT, value); }
        }

        public bool QuickEditModeEnabled {
          get { return GetConsoleMode(ENABLE_QUICK_EDIT_MODE); }
          set { SetConsoleMode(ENABLE_QUICK_EDIT_MODE, value); }
        }

        public bool ProcessedInputEnabled {
          get { return GetConsoleMode(ENABLE_PROCESSED_INPUT); }
          set { SetConsoleMode(ENABLE_PROCESSED_INPUT, value); }
        }

        public bool InsertModeEnabled {
          get { return GetConsoleMode(ENABLE_INSERT_MODE); }
          set { SetConsoleMode(ENABLE_INSERT_MODE, value); }
        }

        public bool LineInputEnabled {
          get { return GetConsoleMode(ENABLE_LINE_INPUT); }
          set { SetConsoleMode(ENABLE_LINE_INPUT, value); }
        }

        public bool MouseInputEnabled {
          get { return GetConsoleMode(ENABLE_MOUSE_INPUT); }
          set { SetConsoleMode(ENABLE_MOUSE_INPUT, value); }
        }

        public bool WindowInputEnabled {
          get { return GetConsoleMode(ENABLE_WINDOW_INPUT); }
          set { SetConsoleMode(ENABLE_WINDOW_INPUT, value); }
        }

        private InputEventType GetNextInputEventType() {
          uint numberOfEvents = 0;
          if (GetNumberOfConsoleInputEvents(inputHandle, ref numberOfEvents) && numberOfEvents > 0) {

            uint eventsRead = 0;
            INPUT_RECORD inputRecord = new INPUT_RECORD();
            if (PeekConsoleInput(inputHandle, ref inputRecord, 1, ref eventsRead) && eventsRead > 0) {

              switch (inputRecord.EventType) {
                case FOCUS_EVENT:
                case WINDOW_BUFFER_SIZE_EVENT:
                  return InputEventType.WindowEvent;
                case KEY_EVENT:
                  return InputEventType.KeyEvent;
                case MOUSE_EVENT:
                  return InputEventType.MouseEvent;

                default:
                  return InputEventType.None;
              }

            }
          }
          return InputEventType.None;
        }

        private WindowEventType GetNextWindowEventType()
        {
          uint eventsRead = 0;
          INPUT_RECORD inputRecord = new INPUT_RECORD();
          if (PeekConsoleInput(inputHandle, ref inputRecord, 1, ref eventsRead) && eventsRead > 0) {
            switch (inputRecord.EventType) {
              case FOCUS_EVENT:
                return WindowEventType.Focus;
              case WINDOW_BUFFER_SIZE_EVENT:
                return WindowEventType.Resized;

              default:
                return WindowEventType.None;
            }
          }
          return WindowEventType.None;
        }

        private KeyEventType GetNextKeyEventType() {
          uint eventsRead = 0;
          INPUT_RECORD inputRecord = new INPUT_RECORD();
          if (PeekConsoleInput(inputHandle, ref inputRecord, 1, ref eventsRead) && eventsRead > 0) {
            return inputRecord.KeyEvent.bKeyDown ? KeyEventType.Pressed : KeyEventType.Released;
          }
          return KeyEventType.None;
        }

        private MouseEventType GetNextMouseEventType() {
          uint eventsRead = 0;
          INPUT_RECORD inputRecord = new INPUT_RECORD();
          if (PeekConsoleInput(inputHandle, ref inputRecord, 1, ref eventsRead) && eventsRead > 0) {

            switch (inputRecord.MouseEvent.dwEventFlags) {
              case DOUBLE_CLICK:
                return MouseEventType.DoubleClick;
              case MOUSE_HWHEELED:
                return MouseEventType.HorizontalWheel;
              case MOUSE_MOVED:
                return MouseEventType.MouseMoved;
              case MOUSE_WHEELED:
                return MouseEventType.VerticalWheel;

              default:
                return inputRecord.MouseEvent.dwButtonState != 0 ? MouseEventType.SingleClick : MouseEventType.ButtonReleased;
            }
          }
          return MouseEventType.None;
        }

        private KeyEventBase GetNextKeyEvent(bool isPressed) {
          KeyEventBase keyEvent;
          if (isPressed) {
            keyEvent = new PressedKeyEvent();
          } else {
            keyEvent = new ReleasedKeyEvent();
          };

          uint eventsRead = 0;
          INPUT_RECORD inputRecord = new INPUT_RECORD();
          ReadConsoleInput(inputHandle, ref inputRecord, 1, ref eventsRead);

          var keyEventDetails = inputRecord.KeyEvent;
          keyEvent.ScanCode = keyEventDetails.wVirtualScanCode;
          keyEvent.KeyCode = keyEventDetails.wVirtualKeyCode;
          keyEvent.UnicodeChar = keyEventDetails.uChar.UnicodeChar;
          keyEvent.AsciiCharCode = keyEventDetails.uChar.AsciiChar;
          keyEvent.AsciiChar = (char)keyEventDetails.uChar.AsciiChar;
          keyEvent.ControlKeyState = keyEventDetails.dwControlKeyState;
          keyEvent.ControlKeys = GetControlKeys(keyEventDetails.dwControlKeyState);

          if (isPressed && keyEvent is PressedKeyEvent pressedKeyEvent) {
            pressedKeyEvent.Repetitions = keyEventDetails.wRepeatCount;
          }
          return keyEvent;
        }

        private object GetKeyEventDetails() {
          var keyEventType = GetNextKeyEventType();
          switch (keyEventType) {
            case KeyEventType.Pressed:
              return CreateKeyEventDetails("KeyPressed", (PressedKeyEvent)GetNextKeyEvent(true));
            case KeyEventType.Released:
              return CreateKeyEventDetails("KeyReleased", (ReleasedKeyEvent)GetNextKeyEvent(false));

            default:
              return new { type = "None" };
          }
        }

        private MouseEventBase GetNextMouseEvent(MouseEventType eventType) {
          MouseEventBase mouseEvent;
          switch (eventType) {
            case MouseEventType.DoubleClick:
            case MouseEventType.SingleClick:
              mouseEvent = new MouseButtonClickedEvent();
              break;
            case MouseEventType.HorizontalWheel:
              mouseEvent = new MouseHorizontalWheelEvent();
              break;
            case MouseEventType.VerticalWheel:
              mouseEvent = new MouseVerticalWheelEvent();
              break;
            default:
              mouseEvent = new MouseEvent();
              break;
          }

          uint eventsRead = 0;
          INPUT_RECORD inputRecord = new INPUT_RECORD();
          ReadConsoleInput(inputHandle, ref inputRecord, 1, ref eventsRead);

          mouseEvent.CursorLocation = new MouseCursorLocation { Row = inputRecord.MouseEvent.dwMousePosition.Y, Column = inputRecord.MouseEvent.dwMousePosition.X };
          mouseEvent.ControlKeys = GetControlKeys(inputRecord.MouseEvent.dwControlKeyState);

          if (mouseEvent is MouseButtonClickedEvent clickedEvent) {
              clickedEvent.Buttons = GetMouseButtons(inputRecord.MouseEvent.dwButtonState);
          } else if (mouseEvent is MouseWheelEvent wheelEvent) {
              wheelEvent.WheelDirection = (short)HiWord((int)inputRecord.MouseEvent.dwButtonState) > 0 ?
                (mouseEvent is MouseHorizontalWheelEvent ? "right" : "forward") :
                (mouseEvent is MouseHorizontalWheelEvent ? "left" : "backwards");
          }

          return mouseEvent;
        }

        private object GetMouseEventDetails() {
            var mouseEventType = GetNextMouseEventType();
            var mouseEvent = GetNextMouseEvent(mouseEventType);

            var cursorLocation = new {
              row = mouseEvent.CursorLocation.Row,
              column = mouseEvent.CursorLocation.Column
            };

            var controlKeys = GetControlKeysDetails(mouseEvent.ControlKeys);

            switch (mouseEventType) {
              case MouseEventType.DoubleClick:
                return new {
                  type = "MouseDoubleClick",
                  cursorLocation,
                  buttons = GetMouseButtonsDetails(((MouseButtonClickedEvent)mouseEvent).Buttons),
                  controlKeys
                };
              case MouseEventType.SingleClick:
                return new {
                  type = "MouseClick",
                  cursorLocation,
                  buttons = GetMouseButtonsDetails(((MouseButtonClickedEvent)mouseEvent).Buttons),
                  controlKeys
                };
              case MouseEventType.MouseMoved:
                return new {
                  type = "MouseMoved",
                  cursorLocation,
                  controlKeys
                };
              case MouseEventType.HorizontalWheel:
                return new {
                  type = "HorizontalWheel",
                  cursorLocation,
                  direction = ((MouseHorizontalWheelEvent)mouseEvent).WheelDirection,
                  controlKeys
                };
              case MouseEventType.VerticalWheel:
                return new {
                  type = "VerticalWheel",
                  cursorLocation,
                  direction = ((MouseVerticalWheelEvent)mouseEvent).WheelDirection,
                  controlKeys
                };
              case MouseEventType.ButtonReleased:
                return new {
                  type = "MouseButtonReleased",
                  cursorLocation,
                  controlKeys
                };
              default:
                return new { type = "None" };
            }
        }

        private object GetWindowEventDetails() {
          var windowEventType = GetNextWindowEventType();
          switch (windowEventType) {
            case WindowEventType.Focus:
              return new {
                type = "WindowFocus",
                focused = GetNextWindowFocusEvent().Focused
              };
            case WindowEventType.Resized:
              return new {
                type = "WindowResized",
                rows = GetNextWindowResizedEvent().Rows,
                columns = GetNextWindowResizedEvent().Columns
              };
            default:
              return new { type = "None" };
          }
        }

        private bool IsModifierKeyPressed(int keyCode, uint controlKeyState) {
            switch (keyCode) {
                case 0x10: // Shift
                    return (controlKeyState & SHIFT_PRESSED) != 0;
                case 0x11: // Ctrl
                    return ((controlKeyState & LEFT_CTRL_PRESSED) != 0 || 
                            (controlKeyState & RIGHT_CTRL_PRESSED) != 0);
                case 0x12: // Alt
                    return ((controlKeyState & LEFT_ALT_PRESSED) != 0 || 
                            (controlKeyState & RIGHT_ALT_PRESSED) != 0);
                default:
                    return false;
            }
        }

        private object CreateKeyEventDetails(string eventType, KeyEventBase keyEvent) {
          return new {
            type = eventType,
            scanCode = keyEvent.ScanCode,
            keyCode = keyEvent.KeyCode,
            unicodeChar = keyEvent.UnicodeChar,
            unicodeCharCode = (int)keyEvent.UnicodeChar,
            asciiCharCode = (int)keyEvent.AsciiCharCode,
            asciiChar = keyEvent.AsciiCharCode > 0 && keyEvent.AsciiCharCode < 127 ? keyEvent.AsciiChar.ToString() : "",
            repetitions = keyEvent is PressedKeyEvent pressedKeyEvent ? pressedKeyEvent.Repetitions : (int?)null,
            keyType = GetKeyType(keyEvent.KeyCode),
            shiftPressed = IsModifierKeyPressed(0x10, keyEvent.ControlKeyState),
            ctrlPressed = IsModifierKeyPressed(0x11, keyEvent.ControlKeyState),
            altPressed = IsModifierKeyPressed(0x12, keyEvent.ControlKeyState),
            controlKeys = GetControlKeysDetails(keyEvent.ControlKeys)
          };
        }

        private ControlKeys GetControlKeys(uint controlKeyState) {
          return new ControlKeys {
            CapsLockOn = (controlKeyState & CAPSLOCK_ON) != 0,
            EnhancedKey = (controlKeyState & ENHANCED_KEY) != 0,
            LeftAltPressed = (controlKeyState & LEFT_ALT_PRESSED) != 0,
            LeftCtrlPressed = (controlKeyState & LEFT_CTRL_PRESSED) != 0,
            NumLockOn = (controlKeyState & NUMLOCK_ON) != 0,
            RightAltPressed = (controlKeyState & RIGHT_ALT_PRESSED) != 0,
            RightCtrlPressed = (controlKeyState & RIGHT_CTRL_PRESSED) != 0,
            ScrollLockOn = (controlKeyState & SCROLLLOCK_ON) != 0,
            ShiftPressed = (controlKeyState & SHIFT_PRESSED) != 0
          };
        }

        private MouseButtons GetMouseButtons(uint buttonState) {
          return new MouseButtons {
            Left1Pressed = (buttonState & FROM_LEFT_1ST_BUTTON_PRESSED) != 0,
            Left2Pressed = (buttonState & FROM_LEFT_2ND_BUTTON_PRESSED) != 0,
            Left3Pressed = (buttonState & FROM_LEFT_3RD_BUTTON_PRESSED) != 0,
            Left4Pressed = (buttonState & FROM_LEFT_4TH_BUTTON_PRESSED) != 0,
            RightMostPressed = (buttonState & RIGHTMOST_BUTTON_PRESSED) != 0
          };
        }

        private string GetKeyType(int keyCode) {
            // Check for modifier keys first
            if (keyCode == 0x10 || keyCode == 0x11 || keyCode == 0x12) { // Shift, Ctrl, Alt
                return "modifier";
            }

            // Numpad keys
            if ((keyCode >= 0x60 && keyCode <= 0x69) ||  // Numpad 0-9
                (keyCode >= 0x6A && keyCode <= 0x6F)) {  // Numpad operators
                return "numpad";
            }

            // System keys
            if (keyCode == 0x5B || keyCode == 0x5C ||    // Left/Right Windows key
                keyCode == 0x5D ||                        // Menu key
                keyCode == 0x2C ||                        // Print Screen
                keyCode == 0x13) {                        // Pause/Break
                return "system";
            }

            // Lock keys
            if (keyCode == 0x14 ||  // Caps Lock
                keyCode == 0x90 ||  // Num Lock
                keyCode == 0x91) {  // Scroll Lock
                return "lock";
            }

            // Browser keys
            if (keyCode >= 0xA6 && keyCode <= 0xAC) {  // Browser navigation keys
                return "browser";
            }

            // Application keys
            if (keyCode >= 0xB6 && keyCode <= 0xB7 ||  // Launch App1/App2
                keyCode == 0xB4) {                      // Launch Mail
                return "application";
            }

            if (keyCode >= 0x41 && keyCode <= 0x5A)
                return "alphabetic";
            if (keyCode >= 0x30 && keyCode <= 0x39)
                return "numeric";
            if (keyCode >= 0x70 && keyCode <= 0x87)
                return "function";
            if (keyCode >= 0x21 && keyCode <= 0x28)
                return "navigation";
            if (keyCode == 0x20)
                return "alphabetic";
            if (keyCode == 0x09)
                return "navigation";
            if (keyCode == 0x08 || keyCode == 0x2E || keyCode == 0x2D)
                return "edition";
            if (keyCode >= 0xBB && keyCode <= 0xBE)
                return "punctuation";
            if (keyCode >= 0xAD && keyCode <= 0xB3)
                return "media";

            return "none";
        }

        private object GetControlKeysDetails(ControlKeys controlKeys) {
          return new {
            capsLockOn = controlKeys.CapsLockOn,
            enhancedKey = controlKeys.EnhancedKey,
            leftAltPressed = controlKeys.LeftAltPressed,
            leftCtrlPressed = controlKeys.LeftCtrlPressed,
            numLockOn = controlKeys.NumLockOn,
            rightAltPressed = controlKeys.RightAltPressed,
            rightCtrlPressed = controlKeys.RightCtrlPressed,
            scrollLockOn = controlKeys.ScrollLockOn,
            shiftPressed = controlKeys.ShiftPressed
          };
        }

        private object GetMouseButtonsDetails(MouseButtons buttons) {
          return new {
            left1Pressed = buttons.Left1Pressed,
            left2Pressed = buttons.Left2Pressed,
            left3Pressed = buttons.Left3Pressed,
            left4Pressed = buttons.Left4Pressed,
            rightMostPressed = buttons.RightMostPressed
          };
        }

        private WindowResizedEvent GetNextWindowResizedEvent() {
          uint eventsRead = 0;
          INPUT_RECORD inputRecord = new INPUT_RECORD();
          ReadConsoleInput(inputHandle, ref inputRecord, 1, ref eventsRead);
          return new WindowResizedEvent {
            Rows = inputRecord.WindowBufferSizeEvent.dwSize.Y,
            Columns = inputRecord.WindowBufferSizeEvent.dwSize.X
          };
        }

        private WindowFocusEvent GetNextWindowFocusEvent() {
          uint eventsRead = 0;
          INPUT_RECORD inputRecord = new INPUT_RECORD();
          ReadConsoleInput(inputHandle, ref inputRecord, 1, ref eventsRead);
          return new WindowFocusEvent {
              Focused = inputRecord.FocusEvent.bSetFocus
          };
        }

        public int CodePage {
          get => (int)GetConsoleCP();
          set => SetConsoleCP((uint)value);
        }
    }

    internal enum InputEventType {
      None,
      WindowEvent,
      KeyEvent,
      MouseEvent
    }

    internal enum WindowEventType {
      None,
      Focus,
      Resized
    }

    internal enum KeyEventType {
      None,
      Pressed,
      Released
    }

    internal enum MouseEventType {
      None,
      MouseMoved,
      SingleClick,
      DoubleClick,
      HorizontalWheel,
      VerticalWheel,
      ButtonReleased
    }

    internal abstract class InputEventBase {
      public ControlKeys ControlKeys { get; set; }
    }

    internal abstract class KeyEventBase : InputEventBase {
      public int ScanCode { get; set; }
      public int KeyCode { get; set; }
      public char UnicodeChar { get; set; }
      public byte AsciiCharCode { get; set; }
      public char AsciiChar { get; set; }
      public uint ControlKeyState { get; set; }
    }

    internal class PressedKeyEvent : KeyEventBase {
      public int Repetitions { get; set; }
    }

    internal class ReleasedKeyEvent : KeyEventBase { }

    internal abstract class MouseEventBase : InputEventBase {
      public MouseCursorLocation CursorLocation { get; set; }
    }

    internal class MouseButtonClickedEvent : MouseEventBase {
      public MouseButtons Buttons { get; set; }
    }

    internal class MouseEvent : MouseEventBase { }

    internal class MouseWheelEvent : MouseEventBase {
        public string WheelDirection { get; set; }
    }

    internal class MouseHorizontalWheelEvent : MouseWheelEvent { }

    internal class MouseVerticalWheelEvent : MouseWheelEvent { }

    internal class WindowResizedEvent {
      public int Rows { get; set; }
      public int Columns { get; set; }
    }

    internal class WindowFocusEvent {
      public bool Focused { get; set; }
    }

    internal class ControlKeys {
      public bool CapsLockOn { get; set; }
      public bool EnhancedKey { get; set; }
      public bool LeftAltPressed { get; set; }
      public bool LeftCtrlPressed { get; set; }
      public bool NumLockOn { get; set; }
      public bool RightAltPressed { get; set; }
      public bool RightCtrlPressed { get; set; }
      public bool ScrollLockOn { get; set; }
      public bool ShiftPressed { get; set; }
    }

    internal class MouseCursorLocation {
      public int Row { get; set; }
      public int Column { get; set; }
    }

    internal class MouseButtons {
      public bool Left1Pressed { get; set; }
      public bool Left2Pressed { get; set; }
      public bool Left3Pressed { get; set; }
      public bool Left4Pressed { get; set; }
      public bool RightMostPressed { get; set; }
    }
}
