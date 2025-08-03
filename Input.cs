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
            uint numberOfEvents = 0;
            if (!GetNumberOfConsoleInputEvents(inputHandle, ref numberOfEvents) || numberOfEvents == 0) {
                return Serialize(new { type = "None" });
            }

            INPUT_RECORD inputRecord = new INPUT_RECORD();
            uint eventsRead = 0;
            if (!ReadConsoleInput(inputHandle, ref inputRecord, 1, ref eventsRead) || eventsRead == 0) {
                return Serialize(new { type = "None" });
            }

            object eventDetails;
            switch (inputRecord.EventType) {
                case KEY_EVENT:
                    eventDetails = GetKeyEventDetails(inputRecord.KeyEvent);
                    break;
                case MOUSE_EVENT:
                    eventDetails = GetMouseEventDetails(inputRecord.MouseEvent);
                    break;
                case WINDOW_BUFFER_SIZE_EVENT:
                    eventDetails = GetWindowResizedEventDetails(inputRecord.WindowBufferSizeEvent);
                    break;
                case FOCUS_EVENT:
                    eventDetails = GetWindowFocusEventDetails(inputRecord.FocusEvent);
                    break;
                default:
                    eventDetails = new { type = "None" };
                    break;
            }
            return Serialize(eventDetails);
        }

        private bool GetConsoleMode(uint bit) {
            uint mode;
            Kernel32.GetConsoleMode(inputHandle, out mode);
            return (mode & bit) != 0;
        }

        private void SetConsoleMode(uint bit, bool value) {
            uint mode;
            Kernel32.GetConsoleMode(inputHandle, out mode);
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

        private object GetKeyEventDetails(KEY_EVENT_RECORD keyEventRecord) {
            return new {
                type = keyEventRecord.bKeyDown != 0 ? "KeyPressed" : "KeyReleased",
                scanCode = keyEventRecord.wVirtualScanCode,
                keyCode = keyEventRecord.wVirtualKeyCode,
                unicodeChar = keyEventRecord.uChar.UnicodeChar,
                unicodeCharCode = (int)keyEventRecord.uChar.UnicodeChar,
                asciiCharCode = (int)keyEventRecord.uChar.AsciiChar,
                asciiChar = keyEventRecord.uChar.AsciiChar > 0 && keyEventRecord.uChar.AsciiChar < 127 ? ((char)keyEventRecord.uChar.AsciiChar).ToString() : "",
                repetitions = keyEventRecord.bKeyDown != 0 ? (int?)keyEventRecord.wRepeatCount : null,
                keyType = GetKeyType(keyEventRecord.wVirtualKeyCode),
                shiftPressed = IsModifierKeyPressed(0x10, keyEventRecord.dwControlKeyState),
                ctrlPressed = IsModifierKeyPressed(0x11, keyEventRecord.dwControlKeyState),
                altPressed = IsModifierKeyPressed(0x12, keyEventRecord.dwControlKeyState),
                controlKeys = GetControlKeysDetails(keyEventRecord.dwControlKeyState)
            };
        }

        private object GetMouseEventDetails(MOUSE_EVENT_RECORD mouseEventRecord) {
            var cursorLocation = new {
                row = mouseEventRecord.dwMousePosition.Y,
                column = mouseEventRecord.dwMousePosition.X
            };
            var controlKeys = GetControlKeysDetails(mouseEventRecord.dwControlKeyState);

            switch (mouseEventRecord.dwEventFlags) {
                case DOUBLE_CLICK:
                    return new {
                        type = "MouseDoubleClick",
                        cursorLocation,
                        buttons = GetMouseButtonsDetails(mouseEventRecord.dwButtonState),
                        controlKeys
                    };
                case 0: // button press or release
                    if (mouseEventRecord.dwButtonState != 0) {
                        return new {
                            type = "MouseClick",
                            cursorLocation,
                            buttons = GetMouseButtonsDetails(mouseEventRecord.dwButtonState),
                            controlKeys
                        };
                    } else {
                        return new {
                            type = "MouseButtonReleased",
                            cursorLocation,
                            controlKeys
                        };
                    }
                case MOUSE_HWHEELED:
                    return new {
                        type = "HorizontalWheel",
                        cursorLocation,
                        direction = (short)HiWord((int)mouseEventRecord.dwButtonState) > 0 ? "right" : "left",
                        controlKeys
                    };
                case MOUSE_MOVED:
                    return new {
                        type = "MouseMoved",
                        cursorLocation,
                        controlKeys
                    };
                case MOUSE_WHEELED:
                    return new {
                        type = "VerticalWheel",
                        cursorLocation,
                        direction = (short)HiWord((int)mouseEventRecord.dwButtonState) > 0 ? "forward" : "backwards",
                        controlKeys
                    };
                default:
                    return new { type = "None" };
            }
        }

        private object GetWindowResizedEventDetails(WINDOW_BUFFER_SIZE_RECORD windowBufferSizeEvent) {
            return new {
                type = "WindowResized",
                rows = windowBufferSizeEvent.dwSize.Y,
                columns = windowBufferSizeEvent.dwSize.X
            };
        }

        private object GetWindowFocusEventDetails(FOCUS_EVENT_RECORD focusEvent) {
            return new {
                type = "WindowFocus",
                focused = focusEvent.bSetFocus
            };
        }

        private bool IsModifierKeyPressed(int keyCode, uint controlKeyState) {
            switch (keyCode) {
                case 0x10: // Shift
                    return (controlKeyState & SHIFT_PRESSED) != 0;
                case 0x11: // Ctrl
                    return (controlKeyState & (LEFT_CTRL_PRESSED | RIGHT_CTRL_PRESSED)) != 0;
                case 0x12: // Alt
                    return (controlKeyState & (LEFT_ALT_PRESSED | RIGHT_ALT_PRESSED)) != 0;
                default:
                    return false;
            }
        }

        private string GetKeyType(int keyCode) {
            if (keyCode >= 0x10 && keyCode <= 0x12) return "modifier";
            if ((keyCode >= 0x60 && keyCode <= 0x6F)) return "numpad";
            if (keyCode == 0x5B || keyCode == 0x5C || keyCode == 0x5D || keyCode == 0x2C || keyCode == 0x13) return "system";
            if (keyCode == 0x14 || keyCode == 0x90 || keyCode == 0x91) return "lock";
            if (keyCode >= 0xA6 && keyCode <= 0xAC) return "browser";
            if (keyCode >= 0xB6 && keyCode <= 0xB7 || keyCode == 0xB4) return "application";
            if (keyCode >= 0x41 && keyCode <= 0x5A) return "alphabetic";
            if (keyCode >= 0x30 && keyCode <= 0x39) return "numeric";
            if (keyCode >= 0x70 && keyCode <= 0x87) return "function";
            if (keyCode >= 0x21 && keyCode <= 0x28) return "navigation";
            if (keyCode == 0x20) return "alphabetic";
            if (keyCode == 0x09) return "navigation";
            if (keyCode == 0x08 || keyCode == 0x2E || keyCode == 0x2D) return "edition";
            if (keyCode >= 0xBB && keyCode <= 0xBE) return "punctuation";
            if (keyCode >= 0xAD && keyCode <= 0xB3) return "media";
            return "none";
        }

        private object GetControlKeysDetails(uint controlKeyState) {
            return new {
                capsLockOn = (controlKeyState & CAPSLOCK_ON) != 0,
                enhancedKey = (controlKeyState & ENHANCED_KEY) != 0,
                leftAltPressed = (controlKeyState & LEFT_ALT_PRESSED) != 0,
                leftCtrlPressed = (controlKeyState & LEFT_CTRL_PRESSED) != 0,
                numLockOn = (controlKeyState & NUMLOCK_ON) != 0,
                rightAltPressed = (controlKeyState & RIGHT_ALT_PRESSED) != 0,
                rightCtrlPressed = (controlKeyState & RIGHT_CTRL_PRESSED) != 0,
                scrollLockOn = (controlKeyState & SCROLLLOCK_ON) != 0,
                shiftPressed = (controlKeyState & SHIFT_PRESSED) != 0
            };
        }

        private object GetMouseButtonsDetails(uint buttonState) {
            return new {
                left1Pressed = (buttonState & FROM_LEFT_1ST_BUTTON_PRESSED) != 0,
                left2Pressed = (buttonState & FROM_LEFT_2ND_BUTTON_PRESSED) != 0,
                left3Pressed = (buttonState & FROM_LEFT_3RD_BUTTON_PRESSED) != 0,
                left4Pressed = (buttonState & FROM_LEFT_4TH_BUTTON_PRESSED) != 0,
                rightMostPressed = (buttonState & RIGHTMOST_BUTTON_PRESSED) != 0
            };
        }

        public int CodePage {
            get => (int)GetConsoleCP();
            set => SetConsoleCP((uint)value);
        }

        private string Serialize(object obj) {
            var serializer = new JavaScriptSerializer();
            return serializer.Serialize(obj);
        }
    }
}