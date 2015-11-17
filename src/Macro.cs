using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace AutoTypeEX
{
    public static class Macro
    {
        //===================================================================== API
        [DllImport("user32.dll")] private static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

        //===================================================================== VARIABLES
        private static Random _randGen = new Random();

        //===================================================================== FUNCTIONS
        private static int Rand(int low, int high)
        {
            return _randGen.Next(low, high + 1);
        }

        public static void Wait(long ms)
        {
            // set now time
            long time = DateTime.Now.Ticks;
            // loop until specified wait time has passed
            do
            {
                Application.DoEvents();
            } while (time + ms * 10000 > DateTime.Now.Ticks);
        }

        public static void TypeText(string text, int minSpeed = 10, int maxSpeed = 35)
        {
            foreach (char letter in text)
            {
                char formattedLetter = letter;
                byte code;
                bool needShift = false;
                // if uppercase, need shift
                if ((byte)letter >= 65 && (byte)letter <= 90)
                {
                    formattedLetter = char.ToUpper(letter);
                    needShift = true;
                }
                else if ((byte)letter >= 97 && (byte)letter <= 122)
                {
                    formattedLetter = char.ToUpper(letter);
                }
                // special characters
                switch (formattedLetter)
                {
                    case '`': code = 192; break;
                    case '~': code = 192; needShift = true; break;
                    case '!': code = 49; needShift = true; break;
                    case '@': code = 50; needShift = true; break;
                    case '#': code = 51; needShift = true; break;
                    case '$': code = 52; needShift = true; break;
                    case '%': code = 53; needShift = true; break;
                    case '^': code = 54; needShift = true; break;
                    case '&': code = 55; needShift = true; break;
                    case '*': code = 56; needShift = true; break;
                    case '(': code = 57; needShift = true; break;
                    case ')': code = 48; needShift = true; break;

                    case '-': code = 189; break;
                    case '_': code = 189; needShift = true; break;
                    case '=': code = 187; break;
                    case '+': code = 187; needShift = true; break;

                    case '[': code = 219; break;
                    case '{': code = 219; needShift = true; break;
                    case '\\': code = 220; break;
                    case '|': code = 220; needShift = true; break;
                    case ']': code = 221; break;
                    case '}': code = 221; needShift = true; break;

                    case ';': code = 186; break;
                    case ':': code = 186; needShift = true; break;
                    // need more symbols
                    case '\'': code = 222; break;
                    case '\"': code = 222; needShift = true; break;

                    case ',': code = 188; break;
                    case '<': code = 188; needShift = true; break;
                    case '.': code = 190; break;
                    case '>': code = 190; needShift = true; break;
                    case '/': code = 191; break;
                    case '?': code = 191; needShift = true; break;

                    default: code = (byte)formattedLetter; break;
                }
                // typing procedure
                if (needShift) keybd_event(0xA0, 0, 0, 0); // If need shift, hold shift
                keybd_event(code, 0, 0, 0);
                keybd_event(code, 0, 2, 0);
                if (needShift) keybd_event(0xA0, 0, 2, 0); // If need shift, release shift
                Wait(Rand(minSpeed, maxSpeed));
            }
        }
    }
}
