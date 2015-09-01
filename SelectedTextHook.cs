using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace dotSwitcher
{
    public class SelectedTextHook
    {
        private System.Windows.Forms.Timer _timer = new System.Windows.Forms.Timer();

        private void Form2_Load(object sender, EventArgs e)
        {
            //Initialize a timer to update the control text.
            _timer.Interval = 1000;
            _timer.Tick += new EventHandler(_timer_Tick);
        }
        void _timer_Tick(object sender, EventArgs e)
        {
            try
            {
               // textBox1.Text = GetTextFromFocusedControl();
                Console.WriteLine(GetTextFromFocusedControl());
            }
            catch (Exception exp)
            {
                Console.WriteLine(exp.Message);
                //textBox1.Text += exp.Message;
            }
        }

        //Start to monitor and show the text of the related control.
        private void button1_Click(object sender, EventArgs e)
        {
            _timer.Start();
        }

        //Get the text of the focused control
        private string GetTextFromFocusedControl()
        {
            try
            {
                int activeWinPtr = GetForegroundWindow().ToInt32();
                int activeThreadId = 0, processId;
                activeThreadId = GetWindowThreadProcessId(activeWinPtr, out processId);
                int currentThreadId = GetCurrentThreadId();
                if (activeThreadId != currentThreadId)
                    AttachThreadInput(activeThreadId, currentThreadId, true);
                IntPtr activeCtrlId = GetFocus();

                return GetText(activeCtrlId);
            }
            catch (Exception exp)
            {
                return exp.Message;
            }
        }

        //Get the text of the control at the mouse position
        private string GetTextFromControlAtMousePosition()
        {
            try
            {
                Point p;
                if (GetCursorPos(out p))
                {
                    IntPtr ptr = WindowFromPoint(p);
                    if (ptr != IntPtr.Zero)
                    {
                        return GetText(ptr);
                    }
                }
                return "";
            }
            catch (Exception exp)
            {
                return exp.Message;
            }
        }

        //Get the text of a control with its handle
        private string GetText(IntPtr handle)
        {
            int maxLength = 100;
            IntPtr buffer = Marshal.AllocHGlobal((maxLength + 1) * 2);
            SendMessageW(handle, WM_GETTEXT, maxLength, buffer);
            string w = Marshal.PtrToStringUni(buffer);
            Marshal.FreeHGlobal(buffer);
            return w;
        }

        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out Point pt);

        [DllImport("user32.dll", EntryPoint = "WindowFromPoint", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr WindowFromPoint(Point pt);

        [DllImport("user32.dll", EntryPoint = "SendMessageW")]
        public static extern int SendMessageW([InAttribute] System.IntPtr hWnd, int Msg, int wParam, IntPtr lParam);
        public const int WM_GETTEXT = 13;

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        internal static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        internal static extern IntPtr GetFocus();

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern int GetWindowThreadProcessId(int handle, out int processId);

        [DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        internal static extern int AttachThreadInput(int idAttach, int idAttachTo, bool fAttach);
        [DllImport("kernel32.dll")]
        internal static extern int GetCurrentThreadId();

        [DllImport("user32", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern int GetWindowText(IntPtr hWnd, [Out, MarshalAs(UnmanagedType.LPTStr)] StringBuilder lpString, int nMaxCount);
    }
}