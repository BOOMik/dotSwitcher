using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Automation.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace dotSwitcher
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
         private readonly Timer _timer = new Timer();

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
                var text = getSelectedText();
                if (string.IsNullOrEmpty(text)) text = getSelectedText1();
                //textBox1.Text = GetTextFromFocusedControl();
                textBox1.Text = text;
            }
            catch (Exception exp)
            {
                textBox1.Text += exp.Message;
            }
        }

        //work in vs
        private string getSelectedText()
        {
            var element = AutomationElement.FocusedElement;

            if (element != null)
            {
                object pattern;
                if (element.TryGetCurrentPattern(TextPattern.Pattern, out pattern))
                {
                    var tp = (TextPattern)pattern;
                    var sb = new StringBuilder();

                    foreach (var r in tp.GetSelection())
                    {
                        sb.AppendLine(r.GetText(-1));
                    }

                    var selectedText = sb.ToString();
                    return selectedText;
                }
            }
            return null;
        }

        //work in notepad
        private string getSelectedText1()
        {
            IntPtr hWnd = GetForegroundWindow();

            uint processId;

            uint activeThreadId = GetWindowThreadProcessId(hWnd, out processId);

            uint currentThreadId = GetCurrentThreadId();

            AttachThreadInput(activeThreadId, currentThreadId, true);

            IntPtr focusedHandle = GetFocus();

            AttachThreadInput(activeThreadId, currentThreadId, false);

            int len = SendMessage(focusedHandle, WM_GETTEXTLENGTH, 0, null);

            StringBuilder sb = new StringBuilder(len);
            int numChars = SendMessage(focusedHandle, WM_GETTEXT, len + 1, sb);
            int start, next;

            SendMessage(focusedHandle, EM_GETSEL, out start, out next);
            if (len > start && start <= 0 && next <= start) return null;
            string selectedText = sb.ToString().Substring(start, next - start);
            return selectedText;
        }

        private string getSelectedText2()
        {
            System.Drawing.Point mouse = System.Windows.Forms.Cursor.Position; // use Windows forms mouse code instead of WPF
            AutomationElement element = AutomationElement.FromPoint(new System.Windows.Point(mouse.X, mouse.Y));
            if (element == null)
            {
                // no element under mouse
                return null;
            }

            Console.WriteLine("Element at position " + mouse + " is '" + element.Current.Name + "'");

            object pattern;
            // the "Value" pattern is supported by many application (including IE & FF)
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out pattern))
            {
                ValuePattern valuePattern = (ValuePattern)pattern;
                //Console.WriteLine(" Value=" + valuePattern.Current.Value);
                return valuePattern.Current.Value;
            }

            // the "Text" pattern is supported by some applications (including Notepad)and returns the current selection for example
            if (element.TryGetCurrentPattern(TextPattern.Pattern, out pattern))
            {
                TextPattern textPattern = (TextPattern)pattern;
                foreach (TextPatternRange range in textPattern.GetSelection())
                {
                    //Console.WriteLine(" SelectionRange=" + range.GetText(-1));
                    return range.GetText(-1);
                }
            }
            return null;
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
                int currentThreadId = (int) GetCurrentThreadId();
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


// declarations required for Win32 functions

[DllImport("user32.dll")]

static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);


[DllImport("user32.dll", SetLastError = true)]

static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);


[DllImport("user32.dll")]

static extern bool AttachThreadInput(uint idAttach, uint idAttachTo,

bool fAttach);


[DllImport("user32.dll")]

static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, StringBuilder lParam);

// second overload of SendMessage

[DllImport("user32.dll")]

static extern int SendMessage(IntPtr hWnd, uint Msg, out int wParam, out int lParam);


const uint WM_GETTEXTLENGTH = 0x0E;

const uint EM_GETSEL = 0xB0;

        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out Point pt);

        [DllImport("user32.dll", EntryPoint = "WindowFromPoint", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr WindowFromPoint(Point pt);

        [DllImport("user32.dll", EntryPoint = "SendMessageW")]
        public static extern int SendMessageW([In] System.IntPtr hWnd, int Msg, int wParam, IntPtr lParam);
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
        internal static extern uint GetCurrentThreadId();

        [DllImport("user32", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern int GetWindowText(IntPtr hWnd, [Out, MarshalAs(UnmanagedType.LPTStr)] StringBuilder lpString, int nMaxCount);
    }
    
}
