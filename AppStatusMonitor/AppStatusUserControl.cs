using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AppStatusControl
{
    public partial class AppStatusUserControl : UserControl
    {
        public enum LEDs
        {
            LEDActivity,
            LEDStatus
        }

        public int LEDNum;

        private const int LED_SIZE = 8;
        private Point INITIAL_LED_LOCATION = new Point(LED_SIZE, LED_SIZE);
        private int LED_OFFSET = 4; // distance (in pixels) between LEDs
        private int NumLEDsToBuild = 10; // default number of leds to create on this control

        public AppStatusUserControl()
        {
            InitializeComponent();

            // A parameterless instantiation requires a default set of LEDs to be constructed. 
            //  (This should not normally be invoked except by the main program's InitializeComponent() function)
            CreateLEDs(NumLEDsToBuild);
        }

        public AppStatusUserControl(int numLedsToBuild)
        {
            InitializeComponent();

            // Create as many LEDs as is specified in constant NumLEDS.
            //   First LED is named LED00, all others are named sequentially.
            //   They are placed in a single row.

            CreateLEDs(numLedsToBuild);
        }

        private void CreateLEDs(int numLedsToBuild)
        {
            NumLEDsToBuild = numLedsToBuild;
            for (int i = 0; i < NumLEDsToBuild; i++)
            {
                PictureBox led = new PictureBox()
                {
                    Name = "LED" + i.ToString("D2"),
                    Size = new Size(8, 8),
                    BorderStyle = BorderStyle.None,
                    Anchor = AnchorStyles.Top | AnchorStyles.Left,
                    BackColor = Color.FromKnownColor(KnownColor.Desktop),
                    Location = new Point(INITIAL_LED_LOCATION.X + i * (LED_OFFSET + LED_SIZE), INITIAL_LED_LOCATION.Y)
                };

                // Add a mouse hover event handler to this control (TBD as well as an enumeration? We can get it from the control's name)
                led.MouseHover += new EventHandler((sender, e) => Mouse_Hover(sender, e));
                this.Controls.Add(led);
            }
        }

        public void ClearLEDs(int firstLED)
        {
            ClearLEDs(firstLED, NumLEDsToBuild);
        }

        public void ClearLEDs (int firstLED, int lastLED)
        {
            for (int i = firstLED; i < lastLED; i++)  
            {
                SetLEDColor(LEDs.LEDStatus, i, Color.Black);
            }
        }

        public void SetLEDColor(LEDs led, int ledNum, Color LEDcolor)
        {
            // Select the associated bitmap from our resources and stick it in the selected picture box

            try
            {
                PictureBox pic;
                if (led == LEDs.LEDActivity)
                {
                    pic = (PictureBox)this.Controls["LEDActivity"];
                }
                else
                {
                    pic = (PictureBox)this.Controls["LED" + ledNum.ToString("D2")];
                }
                pic.BackColor = LEDcolor;
            }
            catch (Exception ex)
            {
                // TBD - No exception should ever occur here unless we run out of streams or some odd thing like that.            
            }

        }

        public string GetLEDMessage(int LEDNumber)
        {
            PictureBox p = (PictureBox)this.Controls["LED" + LEDnumber.ToString("D2")];
            if (p.Tag != null)
            {
                return (p.Tag.ToString());
            }
            else
            {
                return ("");
            }
        }

        public void SetLEDMessage(int LEDNumber, string msg)
        {
            PictureBox p = (PictureBox)this.Controls["LED" + LEDNumber.ToString("D2")];
            p.Tag = msg;
        }

        public string AppName
        {
            get
            {
                return LblAppName.Text;
            }
            set
            {
                LblAppName.Text = value;
            }
        }

        public DateTime RunTime
        {
            get
            {
                return DateTime.Parse(LBLLastRunTime.Text);
            }
            set
            {
                LBLLastRunTime.Text = value.ToLongDateString() + " " + value.ToLongTimeString();
            }
        }

        public int LEDnumber
        {
            get
            {
                return LEDNum;
            }
        }

#region Events
        private void Mouse_Hover(object sender, EventArgs e)
        {
            // Generate a tool tip if this LED has attached messages
            PictureBox p = (PictureBox)sender;
            string s = p.Name.Substring(p.Name.Length - 2);
            int.TryParse(s, out LEDNum);

            if (p.Tag != null)
            {
                ToolTip tt = new ToolTip
                {
                    IsBalloon = true,
                    ShowAlways = true,
                    ToolTipIcon = ToolTipIcon.Info,
                    AutomaticDelay = 100,
                    AutoPopDelay = 60000
                };
                tt.SetToolTip(p, p.Tag.ToString());
            }
        }
#endregion

    }
}
