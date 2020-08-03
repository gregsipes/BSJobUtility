using System;
using System.Collections.Generic;
using System.Drawing;
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
        public event EventHandler ucMouse_Click;

        private const int LED_SIZE = 8;
        private Point INITIAL_LED_LOCATION = new Point(LED_SIZE, LED_SIZE);
        private int LED_OFFSET = 4; // distance (in pixels) between LEDs

        private const int MULTILINE_MINIMUMWIDTH = 250;
        private const int MULTILINE_MINIMUM_HEIGHT = 64;
        private const int MULTILINE_ACTIVITY_LED_X = 4;
        private const int MULTILINE_ACTIVITY_LED_Y = 26;
        private const int MULTILINE_APPNAME_X = 24;
        private const int MULTILINE_APPNAME_Y = 24;
        private const int MULTILINE_LASTRUNTIME_X = 24;
        private const int MULTILINE_LASTRUNTIME_Y = 44;

        private const int SINGLELINE_MINIMUMWIDTH = 500;
        private const int SINGLELINE_MINIMUMHEIGHT = 24;
        private const int SINGLELINE_LASTRUNTIME_X = 400;
        private const int SINGLELINE_ACTIVITY_LED_X = 600;
        private const int SINGLELINE_ACTIVITY_LED_Y = 4;
        private const int SINGLELINE_ALIGNMENT_Y = 4;

        private int NumLEDsToBuild = 10; // default number of leds to create on this control
        private bool SingleLine = false;
        private List<Color> LEDColors;
        private Color LEDActivityColor;

        public AppStatusUserControl()
        {
            InitializeComponent();

            // A parameterless instantiation requires a default set of LEDs to be constructed. 
            //  (This should not normally be invoked except by the main program's InitializeComponent() function)
            CreateLEDs(NumLEDsToBuild);
        }

        public AppStatusUserControl(int numLedsToBuild, bool singleLine)
        {
            InitializeComponent();

            // Create as many LEDs as is specified in constant NumLEDS.
            //   First LED is named LED00, all others are named sequentially.
            //   They are placed in a single row.

            SingleLine = singleLine;
            CreateLEDs(numLedsToBuild);
            //ucMouse_Click += new EventHandler(OnMouse_Click);
            ucMouse_Click += OnMouse_Click;

            // Some dynamics:  Logic to stretch/shape this control based on the number of LEDs, text width and
            //   single-line or multi-line shape.
            UpdateControlArea();
        }

        private void CreateLEDs(int numLedsToBuild)
        {

            Color DefaultColor = Color.FromKnownColor(KnownColor.Desktop);
            LEDColors = new List<Color>();
            LEDActivityColor = DefaultColor;
            NumLEDsToBuild = numLedsToBuild;
            for (int i = 0; i < NumLEDsToBuild; i++)
            {
                PictureBox led = new PictureBox()
                {
                    Name = "LED" + i.ToString("D2"),
                    Size = new Size(8, 8),
                    BorderStyle = BorderStyle.None,
                    Anchor = AnchorStyles.Top | AnchorStyles.Left,
                    BackColor = DefaultColor,
                    Location = new Point(INITIAL_LED_LOCATION.X + i * (LED_OFFSET + LED_SIZE), INITIAL_LED_LOCATION.Y)
                };
                LEDColors.Add(DefaultColor);

                // Add a mouse hover event handler to this control, to be used to display messages associated with any LED/activity
                led.MouseHover += new EventHandler((sender, e) => Mouse_Hover(sender, e));
                this.Controls.Add(led);
            }
        }

        private void UpdateControlArea()
        {
            //  Logic to stretch / shape this control based on the number of LEDs, text width and
            //   single-line or multi-line shape.

            int EndofLEDs = NumLEDsToBuild * (LED_OFFSET + LED_SIZE) + 2 * LED_OFFSET;

            if (SingleLine)
            {
                // Shape 1:  Singleline - Everything is stretched out and a fixed distance apart after LEDs have been placed.

                this.LblAppName.Location = new Point(EndofLEDs, SINGLELINE_ALIGNMENT_Y);
                this.LBLLastRunTime.Location = new Point(SINGLELINE_LASTRUNTIME_X, SINGLELINE_ALIGNMENT_Y);
                this.LEDActivity.Location = new Point(SINGLELINE_ACTIVITY_LED_X, SINGLELINE_ACTIVITY_LED_Y);

                this.Width = this.LEDActivity.Left + this.LEDActivity.Width + LED_OFFSET;
                this.Height = SINGLELINE_MINIMUMHEIGHT;

            }
            else
            {
                // Shape 2:  Multiline
                //   Width is the larger of MINIMUM_MULTILINE_WIDTH or NumLEDsToBuild * (LED_OFFSET + LED_SIZE) + LED_OFFSET
                //   Height is fixed at MINIMUM_MULTILINE_HEIGHT

                this.Width = Math.Max(MULTILINE_MINIMUMWIDTH, EndofLEDs);
                this.Height = MULTILINE_MINIMUM_HEIGHT;

                // LEDS are positioned elsewhere
                // AppName is positioned below LEDs
                // LastRunTime is positioned below AppName
                this.LblAppName.Location = new Point(MULTILINE_APPNAME_X, MULTILINE_APPNAME_Y);
                this.LBLLastRunTime.Location = new Point(MULTILINE_LASTRUNTIME_X, MULTILINE_LASTRUNTIME_Y);
                this.LEDActivity.Location = new Point(MULTILINE_ACTIVITY_LED_X, MULTILINE_ACTIVITY_LED_Y);

            }
        }

        public void ToggleDisplayMode()
        {
            // Toggle between singleline and multiline mode
            SingleLine = !SingleLine;
            UpdateControlArea();
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
            // OPTIMIZATION:  Set color only if it has changed.

            try
            {
                bool ColorChanged = false;
                PictureBox pic = null;
                if (led == LEDs.LEDActivity)
                {
                    if (LEDcolor != LEDActivityColor)
                    {
                        pic = (PictureBox)this.Controls["LEDActivity"];
                        LEDActivityColor = LEDcolor;
                        ColorChanged = true;
                    }
                }
                else
                {
                    if (LEDcolor != LEDColors[ledNum])
                    {
                        pic = (PictureBox)this.Controls["LED" + ledNum.ToString("D2")];
                        LEDColors[ledNum] = LEDcolor;
                        ColorChanged = true;
                    }
                }

                if (ColorChanged)
                {
                    pic.BackColor = LEDcolor;
                }
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
                DateTime dt;
                try
                {
                    dt = DateTime.Parse(LBLLastRunTime.Text);
                } catch (Exception ex)
                {
                    dt = DateTime.MinValue;
                }
                return dt;
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

        private void OnMouse_Click(object sender, EventArgs e)
        {
            // Expose a mouse click to the outside
            ucMouse_Click(sender, e);
        }

        private void LEDActivity_Click(object sender, EventArgs e)
        {
            //
            OnMouse_Click(sender, e);
        }

        private void LblAppName_Click(object sender, EventArgs e)
        {
            //
            OnMouse_Click(sender, e);
        }

        private void LBLLastRunTime_Click(object sender, EventArgs e)
        {
            //
            OnMouse_Click(sender, e);
        }
        #endregion

        private void AppStatusUserControl_Click(object sender, EventArgs e)
        {
            //
            OnMouse_Click(sender, e);
        }
    }
}
