using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace BSGlobals
{

    // A general status bar that can be used on any UI.  The goal is to get a consistent status bar
    //   across all apps that automatically displays the app's version number while leaving 
    //   room and versatility for status bar text.
    //
    // The simplest instantiation only needs the parent form identified.
    // The more versatile instantiation allows for multiple segments (2 or more) and 
    //   manual or automatic control of the version control box.
    //
    // This status bar has a minimum of 2 segments.

    public class VersionStatusBar
    {
        private Form ParentForm;
        public StatusStrip Strip;
        public int NumSegments { get; }
        public int VersionSegmentWidth = 80;

        private bool PaintEnabled;

        public VersionStatusBar(Form parentForm)
        {
            // This is the default (simple) status bar: 
            //   2 segments
            //   right-most segment is fixed-size and contains the app's version number

            Strip = new StatusStrip();
            ParentForm = parentForm;
            NumSegments = 2;

            CreateStatusBar(true);
        }

        public VersionStatusBar(Form parentForm, int numSegments, bool addVersion)
        {
            // A more adaptable status bar that allows multiple segments (2 or more) and
            //   optional automatic insertion of the app's version number

            // Instantiate
            Strip = new StatusStrip();
            ParentForm = parentForm;
            NumSegments = (numSegments > 1) ? numSegments : 2;

            CreateStatusBar(addVersion);
        }

        private void CreateStatusBar(bool addVersion)
        {
            try
            {
                // StatusStrip doesn't support a resize event, so use the Paint event to re-paint the strip every time the parent form is resized.
                PaintEnabled = false;
                Strip.Paint += new PaintEventHandler(this.RepaintStrip);

                Strip.SuspendLayout();
                ParentForm.SuspendLayout();

                // Dock at the bottom and add a grip to allow expansion/contraction of parent form
                Strip.Dock = DockStyle.Bottom;
                Strip.GripStyle = ToolStripGripStyle.Visible;
                Strip.SizingGrip = true;

                // Specific, default styling
                //Strip.LayoutStyle = ToolStripLayoutStyle.HorizontalStackWithOverflow;
                Strip.LayoutStyle = ToolStripLayoutStyle.Flow;
                Strip.ShowItemToolTips = true;
                Strip.Stretch = true;

                // Add status bar segements
                for (int i = 0; i < NumSegments; i++)
                {
                    Strip.Items.AddRange(new ToolStripItem[] { new ToolStripStatusLabel() });
                    Strip.Items[i].AutoSize = false;
                    Strip.Items[i].TextAlign = ContentAlignment.MiddleLeft;
                }

                PaintEnabled = true;

                // Last segment is special.  It will normally contain the app's version number.
                // Force the last segment to take up whatever remains of the status bar after repaint
                Strip.Items[NumSegments - 1].AutoSize = true;

                if (addVersion)
                {
                    Strip.Items[NumSegments - 1].TextAlign = ContentAlignment.MiddleRight;
                    //Strip.Items[NumSegments - 1].Alignment = ToolStripItemAlignment.Right;
                    AddVersion();
                }

                // Renable parent layout, and force this status strip to the front in case there are other controls extending the bottom of the parent form.

                ParentForm.Controls.Add(Strip);
                Strip.ResumeLayout(false);
                Strip.PerformLayout();
                Strip.BringToFront();
                ParentForm.ResumeLayout();
                ParentForm.PerformLayout();

            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR - Unable to create status bar:  " + ex.ToString());
            }

        }

        public bool AddVersion()
        {
            try
            {
                // Allow manual insertion of verion
                Strip.Items[NumSegments - 1].Text = "V" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
                return (true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR - Unable to add version to status bar:  " + ex.ToString());
                return (false);
            }
        }

        public bool AddVersion(int segmentNum, string version)
        {
            try
            {
                // Allow manual insertion of version into any segment
                if (segmentNum < NumSegments)
                {
                    Strip.Items[segmentNum].Text = version;
                    return (true);
                }
                return (false);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR - Unable to add version to status bar:  " + ex.ToString());
                return (false);
            }
        }

        public bool AddText(int segmentNum, string s)
        {
            try
            {
                // Adds text to the selected segment
                if (segmentNum < NumSegments)
                {
                    Strip.Items[segmentNum].Text = s;
                    return (true);
                }
                return (false);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR trying to add text to status bar:  " + ex.ToString());
                return (false);
            }
        }

        private void RepaintStrip(object sender, PaintEventArgs e)
        {
            // Resize the status bar here whenever the size is changed
            try
            {
                if (PaintEnabled)
                {
                    // We want to resize all but the last segment to the new client width
                    if (NumSegments > 1)
                    {
                        int newwidth = (ParentForm.ClientSize.Width - VersionSegmentWidth) / (NumSegments - 1);
                        for (int i = 0; i < NumSegments - 1; i++)
                        {
                            Strip.Items[i].Width = newwidth;
                        }
                    }
                    else
                    {
                        Strip.Items[0].Width = ParentForm.ClientSize.Width;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Status Bar Repaint Error: " + ex.ToString());
            }
        }
    }
}
