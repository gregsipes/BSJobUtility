using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

// This is an extension of the Windows Combobox.  It provides an additional function to fire a 
//   "DataEntryComplete" event whenever the combobox loses focus or when the ENTER key is depressed.
//   This functionality allows us to process the user entry immediately, if so desired.

// NOTES: To create a similar control, create a naked UserControl (don't put anything in it).
// Change the in
namespace IS_Inventory
{
    public partial class ComboBoxUnlocked : ComboBox
    {

        // Public event handlers
        public event EventHandler DataEntryComplete;

        // Public properties.
        public ComboBoxUnlocked()
        {
            InitializeComponent();
        }

        private void ComboBoxUnlocked_Leave(object sender, EventArgs e)
        {
            // Fire a DataComplete event when the combobox loses focus.
            // This event (Leave) will fire after DataEntryComplete fires.
            DataEntryComplete?.Invoke(this, (EventArgs)e);
        }

        private void ComboBoxUnlocked_KeyDown(object sender, KeyEventArgs e)
        {
            // Fire a DataComplete event if the ENTER key was depressed.
            // This event (KeyDown) will fire after DataEntryComplete fires.
            if (e.KeyValue == 13)
            {
                DataEntryComplete?.Invoke(this, (EventArgs)e);
            }
        }
    }
}
