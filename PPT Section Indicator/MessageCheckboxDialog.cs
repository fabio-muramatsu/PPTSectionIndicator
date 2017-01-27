using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PPT_Section_Indicator
{
    public partial class MessageCheckboxDialog : Form
    {

        public MessageCheckboxDialog()
        {
            InitializeComponent();
        }

        public MessageCheckboxDialog(string message) : this()
        {
            MessageLabel.Text = message;
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        public bool ShowDialogForResult()
        {
            ShowDialog();
            return ShowCheckBox.Checked;
        }

        public void SetCheckBoxState(bool state)
        {
            ShowCheckBox.Checked = state;
        }
    }
}
