using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PPT_Section_Indicator
{
    public partial class ProgressDialogBox : Form
    {
        private Action dialogBoxShownCallback;

        public ProgressDialogBox()
        {
            InitializeComponent();
        }

        private void ProgressDialogBox_Shown(object sender, EventArgs e)
        {
            Debug.WriteLine("Progress dialog shown");
            dialogBoxShownCallback();
        }

        public void SetDialogBoxShownCallback(Action callback)
        {
            dialogBoxShownCallback = callback;
        }

        public void UpdateProgressMessage(int current, int total)
        {
            this.Invoke(new Action(() => ProgressSecondaryMessageLabel.Text = PROGRESS_MESSAGE + current + " of " + total));
        }
    }
}
