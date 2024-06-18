using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EYOC_Team_Score
{
    public partial class OptionsBox : Form
    {
        public OptionsBox()
        {
            InitializeComponent();
            ExportCSS.Checked = Properties.Settings.Default.ExportCSS;
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void OK_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save();
            Close();
        }

        private void ExportCSS_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.ExportCSS = ExportCSS.Checked;
        }
    }
}
