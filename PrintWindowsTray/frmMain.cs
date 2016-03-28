using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;

namespace PrintWindowsService
{
    public partial class frmMain : Form
    {
        private PrintJobs pJobs;

        public frmMain()
        {
            InitializeComponent();
            this.ShowInTaskbar = false;
            this.Visible = false;
            pJobs = new PrintJobs();
            pJobs.StartJob();
        }

        private void mItemStart_Click(object sender, EventArgs e)
        {
            pJobs.StartJob();
            this.mItemStart.Enabled = false;
            this.mItemStop.Enabled = true;
        }

        private void mItemStop_Click(object sender, EventArgs e)
        {
            pJobs.StopJob();
            this.mItemStart.Enabled = true;
            this.mItemStop.Enabled = false;
        }

        private void mItemRestart_Click(object sender, EventArgs e)
        {
            Application.Restart();
        }

        private void mItemExit_Click(object sender, EventArgs e)
        {
            pJobs.StopJob();
            Application.Exit();
        }
    }
}
