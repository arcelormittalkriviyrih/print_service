namespace PrintWindowsService
{
    partial class frmMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.notifyIconPrint = new System.Windows.Forms.NotifyIcon(this.components);
            this.contextMenuTray = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.mItemStart = new System.Windows.Forms.ToolStripMenuItem();
            this.mItemStop = new System.Windows.Forms.ToolStripMenuItem();
            this.mItemRestart = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.mItemExit = new System.Windows.Forms.ToolStripMenuItem();
            this.contextMenuTray.SuspendLayout();
            this.SuspendLayout();
            // 
            // notifyIconPrint
            // 
            this.notifyIconPrint.ContextMenuStrip = this.contextMenuTray;
            this.notifyIconPrint.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIconPrint.Icon")));
            this.notifyIconPrint.Text = "Сервис печати этикеток";
            this.notifyIconPrint.Visible = true;
            // 
            // contextMenuTray
            // 
            this.contextMenuTray.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mItemStart,
            this.mItemStop,
            this.mItemRestart,
            this.toolStripSeparator1,
            this.mItemExit});
            this.contextMenuTray.Name = "contextMenuTray";
            this.contextMenuTray.Size = new System.Drawing.Size(156, 98);
            // 
            // mItemStart
            // 
            this.mItemStart.Enabled = false;
            this.mItemStart.Name = "mItemStart";
            this.mItemStart.Size = new System.Drawing.Size(155, 22);
            this.mItemStart.Text = "Запустить";
            this.mItemStart.Click += new System.EventHandler(this.mItemStart_Click);
            // 
            // mItemStop
            // 
            this.mItemStop.Name = "mItemStop";
            this.mItemStop.Size = new System.Drawing.Size(155, 22);
            this.mItemStop.Text = "Остановить";
            this.mItemStop.Click += new System.EventHandler(this.mItemStop_Click);
            // 
            // mItemRestart
            // 
            this.mItemRestart.Name = "mItemRestart";
            this.mItemRestart.Size = new System.Drawing.Size(155, 22);
            this.mItemRestart.Text = "Перезапустить";
            this.mItemRestart.Click += new System.EventHandler(this.mItemRestart_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(152, 6);
            // 
            // mItemExit
            // 
            this.mItemExit.Name = "mItemExit";
            this.mItemExit.Size = new System.Drawing.Size(155, 22);
            this.mItemExit.Text = "Выйти";
            this.mItemExit.Click += new System.EventHandler(this.mItemExit_Click);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Name = "frmMain";
            this.ShowInTaskbar = false;
            this.Text = "Label Print";
            this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
            this.contextMenuTray.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.NotifyIcon notifyIconPrint;
        private System.Windows.Forms.ContextMenuStrip contextMenuTray;
        private System.Windows.Forms.ToolStripMenuItem mItemStart;
        private System.Windows.Forms.ToolStripMenuItem mItemStop;
        private System.Windows.Forms.ToolStripMenuItem mItemRestart;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem mItemExit;
    }
}

