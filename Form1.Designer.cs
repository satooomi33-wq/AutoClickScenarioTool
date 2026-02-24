namespace AutoClickScenarioTool
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button btnPause;
        private System.Windows.Forms.Button btnStop;
        
        private System.Windows.Forms.ToolStrip captureToolStrip;
        private System.Windows.Forms.ToolStripButton tsbDisable;
        private System.Windows.Forms.ToolStripButton tsbMouse;
        private System.Windows.Forms.ToolStripButton tsbKey;
        private System.Windows.Forms.ToolStripLabel tslCaptureStatus;
        private System.Windows.Forms.TextBox txtDataFolder;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.ComboBox cmbFiles;
        private System.Windows.Forms.Button btnSave;
        public System.Windows.Forms.DataGridView dgvScenario;
        private System.Windows.Forms.TextBox txtLog;

        private void InitializeComponent()
        {
            btnStart = new Button();
            btnPause = new Button();
            btnStop = new Button();
            captureToolStrip = new ToolStrip();
            tsbDisable = new ToolStripButton();
            tsbMouse = new ToolStripButton();
            tsbKey = new ToolStripButton();
            tslCaptureStatus = new ToolStripLabel();
            txtDataFolder = new TextBox();
            btnBrowse = new Button();
            cmbFiles = new ComboBox();
            btnSave = new Button();
            dgvScenario = new DataGridView();
            txtLog = new TextBox();
            ((System.ComponentModel.ISupportInitialize)dgvScenario).BeginInit();
            SuspendLayout();
            // 
            // btnStart
            // 
            btnStart.Font = new Font("Segoe UI Symbol", 9.857143F);
            btnStart.Location = new Point(10, 129);
            btnStart.Name = "btnStart";
            btnStart.Size = new Size(75, 45);
            btnStart.TabIndex = 0;
            btnStart.Text = "▶";
            btnStart.UseVisualStyleBackColor = true;
            btnStart.Click += btnStart_Click;
            // 
            // btnPause
            // 
            btnPause.Font = new Font("Segoe UI Symbol", 9.857143F);
            btnPause.Location = new Point(91, 129);
            btnPause.Name = "btnPause";
            btnPause.Size = new Size(75, 45);
            btnPause.TabIndex = 1;
            btnPause.Text = "⏸";
            btnPause.UseVisualStyleBackColor = true;
            btnPause.Click += btnPause_Click;
            // 
            // btnStop
            // 
            btnStop.Font = new Font("Segoe UI Symbol", 9.857143F);
            btnStop.Location = new Point(172, 129);
            btnStop.Name = "btnStop";
            btnStop.Size = new Size(75, 45);
            btnStop.TabIndex = 2;
            btnStop.Text = "⏹";
            btnStop.UseVisualStyleBackColor = true;
            btnStop.Click += btnStop_Click;
            // 
            // txtDataFolder
            // 
            txtDataFolder.Location = new Point(12, 42);
            txtDataFolder.Name = "txtDataFolder";
            txtDataFolder.Size = new Size(600, 35);
            txtDataFolder.TabIndex = 4;
            // 
            // btnBrowse
            // 
            btnBrowse.Location = new Point(618, 40);
            btnBrowse.Name = "btnBrowse";
            btnBrowse.Size = new Size(75, 37);
            btnBrowse.TabIndex = 5;
            btnBrowse.Text = "参照";
            btnBrowse.UseVisualStyleBackColor = true;
            btnBrowse.Click += btnBrowse_Click;
            // 
            // cmbFiles
            // 
            cmbFiles.FormattingEnabled = true;
            cmbFiles.Location = new Point(11, 97);
            cmbFiles.Name = "cmbFiles";
            cmbFiles.Size = new Size(300, 38);
            cmbFiles.TabIndex = 6;
            cmbFiles.SelectedIndexChanged += cmbFiles_SelectedIndexChanged;
            cmbFiles.TextChanged += cmbFiles_TextChanged;
            // 
            // btnSave
            // 
            btnSave.Location = new Point(493, 93);
            btnSave.Name = "btnSave";
            btnSave.Size = new Size(85, 42);
            btnSave.TabIndex = 8;
            btnSave.Text = "保存";
            btnSave.UseVisualStyleBackColor = true;
            btnSave.Click += btnSave_Click;
            // 
            // dgvScenario
            // 
            dgvScenario.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvScenario.Location = new Point(12, 223);
            dgvScenario.Name = "dgvScenario";
            dgvScenario.RowHeadersWidth = 72;
            dgvScenario.RowTemplate.Height = 25;
            dgvScenario.Size = new Size(1167, 636);
            dgvScenario.TabIndex = 9;
            dgvScenario.CellValueChanged += dgvScenario_CellValueChanged;
            dgvScenario.RowsAdded += dgvScenario_RowsChanged;
            dgvScenario.RowsRemoved += dgvScenario_RowsChanged;
            dgvScenario.UserAddedRow += dgvScenario_UserAddedRow;
            dgvScenario.UserDeletedRow += dgvScenario_RowsChanged;
            // 
            // txtLog
            // 
            txtLog.Location = new Point(12, 910);
            txtLog.Multiline = true;
            txtLog.Name = "txtLog";
            txtLog.ReadOnly = true;
            txtLog.ScrollBars = ScrollBars.Vertical;
            txtLog.Size = new Size(1167, 88);
            txtLog.TabIndex = 10;
            // 
            // captureToolStrip
            // 
            captureToolStrip.Items.AddRange(new ToolStripItem[] { tsbDisable, tsbMouse, tsbKey, new ToolStripSeparator(), tslCaptureStatus });
            captureToolStrip.Location = new Point(0, 0);
            captureToolStrip.Name = "captureToolStrip";
            captureToolStrip.Size = new Size(1211, 30);
            captureToolStrip.TabIndex = 100;
            captureToolStrip.Text = "captureToolStrip";

            // tsbDisable
            tsbDisable.Text = "無効";
            tsbDisable.CheckOnClick = true;
            tsbDisable.Checked = true;
            tsbDisable.Click += tsbDisable_Click;

            // tsbMouse
            tsbMouse.Text = "座標抽出";
            tsbMouse.CheckOnClick = true;
            tsbMouse.Click += tsbMouse_Click;

            // tsbKey
            tsbKey.Text = "キー抽出";
            tsbKey.CheckOnClick = true;
            tsbKey.Click += tsbKey_Click;

            // tslCaptureStatus
            tslCaptureStatus.Text = "キャプチャ: 無効";

            // add toolstrip first
            Controls.Add(captureToolStrip);

            
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(12F, 30F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1211, 980);
            Controls.Add(txtLog);
            Controls.Add(dgvScenario);
            Controls.Add(btnSave);
            Controls.Add(cmbFiles);
            Controls.Add(btnBrowse);
            Controls.Add(txtDataFolder);
            Controls.Add(btnStop);
            Controls.Add(btnPause);
            Controls.Add(btnStart);
            Name = "Form1";
            Text = "AutoClickScenarioTool";
            ((System.ComponentModel.ISupportInitialize)dgvScenario).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
    }
}
