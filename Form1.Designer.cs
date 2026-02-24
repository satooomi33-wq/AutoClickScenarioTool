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
        private System.Windows.Forms.Label lblDefaultDelay;
        private System.Windows.Forms.TextBox txtDefaultDelay;
        private System.Windows.Forms.Label lblDefaultPressDuration;
        private System.Windows.Forms.TextBox txtDefaultPressDuration;
        private System.Windows.Forms.Button btnSaveDefaults;
        private System.Windows.Forms.TextBox txtDataFolder;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.ComboBox cmbFiles;
        private System.Windows.Forms.Button btnSave;
        public System.Windows.Forms.DataGridView dgvScenario;
        private System.Windows.Forms.TextBox txtLog;
        private System.Windows.Forms.Label lblHumanizeRange;
        private System.Windows.Forms.TextBox txtHumanizeLower;
        private System.Windows.Forms.TextBox txtHumanizeUpper;
        private System.Windows.Forms.Button btnToggleHumanize;

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
            lblDefaultDelay = new Label();
            txtDefaultDelay = new TextBox();
            lblDefaultPressDuration = new Label();
            txtDefaultPressDuration = new TextBox();
            btnSaveDefaults = new Button();
            lblHumanizeRange = new Label();
            txtHumanizeLower = new TextBox();
            txtHumanizeUpper = new TextBox();
            btnToggleHumanize = new Button();
            captureToolStrip.SuspendLayout();
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
            // captureToolStrip
            // 
            captureToolStrip.ImageScalingSize = new Size(28, 28);
            captureToolStrip.Items.AddRange(new ToolStripItem[] { tsbDisable, tsbMouse, tsbKey, tslCaptureStatus });
            captureToolStrip.Location = new Point(0, 0);
            captureToolStrip.Name = "captureToolStrip";
            captureToolStrip.Size = new Size(1273, 40);
            captureToolStrip.TabIndex = 100;
            captureToolStrip.Text = "captureToolStrip";
            // 
            // tsbDisable
            // 
            tsbDisable.Checked = true;
            tsbDisable.CheckOnClick = true;
            tsbDisable.CheckState = CheckState.Checked;
            tsbDisable.Name = "tsbDisable";
            tsbDisable.Size = new Size(59, 34);
            tsbDisable.Text = "無効";
            tsbDisable.Click += tsbDisable_Click;
            // 
            // tsbMouse
            // 
            tsbMouse.CheckOnClick = true;
            tsbMouse.Name = "tsbMouse";
            tsbMouse.Size = new Size(101, 34);
            tsbMouse.Text = "座標抽出";
            tsbMouse.Click += tsbMouse_Click;
            // 
            // tsbKey
            // 
            tsbKey.CheckOnClick = true;
            tsbKey.Name = "tsbKey";
            tsbKey.Size = new Size(89, 34);
            tsbKey.Text = "キー抽出";
            tsbKey.Click += tsbKey_Click;
            // 
            // tslCaptureStatus
            // 
            tslCaptureStatus.Name = "tslCaptureStatus";
            tslCaptureStatus.Size = new Size(142, 34);
            tslCaptureStatus.Text = "キャプチャ: 無効";
            // 
            // txtDataFolder
            // 
            txtDataFolder.Location = new Point(12, 50);
            txtDataFolder.Name = "txtDataFolder";
            txtDataFolder.Size = new Size(600, 35);
            txtDataFolder.TabIndex = 4;
            // 
            // btnBrowse
            // 
            btnBrowse.Location = new Point(618, 43);
            btnBrowse.Name = "btnBrowse";
            btnBrowse.Size = new Size(75, 42);
            btnBrowse.TabIndex = 5;
            btnBrowse.Text = "参照";
            btnBrowse.UseVisualStyleBackColor = true;
            btnBrowse.Click += btnBrowse_Click;
            // 
            // cmbFiles
            // 
            cmbFiles.FormattingEnabled = true;
            cmbFiles.Location = new Point(699, 43);
            cmbFiles.Name = "cmbFiles";
            cmbFiles.Size = new Size(367, 38);
            cmbFiles.TabIndex = 6;
            cmbFiles.SelectedIndexChanged += cmbFiles_SelectedIndexChanged;
            cmbFiles.TextChanged += cmbFiles_TextChanged;
            // 
            // btnSave
            // 
            btnSave.Location = new Point(1072, 47);
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
            dgvScenario.Location = new Point(12, 187);
            dgvScenario.Name = "dgvScenario";
            dgvScenario.RowHeadersWidth = 72;
            dgvScenario.RowTemplate.Height = 25;
            dgvScenario.Size = new Size(1167, 537);
            dgvScenario.TabIndex = 9;
            dgvScenario.CellValueChanged += dgvScenario_CellValueChanged;
            dgvScenario.RowsAdded += dgvScenario_RowsChanged;
            dgvScenario.RowsRemoved += dgvScenario_RowsChanged;
            dgvScenario.UserAddedRow += dgvScenario_UserAddedRow;
            dgvScenario.UserDeletedRow += dgvScenario_RowsChanged;
            // 
            // txtLog
            // 
            txtLog.Location = new Point(10, 746);
            txtLog.Multiline = true;
            txtLog.Name = "txtLog";
            txtLog.ReadOnly = true;
            txtLog.ScrollBars = ScrollBars.Vertical;
            txtLog.Size = new Size(1167, 222);
            txtLog.TabIndex = 10;
            // 
            // lblDefaultDelay
            // 
            lblDefaultDelay.Location = new Point(531, 92);
            lblDefaultDelay.Name = "lblDefaultDelay";
            lblDefaultDelay.Size = new Size(180, 35);
            lblDefaultDelay.TabIndex = 101;
            lblDefaultDelay.Text = "デフォルト 遅延(ms)";
            // 
            // txtDefaultDelay
            // 
            txtDefaultDelay.Location = new Point(717, 90);
            txtDefaultDelay.Name = "txtDefaultDelay";
            txtDefaultDelay.Size = new Size(60, 35);
            txtDefaultDelay.TabIndex = 11;
            txtDefaultDelay.Text = "500";
            // 
            // lblDefaultPressDuration
            // 
            lblDefaultPressDuration.Location = new Point(799, 92);
            lblDefaultPressDuration.Name = "lblDefaultPressDuration";
            lblDefaultPressDuration.Size = new Size(180, 33);
            lblDefaultPressDuration.TabIndex = 102;
            lblDefaultPressDuration.Text = "デフォルト 押下時間(ms)";
            // 
            // txtDefaultPressDuration
            // 
            txtDefaultPressDuration.Location = new Point(985, 90);
            txtDefaultPressDuration.Name = "txtDefaultPressDuration";
            txtDefaultPressDuration.Size = new Size(66, 35);
            txtDefaultPressDuration.TabIndex = 12;
            txtDefaultPressDuration.Text = "100";
            // 
            // btnSaveDefaults
            // 
            btnSaveDefaults.Location = new Point(1057, 90);
            btnSaveDefaults.Name = "btnSaveDefaults";
            btnSaveDefaults.Size = new Size(100, 35);
            btnSaveDefaults.TabIndex = 13;
            btnSaveDefaults.Text = "初期値保存";
            btnSaveDefaults.UseVisualStyleBackColor = true;
            btnSaveDefaults.Click += btnSaveDefaults_Click;
            // 
            // lblHumanizeRange
            // 
            lblHumanizeRange.Location = new Point(650, 92);
            lblHumanizeRange.Name = "lblHumanizeRange";
            lblHumanizeRange.Size = new Size(100, 23);
            lblHumanizeRange.TabIndex = 0;
            lblHumanizeRange.Text = "擬人化範囲(ms±)";
            // 
            // txtHumanizeLower
            // 
            txtHumanizeLower.Location = new Point(760, 90);
            txtHumanizeLower.Name = "txtHumanizeLower";
            txtHumanizeLower.Size = new Size(60, 35);
            txtHumanizeLower.TabIndex = 0;
            // 
            // txtHumanizeUpper
            // 
            txtHumanizeUpper.Location = new Point(830, 90);
            txtHumanizeUpper.Name = "txtHumanizeUpper";
            txtHumanizeUpper.Size = new Size(60, 35);
            txtHumanizeUpper.TabIndex = 0;
            // 
            // btnToggleHumanize
            // 
            btnToggleHumanize.Location = new Point(900, 90);
            btnToggleHumanize.Name = "btnToggleHumanize";
            btnToggleHumanize.Size = new Size(100, 35);
            btnToggleHumanize.TabIndex = 0;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(12F, 30F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1273, 980);
            Controls.Add(captureToolStrip);
            Controls.Add(txtLog);
            Controls.Add(dgvScenario);
            Controls.Add(btnSave);
            Controls.Add(cmbFiles);
            Controls.Add(btnBrowse);
            Controls.Add(txtDataFolder);
            Controls.Add(lblDefaultDelay);
            Controls.Add(txtDefaultDelay);
            Controls.Add(lblDefaultPressDuration);
            Controls.Add(txtDefaultPressDuration);
            Controls.Add(btnSaveDefaults);
            Controls.Add(lblHumanizeRange);
            Controls.Add(txtHumanizeLower);
            Controls.Add(txtHumanizeUpper);
            Controls.Add(btnToggleHumanize);
            Controls.Add(btnStop);
            Controls.Add(btnPause);
            Controls.Add(btnStart);
            Name = "Form1";
            Text = "AutoClickScenarioTool";
            captureToolStrip.ResumeLayout(false);
            captureToolStrip.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dgvScenario).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
    }
}
