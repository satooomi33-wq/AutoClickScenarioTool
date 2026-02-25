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
        private System.Windows.Forms.ToolStripButton tsbScanCode;
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
        private System.Windows.Forms.Label lblHumanizeSeparator;
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
            tsbScanCode = new ToolStripButton();
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
            lblHumanizeSeparator = new Label();
            txtHumanizeUpper = new TextBox();
            btnToggleHumanize = new Button();
            captureToolStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvScenario).BeginInit();
            SuspendLayout();
            // 
            // btnStart
            // 
            btnStart.Font = new Font("Segoe UI Symbol", 8.142858F);
            btnStart.Location = new Point(10, 137);
            btnStart.Name = "btnStart";
            btnStart.Size = new Size(75, 37);
            btnStart.TabIndex = 0;
            btnStart.Text = "▶";
            btnStart.UseVisualStyleBackColor = true;
            btnStart.Click += btnStart_Click;
            // 
            // btnPause
            // 
            btnPause.Font = new Font("Segoe UI Symbol", 8.142858F);
            btnPause.Location = new Point(91, 137);
            btnPause.Name = "btnPause";
            btnPause.Size = new Size(75, 37);
            btnPause.TabIndex = 1;
            btnPause.Text = "⏸";
            btnPause.UseVisualStyleBackColor = true;
            btnPause.Click += btnPause_Click;
            // 
            // btnStop
            // 
            btnStop.Font = new Font("Segoe UI Symbol", 8.142858F);
            btnStop.Location = new Point(172, 137);
            btnStop.Name = "btnStop";
            btnStop.Size = new Size(75, 37);
            btnStop.TabIndex = 2;
            btnStop.Text = "⏹";
            btnStop.UseVisualStyleBackColor = true;
            btnStop.Click += btnStop_Click;
            // 
            // captureToolStrip
            // 
            captureToolStrip.ImageScalingSize = new Size(28, 28);
            captureToolStrip.Items.AddRange(new ToolStripItem[] { tsbDisable, tsbMouse, tsbKey, tsbScanCode, tslCaptureStatus });
            captureToolStrip.Location = new Point(0, 0);
            captureToolStrip.Name = "captureToolStrip";
            captureToolStrip.Size = new Size(1227, 40);
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
            // tsbScanCode
            // 
            tsbScanCode.CheckOnClick = true;
            tsbScanCode.Name = "tsbScanCode";
            tsbScanCode.Size = new Size(80, 34);
            tsbScanCode.Text = "SC";
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
            btnSave.Location = new Point(1072, 39);
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
            dgvScenario.Location = new Point(12, 200);
            dgvScenario.Name = "dgvScenario";
            dgvScenario.RowHeadersWidth = 72;
            dgvScenario.RowTemplate.Height = 25;
            dgvScenario.Size = new Size(1202, 520);
            dgvScenario.TabIndex = 9;
            dgvScenario.CellValueChanged += dgvScenario_CellValueChanged;
            dgvScenario.CellValidating += DgvScenario_CellValidating;
            dgvScenario.RowsAdded += dgvScenario_RowsChanged;
            dgvScenario.RowsRemoved += dgvScenario_RowsChanged;
            dgvScenario.UserAddedRow += dgvScenario_UserAddedRow;
            dgvScenario.UserDeletedRow += dgvScenario_RowsChanged;
            // 
            // txtLog
            // 
            txtLog.Location = new Point(10, 730);
            txtLog.Multiline = true;
            txtLog.Name = "txtLog";
            txtLog.ReadOnly = true;
            txtLog.ScrollBars = ScrollBars.Vertical;
            txtLog.Size = new Size(1204, 230);
            txtLog.TabIndex = 10;
            // 
            // lblDefaultDelay
            // 
            lblDefaultDelay.Location = new Point(409, 93);
            lblDefaultDelay.Name = "lblDefaultDelay";
            lblDefaultDelay.Size = new Size(180, 35);
            lblDefaultDelay.TabIndex = 101;
            lblDefaultDelay.Text = "デフォルト 遅延(ms)";
            lblDefaultDelay.TextAlign = ContentAlignment.MiddleRight;
            // 
            // txtDefaultDelay
            // 
            txtDefaultDelay.Location = new Point(595, 90);
            txtDefaultDelay.Name = "txtDefaultDelay";
            txtDefaultDelay.Size = new Size(66, 35);
            txtDefaultDelay.TabIndex = 11;
            txtDefaultDelay.Text = "500";
            // 
            // lblDefaultPressDuration
            // 
            lblDefaultPressDuration.Location = new Point(677, 92);
            lblDefaultPressDuration.Name = "lblDefaultPressDuration";
            lblDefaultPressDuration.Size = new Size(229, 33);
            lblDefaultPressDuration.TabIndex = 102;
            lblDefaultPressDuration.Text = "デフォルト 押下時間(ms)";
            lblDefaultPressDuration.TextAlign = ContentAlignment.MiddleRight;
            // 
            // txtDefaultPressDuration
            // 
            txtDefaultPressDuration.Location = new Point(912, 93);
            txtDefaultPressDuration.Name = "txtDefaultPressDuration";
            txtDefaultPressDuration.Size = new Size(66, 35);
            txtDefaultPressDuration.TabIndex = 12;
            txtDefaultPressDuration.Text = "100";
            // 
            // btnSaveDefaults
            // 
            btnSaveDefaults.Location = new Point(997, 87);
            btnSaveDefaults.Name = "btnSaveDefaults";
            btnSaveDefaults.Size = new Size(131, 35);
            btnSaveDefaults.TabIndex = 13;
            btnSaveDefaults.Text = "初期値保存";
            btnSaveDefaults.UseVisualStyleBackColor = true;
            btnSaveDefaults.Click += btnSaveDefaults_Click;
            // 
            // lblHumanizeRange
            // 
            lblHumanizeRange.Location = new Point(9, 93);
            lblHumanizeRange.Name = "lblHumanizeRange";
            lblHumanizeRange.Size = new Size(183, 35);
            lblHumanizeRange.TabIndex = 14;
            lblHumanizeRange.Text = "擬人化範囲(ms±)";
            lblHumanizeRange.TextAlign = ContentAlignment.MiddleRight;
            // 
            // txtHumanizeLower
            // 
            txtHumanizeLower.Location = new Point(198, 93);
            txtHumanizeLower.Name = "txtHumanizeLower";
            txtHumanizeLower.Size = new Size(66, 35);
            txtHumanizeLower.TabIndex = 15;
            // 
            // lblHumanizeSeparator
            // 
            lblHumanizeSeparator.Location = new Point(270, 93);
            lblHumanizeSeparator.Name = "lblHumanizeSeparator";
            lblHumanizeSeparator.Size = new Size(24, 35);
            lblHumanizeSeparator.TabIndex = 16;
            lblHumanizeSeparator.Text = "～";
            lblHumanizeSeparator.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // txtHumanizeUpper
            // 
            txtHumanizeUpper.Location = new Point(300, 93);
            txtHumanizeUpper.Name = "txtHumanizeUpper";
            txtHumanizeUpper.Size = new Size(66, 35);
            txtHumanizeUpper.TabIndex = 17;
            // 
            // btnToggleHumanize
            // 
            btnToggleHumanize.Location = new Point(257, 137);
            btnToggleHumanize.Name = "btnToggleHumanize";
            btnToggleHumanize.Size = new Size(146, 37);
            btnToggleHumanize.TabIndex = 3;
            btnToggleHumanize.Text = "擬人化切替";
            btnToggleHumanize.UseVisualStyleBackColor = true;
            btnToggleHumanize.Click += btnToggleHumanize_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(12F, 30F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1227, 980);
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
            Controls.Add(lblHumanizeSeparator);
            Controls.Add(txtHumanizeUpper);
            Controls.Add(btnToggleHumanize);
            Controls.Add(btnStop);
            Controls.Add(btnPause);
            Controls.Add(btnStart);
            Name = "Form1";
            captureToolStrip.ResumeLayout(false);
            captureToolStrip.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dgvScenario).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
    }
}
