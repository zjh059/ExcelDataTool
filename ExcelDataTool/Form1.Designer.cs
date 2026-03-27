namespace ExcelDataTool
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Button btnSelectFile;

        // 新增的两个控件：标签和下拉菜单
        private System.Windows.Forms.Label lblTaskSelect;
        private System.Windows.Forms.ComboBox cmbTaskSelect;

        private System.Windows.Forms.Button btnProcess;
        private System.Windows.Forms.RichTextBox rtbLogs;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.lblTaskSelect = new System.Windows.Forms.Label();
            this.cmbTaskSelect = new System.Windows.Forms.ComboBox();
            this.btnProcess = new System.Windows.Forms.Button();
            this.rtbLogs = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();

            // txtFilePath (第一行)
            this.txtFilePath.Location = new System.Drawing.Point(12, 12);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.ReadOnly = true;
            this.txtFilePath.Size = new System.Drawing.Size(400, 23);
            this.txtFilePath.TabIndex = 0;

            // btnSelectFile (第一行)
            this.btnSelectFile.Location = new System.Drawing.Point(420, 11);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(100, 25);
            this.btnSelectFile.TabIndex = 1;
            this.btnSelectFile.Text = "选择 Excel";
            this.btnSelectFile.UseVisualStyleBackColor = true;
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);

            // lblTaskSelect (第二行：提示文字)
            this.lblTaskSelect.AutoSize = true;
            this.lblTaskSelect.Location = new System.Drawing.Point(12, 48);
            this.lblTaskSelect.Name = "lblTaskSelect";
            this.lblTaskSelect.Size = new System.Drawing.Size(91, 15);
            this.lblTaskSelect.TabIndex = 4;
            this.lblTaskSelect.Text = "请选择提取月份:";

            // cmbTaskSelect (第二行：下拉菜单)
            this.cmbTaskSelect.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList; // 设置为只能选不能手写
            this.cmbTaskSelect.FormattingEnabled = true;
            this.cmbTaskSelect.Location = new System.Drawing.Point(110, 45);
            this.cmbTaskSelect.Name = "cmbTaskSelect";
            this.cmbTaskSelect.Size = new System.Drawing.Size(302, 23);
            this.cmbTaskSelect.TabIndex = 5;

            // btnProcess (第二行：开始按钮移到这里)
            this.btnProcess.Location = new System.Drawing.Point(420, 44);
            this.btnProcess.Name = "btnProcess";
            this.btnProcess.Size = new System.Drawing.Size(100, 25);
            this.btnProcess.TabIndex = 2;
            this.btnProcess.Text = "开始处理";
            this.btnProcess.UseVisualStyleBackColor = true;
            this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);

            // rtbLogs (第三行：日志框整体往下挪一点)
            this.rtbLogs.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rtbLogs.Location = new System.Drawing.Point(12, 80);
            this.rtbLogs.Name = "rtbLogs";
            this.rtbLogs.ReadOnly = true;
            this.rtbLogs.Size = new System.Drawing.Size(618, 320);
            this.rtbLogs.TabIndex = 3;
            this.rtbLogs.Text = "欢迎使用团队版数据处理工具！\n1. 请点击“选择 Excel”加载总表。\n2. 选择你需要处理的月份区间。\n3. 点击“开始处理”。\n";

            // Form1
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(642, 412);
            this.Controls.Add(this.cmbTaskSelect);
            this.Controls.Add(this.lblTaskSelect);
            this.Controls.Add(this.rtbLogs);
            this.Controls.Add(this.btnProcess);
            this.Controls.Add(this.btnSelectFile);
            this.Controls.Add(this.txtFilePath);
            this.Name = "Form1";
            this.Text = "团队协作版 Excel 分类神器";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}