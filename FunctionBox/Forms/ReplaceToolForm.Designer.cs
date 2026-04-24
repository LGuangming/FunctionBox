namespace FunctionBox.Forms
{
    partial class ReplaceToolForm
    {
        private System.ComponentModel.IContainer components = null;

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
            this.rulesPanel = new System.Windows.Forms.Panel();
            this.btnAddRule = new System.Windows.Forms.Button();
            this.btnDeleteRule = new System.Windows.Forms.Button();
            this.btnClearOld = new System.Windows.Forms.Button();
            this.btnClearNew = new System.Windows.Forms.Button();
            this.btnReplaceCurrent = new System.Windows.Forms.Button();
            this.btnReplaceOther = new System.Windows.Forms.Button();
            this.cmbSpecialChars = new System.Windows.Forms.ComboBox();
            this.btnInsertSpecialChar = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // rulesPanel
            // 
            this.rulesPanel.AutoScroll = true;
            this.rulesPanel.Location = new System.Drawing.Point(16, 15);
            this.rulesPanel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.rulesPanel.Name = "rulesPanel";
            this.rulesPanel.Size = new System.Drawing.Size(442, 192);
            this.rulesPanel.TabIndex = 0;
            // 
            // btnAddRule
            // 
            this.btnAddRule.Location = new System.Drawing.Point(16, 219);
            this.btnAddRule.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnAddRule.Name = "btnAddRule";
            this.btnAddRule.Size = new System.Drawing.Size(133, 38);
            this.btnAddRule.TabIndex = 1;
            this.btnAddRule.Text = "添加替换规则";
            this.btnAddRule.UseVisualStyleBackColor = true;
            this.btnAddRule.Click += new System.EventHandler(this.btnAddRule_Click);
            // 
            // btnDeleteRule
            // 
            this.btnDeleteRule.Location = new System.Drawing.Point(16, 264);
            this.btnDeleteRule.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnDeleteRule.Name = "btnDeleteRule";
            this.btnDeleteRule.Size = new System.Drawing.Size(133, 38);
            this.btnDeleteRule.TabIndex = 3;
            this.btnDeleteRule.Text = "删除选中规则";
            this.btnDeleteRule.UseVisualStyleBackColor = true;
            this.btnDeleteRule.Click += new System.EventHandler(this.btnDeleteRule_Click);
            // 
            // btnClearOld
            // 
            this.btnClearOld.Location = new System.Drawing.Point(157, 219);
            this.btnClearOld.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnClearOld.Name = "btnClearOld";
            this.btnClearOld.Size = new System.Drawing.Size(147, 38);
            this.btnClearOld.TabIndex = 2;
            this.btnClearOld.Text = "清空替换前字符";
            this.btnClearOld.UseVisualStyleBackColor = true;
            this.btnClearOld.Click += new System.EventHandler(this.btnClearOld_Click);
            // 
            // btnClearNew
            // 
            this.btnClearNew.Location = new System.Drawing.Point(157, 264);
            this.btnClearNew.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnClearNew.Name = "btnClearNew";
            this.btnClearNew.Size = new System.Drawing.Size(147, 38);
            this.btnClearNew.TabIndex = 4;
            this.btnClearNew.Text = "清空替换后字符";
            this.btnClearNew.UseVisualStyleBackColor = true;
            this.btnClearNew.Click += new System.EventHandler(this.btnClearNew_Click);
            // 
            // btnReplaceCurrent
            // 
            this.btnReplaceCurrent.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.btnReplaceCurrent.Location = new System.Drawing.Point(248, 320);
            this.btnReplaceCurrent.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnReplaceCurrent.Name = "btnReplaceCurrent";
            this.btnReplaceCurrent.Size = new System.Drawing.Size(210, 50);
            this.btnReplaceCurrent.TabIndex = 5;
            this.btnReplaceCurrent.Text = "替换当前文档";
            this.btnReplaceCurrent.UseVisualStyleBackColor = true;
            this.btnReplaceCurrent.Click += new System.EventHandler(this.btnReplaceCurrent_Click);
            // 
            // btnReplaceOther
            // 
            this.btnReplaceOther.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.btnReplaceOther.Location = new System.Drawing.Point(16, 320);
            this.btnReplaceOther.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnReplaceOther.Name = "btnReplaceOther";
            this.btnReplaceOther.Size = new System.Drawing.Size(210, 50);
            this.btnReplaceOther.TabIndex = 6;
            this.btnReplaceOther.Text = "选择多个Word文件";
            this.btnReplaceOther.UseVisualStyleBackColor = true;
            this.btnReplaceOther.Click += new System.EventHandler(this.btnReplaceOther_Click);
            // 
            // cmbSpecialChars
            // 
            this.cmbSpecialChars.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSpecialChars.Items.AddRange(new object[] {
            "\\n (换行符)",
            "\\t (制表符)",
            "\\r (回车符)",
            "\\p (段落符)"});
            this.cmbSpecialChars.Location = new System.Drawing.Point(313, 225);
            this.cmbSpecialChars.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cmbSpecialChars.Name = "cmbSpecialChars";
            this.cmbSpecialChars.Size = new System.Drawing.Size(145, 23);
            this.cmbSpecialChars.TabIndex = 7;
            // 
            // btnInsertSpecialChar
            // 
            this.btnInsertSpecialChar.Location = new System.Drawing.Point(313, 264);
            this.btnInsertSpecialChar.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnInsertSpecialChar.Name = "btnInsertSpecialChar";
            this.btnInsertSpecialChar.Size = new System.Drawing.Size(145, 38);
            this.btnInsertSpecialChar.TabIndex = 8;
            this.btnInsertSpecialChar.Text = "插入特殊字符";
            this.btnInsertSpecialChar.UseVisualStyleBackColor = true;
            this.btnInsertSpecialChar.Click += new System.EventHandler(this.btnInsertSpecialChar_Click);
            // 
            // ReplaceToolForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(471, 383);
            this.Controls.Add(this.cmbSpecialChars);
            this.Controls.Add(this.btnInsertSpecialChar);
            this.Controls.Add(this.btnReplaceOther);
            this.Controls.Add(this.btnReplaceCurrent);
            this.Controls.Add(this.btnClearNew);
            this.Controls.Add(this.btnDeleteRule);
            this.Controls.Add(this.btnClearOld);
            this.Controls.Add(this.btnAddRule);
            this.Controls.Add(this.rulesPanel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ReplaceToolForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "文本批量替换工具";
            this.ResumeLayout(false);

        }

        private System.Windows.Forms.Panel rulesPanel;
        private System.Windows.Forms.Button btnAddRule;
        private System.Windows.Forms.Button btnDeleteRule;
        private System.Windows.Forms.Button btnClearOld;
        private System.Windows.Forms.Button btnClearNew;
        private System.Windows.Forms.Button btnReplaceCurrent;
        private System.Windows.Forms.Button btnReplaceOther;
        private System.Windows.Forms.ComboBox cmbSpecialChars;
        private System.Windows.Forms.Button btnInsertSpecialChar;
    }
}
