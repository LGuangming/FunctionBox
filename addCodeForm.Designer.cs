using System;
using System.Windows.Forms;

namespace FunctionBox
{
    partial class addCodeForm
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
            System.Windows.Forms.PictureBox HelpBoxTab;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(addCodeForm));
            this.CodeNameTextBox = new System.Windows.Forms.TextBox();
            this.AddCodeConfirm = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabCodePage = new System.Windows.Forms.TabPage();
            this.CodeEditTextBox = new System.Windows.Forms.TextBox();
            this.tabHelpPage = new System.Windows.Forms.TabPage();
            this.label3 = new System.Windows.Forms.Label();
            this.RemarkTextBox = new System.Windows.Forms.TextBox();
            this.HelpBoxLabel = new System.Windows.Forms.Label();
            this.ShortcutLabel = new System.Windows.Forms.Label();
            this.ShortcutTextBox = new System.Windows.Forms.TextBox();
            this.ShortcutHintLabel = new System.Windows.Forms.Label();
            HelpBoxTab = new System.Windows.Forms.PictureBox();
            this.tabControl1.SuspendLayout();
            this.tabCodePage.SuspendLayout();
            this.tabHelpPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(HelpBoxTab)).BeginInit();
            this.SuspendLayout();
            // 
            // CodeNameTextBox
            // 
            this.CodeNameTextBox.Location = new System.Drawing.Point(86, 24);
            this.CodeNameTextBox.Name = "CodeNameTextBox";
            this.CodeNameTextBox.Size = new System.Drawing.Size(208, 25);
            this.CodeNameTextBox.TabIndex = 0;
            // 
            // AddCodeConfirm
            // 
            this.AddCodeConfirm.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.AddCodeConfirm.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.AddCodeConfirm.Font = new System.Drawing.Font("宋体", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.AddCodeConfirm.Location = new System.Drawing.Point(506, 514);
            this.AddCodeConfirm.Name = "AddCodeConfirm";
            this.AddCodeConfirm.Size = new System.Drawing.Size(107, 36);
            this.AddCodeConfirm.TabIndex = 5;
            this.AddCodeConfirm.Text = "确定";
            this.AddCodeConfirm.UseVisualStyleBackColor = true;
            this.AddCodeConfirm.Click += new System.EventHandler(this.AddCodeConfirm_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 29);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 15);
            this.label1.TabIndex = 3;
            this.label1.Text = "代码名称:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(38, 63);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(0, 15);
            this.label2.TabIndex = 4;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabCodePage);
            this.tabControl1.Controls.Add(this.tabHelpPage);
            this.tabControl1.Location = new System.Drawing.Point(15, 96);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(600, 410);
            this.tabControl1.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight;
            this.tabControl1.TabIndex = 6;
            // 
            // tabCodePage
            // 
            this.tabCodePage.Controls.Add(this.CodeEditTextBox);
            this.tabCodePage.Location = new System.Drawing.Point(4, 25);
            this.tabCodePage.Name = "tabCodePage";
            this.tabCodePage.Padding = new System.Windows.Forms.Padding(3, 3, 3, 0);
            this.tabCodePage.Size = new System.Drawing.Size(592, 411);
            this.tabCodePage.TabIndex = 0;
            this.tabCodePage.Text = "代码";
            // 
            // CodeEditTextBox
            // 
            this.CodeEditTextBox.AcceptsReturn = true;
            this.CodeEditTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.CodeEditTextBox.Font = new System.Drawing.Font("Consolas", 9F);
            this.CodeEditTextBox.Location = new System.Drawing.Point(3, 3);
            this.CodeEditTextBox.MaxLength = 3276777;
            this.CodeEditTextBox.Multiline = true;
            this.CodeEditTextBox.Name = "CodeEditTextBox";
            this.CodeEditTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.CodeEditTextBox.Size = new System.Drawing.Size(586, 408);
            this.CodeEditTextBox.TabIndex = 0;
            this.CodeEditTextBox.WordWrap = false;
            this.CodeEditTextBox.TextChanged += new System.EventHandler(this.CodeEditTextBox_TextChanged);
            // 
            // tabHelpPage
            // 
            this.tabHelpPage.Controls.Add(this.HelpBoxLabel);
            this.tabHelpPage.Controls.Add(HelpBoxTab);
            this.tabHelpPage.Location = new System.Drawing.Point(4, 25);
            this.tabHelpPage.Name = "tabHelpPage";
            this.tabHelpPage.Padding = new System.Windows.Forms.Padding(3);
            this.tabHelpPage.Size = new System.Drawing.Size(592, 411);
            this.tabHelpPage.TabIndex = 1;
            this.tabHelpPage.Text = "帮助";
            this.tabHelpPage.ToolTipText = "帮不了一点";
            this.tabHelpPage.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(300, 31);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(45, 15);
            this.label3.TabIndex = 8;
            this.label3.Text = "备注:";
            // 
            // RemarkTextBox
            // 
            this.RemarkTextBox.Location = new System.Drawing.Point(346, 26);
            this.RemarkTextBox.Name = "RemarkTextBox";
            this.RemarkTextBox.Size = new System.Drawing.Size(223, 25);
            this.RemarkTextBox.TabIndex = 1;
            // 
            // ShortcutLabel
            // 
            this.ShortcutLabel.AutoSize = true;
            this.ShortcutLabel.Location = new System.Drawing.Point(12, 66);
            this.ShortcutLabel.Name = "ShortcutLabel";
            this.ShortcutLabel.Size = new System.Drawing.Size(67, 15);
            this.ShortcutLabel.TabIndex = 9;
            this.ShortcutLabel.Text = "快捷键：";
            // 
            // ShortcutTextBox
            // 
            this.ShortcutTextBox.Location = new System.Drawing.Point(86, 60);
            this.ShortcutTextBox.Name = "ShortcutTextBox";
            this.ShortcutTextBox.Size = new System.Drawing.Size(208, 25);
            this.ShortcutTextBox.TabIndex = 2;
            this.ShortcutTextBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ShortcutTextBox_KeyDown);
            // 
            // ShortcutHintLabel
            // 
            this.ShortcutHintLabel.AutoSize = true;
            this.ShortcutHintLabel.ForeColor = System.Drawing.SystemColors.GrayText;
            this.ShortcutHintLabel.Location = new System.Drawing.Point(303, 66);
            this.ShortcutHintLabel.Name = "ShortcutHintLabel";
            this.ShortcutHintLabel.Size = new System.Drawing.Size(187, 15);
            this.ShortcutHintLabel.TabIndex = 11;
            this.ShortcutHintLabel.Text = "按下快捷键设置";
            // 
            // HelpBoxTab
            // 
            HelpBoxTab.Image = global::FunctionBox.Properties.Resources.帮助_问号;
            HelpBoxTab.InitialImage = global::FunctionBox.Properties.Resources.帮助_问号;
            HelpBoxTab.Location = new System.Drawing.Point(214, 115);
            HelpBoxTab.Name = "HelpBoxTab";
            HelpBoxTab.Size = new System.Drawing.Size(112, 105);
            HelpBoxTab.TabIndex = 0;
            HelpBoxTab.TabStop = false;
            // 
            // HelpBoxLabel
            // 
            this.HelpBoxLabel.AutoSize = true;
            this.HelpBoxLabel.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.HelpBoxLabel.Location = new System.Drawing.Point(214, 232);
            this.HelpBoxLabel.Name = "HelpBoxLabel";
            this.HelpBoxLabel.Size = new System.Drawing.Size(109, 20);
            this.HelpBoxLabel.TabIndex = 1;
            this.HelpBoxLabel.Text = "帮不了一点";
            // 
            // addCodeForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(120F, 120F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(632, 560);
            this.Controls.Add(this.ShortcutHintLabel);
            this.Controls.Add(this.ShortcutTextBox);
            this.Controls.Add(this.ShortcutLabel);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.RemarkTextBox);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.AddCodeConfirm);
            this.Controls.Add(this.CodeNameTextBox);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "addCodeForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "添加代码";
            this.tabControl1.ResumeLayout(false);
            this.tabCodePage.ResumeLayout(false);
            this.tabCodePage.PerformLayout();
            this.tabHelpPage.ResumeLayout(false);
            this.tabHelpPage.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(HelpBoxTab)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }



        #endregion
        private System.Windows.Forms.Button AddCodeConfirm;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabCodePage;
        private System.Windows.Forms.TabPage tabHelpPage;
        internal System.Windows.Forms.TextBox CodeNameTextBox;
        internal System.Windows.Forms.TextBox CodeEditTextBox;
        private Label label3;
        internal TextBox RemarkTextBox;
        private Label HelpBoxLabel;
        private Label ShortcutLabel;
        internal TextBox ShortcutTextBox;
        private Label ShortcutHintLabel;
    }
}
