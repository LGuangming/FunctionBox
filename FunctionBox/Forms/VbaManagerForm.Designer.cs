using System.Windows.Forms;

namespace FunctionBox
{
    partial class VbaManagerForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(VbaManagerForm));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnBoxEdit = new System.Windows.Forms.Button();
            this.btnBoxDelete = new System.Windows.Forms.Button();
            this.btnBoxAdd = new System.Windows.Forms.Button();
            this.SearchTextBox = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnImport = new System.Windows.Forms.Button();
            this.lstVbaCodes = new System.Windows.Forms.ListView();
            this.代码名称 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.备注 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.是否在右键菜单显示 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnBoxEdit);
            this.groupBox1.Controls.Add(this.btnBoxDelete);
            this.groupBox1.Controls.Add(this.btnBoxAdd);
            this.groupBox1.Location = new System.Drawing.Point(483, 35);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Size = new System.Drawing.Size(100, 146);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "功能";
            // 
            // btnBoxEdit
            // 
            this.btnBoxEdit.Location = new System.Drawing.Point(4, 59);
            this.btnBoxEdit.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnBoxEdit.Name = "btnBoxEdit";
            this.btnBoxEdit.Size = new System.Drawing.Size(92, 35);
            this.btnBoxEdit.TabIndex = 2;
            this.btnBoxEdit.Text = "编辑";
            this.btnBoxEdit.UseVisualStyleBackColor = true;
            this.btnBoxEdit.Click += new System.EventHandler(this.btnBoxEdit_Click);
            // 
            // btnBoxDelete
            // 
            this.btnBoxDelete.Location = new System.Drawing.Point(4, 99);
            this.btnBoxDelete.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnBoxDelete.Name = "btnBoxDelete";
            this.btnBoxDelete.Size = new System.Drawing.Size(92, 35);
            this.btnBoxDelete.TabIndex = 1;
            this.btnBoxDelete.Text = "删除";
            this.btnBoxDelete.UseVisualStyleBackColor = true;
            this.btnBoxDelete.Click += new System.EventHandler(this.btnBoxDelete_Click);
            // 
            // btnBoxAdd
            // 
            this.btnBoxAdd.Location = new System.Drawing.Point(4, 19);
            this.btnBoxAdd.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnBoxAdd.Name = "btnBoxAdd";
            this.btnBoxAdd.Size = new System.Drawing.Size(92, 35);
            this.btnBoxAdd.TabIndex = 0;
            this.btnBoxAdd.Text = "新建";
            this.btnBoxAdd.UseVisualStyleBackColor = true;
            this.btnBoxAdd.Click += new System.EventHandler(this.btnBoxAdd_Click);
            // 
            // SearchTextBox
            // 
            this.SearchTextBox.Location = new System.Drawing.Point(44, 11);
            this.SearchTextBox.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.SearchTextBox.Name = "SearchTextBox";
            this.SearchTextBox.Size = new System.Drawing.Size(253, 21);
            this.SearchTextBox.TabIndex = 2;
            this.SearchTextBox.Tag = "";
            this.SearchTextBox.Text = "在此输入名称，进行模糊查找";
            this.SearchTextBox.TextChanged += new System.EventHandler(this.SearchTextBox_TextChanged);
            this.SearchTextBox.Enter += new System.EventHandler(this.SearchTextBox_Enter);
            this.SearchTextBox.Leave += new System.EventHandler(this.SearchTextBox_Leave);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnExport);
            this.groupBox2.Controls.Add(this.btnImport);
            this.groupBox2.Location = new System.Drawing.Point(483, 257);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox2.Size = new System.Drawing.Size(100, 98);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "备份/还原";
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(4, 57);
            this.btnExport.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(92, 35);
            this.btnExport.TabIndex = 4;
            this.btnExport.Text = "导出VBA清单";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(4, 16);
            this.btnImport.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(92, 35);
            this.btnImport.TabIndex = 3;
            this.btnImport.Text = "导入VBA清单";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // lstVbaCodes
            // 
            this.lstVbaCodes.Alignment = System.Windows.Forms.ListViewAlignment.Left;
            this.lstVbaCodes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lstVbaCodes.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.代码名称,
            this.备注,
            this.是否在右键菜单显示});
            this.lstVbaCodes.FullRowSelect = true;
            this.lstVbaCodes.GridLines = true;
            this.lstVbaCodes.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstVbaCodes.HideSelection = false;
            this.lstVbaCodes.Location = new System.Drawing.Point(9, 40);
            this.lstVbaCodes.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.lstVbaCodes.MultiSelect = false;
            this.lstVbaCodes.Name = "lstVbaCodes";
            this.lstVbaCodes.Size = new System.Drawing.Size(470, 314);
            this.lstVbaCodes.TabIndex = 5;
            this.lstVbaCodes.UseCompatibleStateImageBehavior = false;
            this.lstVbaCodes.View = System.Windows.Forms.View.Details;
            this.lstVbaCodes.DoubleClick += new System.EventHandler(this.btnBoxEdit_Click);
            this.lstVbaCodes.ColumnWidthChanging += new System.Windows.Forms.ColumnWidthChangingEventHandler(this.lstVbaCodes_ColumnWidthChanging);
            this.lstVbaCodes.SizeChanged += new System.EventHandler(this.lstVbaCodes_SizeChanged);
            // 
            // 代码名称
            // 
            this.代码名称.Text = "代码名称";
            this.代码名称.Width = 150;
            // 
            // 备注
            // 
            this.备注.Text = "备注";
            this.备注.Width = 325;
            // 
            // 是否在右键菜单显示
            // 
            this.是否在右键菜单显示.Text = "快捷键";
            this.是否在右键菜单显示.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.是否在右键菜单显示.Width = 150;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 14);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 12);
            this.label1.TabIndex = 6;
            this.label1.Text = "搜索:";
            // 
            // VbaManagerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(586, 362);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.SearchTextBox);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.lstVbaCodes);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.MaximizeBox = false;
            this.Name = "VbaManagerForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "VBA工具箱";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.VbaManagerForm_FormClosing);
            this.Load += new System.EventHandler(this.VbaManagerForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnBoxEdit;
        private System.Windows.Forms.Button btnBoxDelete;
        private System.Windows.Forms.Button btnBoxAdd;
        private System.Windows.Forms.TextBox SearchTextBox;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.ListView lstVbaCodes;
        private System.Windows.Forms.Label label1;
        private ColumnHeader 代码名称;
        private ColumnHeader 备注;
        private ColumnHeader 是否在右键菜单显示;
    }
}
