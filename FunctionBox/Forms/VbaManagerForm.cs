using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;


namespace FunctionBox
{
    public partial class VbaManagerForm : Form
    {
        private RibbonDropDown btnToolList; // 假设 btnToolList 是一个 RibbonDropDown 控件
        private RibbonButton btnExecuteVba;
        private const string EmptyMacroPlaceholder = "添加VBA代码";
        private string SaveFilePath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "FunctionBox",
            "vba_codes.json");
        private List<ListViewItem> allItems = new List<ListViewItem>();
        public VbaManagerForm(RibbonDropDown toolList, RibbonButton executeButton = null)
        {
            btnToolList = toolList ?? throw new ArgumentNullException(nameof(toolList));
            btnExecuteVba = executeButton;
            InitializeComponent();
            LoadVbaCodes(null, btnToolList);
        }
        private void btnBoxAdd_Click(object sender, EventArgs e)
        {
            using (addCodeForm addCodeForm = new addCodeForm())
            {
                if (addCodeForm.ShowDialog() == DialogResult.OK)
                {
                    string codeName = addCodeForm.CodeNameTextBox.Text;
                    string code = addCodeForm.CodeEditTextBox.Text;
                    string remark = addCodeForm.RemarkTextBox.Text;
                    string shortcut = addCodeForm.ShortcutTextBox.Text;
                    if (shortcut != null) shortcut = shortcut.Trim();

                    if (string.IsNullOrEmpty(codeName))
                    {
                        if (!TryExtractMacroName(code, out codeName))
                        {
                            MessageBox.Show("输入的代码不符合VBA程序规范，请检查代码。");
                            return;
                        }
                    }

                    // 如果备注为空，则使用代码名称填充备注
                    if (string.IsNullOrWhiteSpace(remark))
                    {
                        remark = codeName;
                    }

                    VbaCode newCode = new VbaCode { Name = codeName, Code = code, Remark = remark, Shortcut = shortcut };
                    ListViewItem item = new ListViewItem(new[] { codeName, remark, shortcut });
                    item.Tag = newCode;
                    lstVbaCodes.Items.Add(item);
                    allItems.Add(item);
                    SaveVbaCodes();
                    SyncToolListFromAllItems();
                }
            }
        }
        private void btnBoxDelete_Click(object sender, EventArgs e)
        {
            // 删除选中的VBA代码
            if (lstVbaCodes.SelectedItems.Count > 0)
            {
                ListViewItem selectedItem = lstVbaCodes.SelectedItems[0];
                VbaCode selectedCode = (VbaCode)selectedItem.Tag;
                lstVbaCodes.Items.Remove(selectedItem);
                allItems.Remove(selectedItem);
                allItems.RemoveAll(item =>
                {
                    VbaCode code = item.Tag as VbaCode;
                    if (code == null || selectedCode == null)
                    {
                        return false;
                    }

                    return string.Equals(code.Name, selectedCode.Name, StringComparison.Ordinal) &&
                        string.Equals(code.Code, selectedCode.Code, StringComparison.Ordinal);
                });

                SaveVbaCodes();
                SyncToolListFromAllItems();
            }
        }
        private void btnBoxEdit_Click(object sender, EventArgs e)
        {
            // 编辑选中的VBA代码
            if (lstVbaCodes.SelectedItems.Count > 0)
            {
                ListViewItem selectedItem = lstVbaCodes.SelectedItems[0];
                VbaCode selectedCode = (VbaCode)selectedItem.Tag;

                using (addCodeForm addCodeForm = new addCodeForm())
                {
                    addCodeForm.CodeNameTextBox.Text = selectedCode.Name;
                    addCodeForm.CodeEditTextBox.Text = selectedCode.Code;
                    addCodeForm.RemarkTextBox.Text = selectedCode.Remark;
                    addCodeForm.ShortcutTextBox.Text = selectedCode.Shortcut;

                    if (addCodeForm.ShowDialog() == DialogResult.OK)
                    {
                        string codeName = addCodeForm.CodeNameTextBox.Text;
                        string code = addCodeForm.CodeEditTextBox.Text;
                        string remark = addCodeForm.RemarkTextBox.Text;
                        string shortcut = addCodeForm.ShortcutTextBox.Text;
                        if (shortcut != null) shortcut = shortcut.Trim();

                        if (string.IsNullOrEmpty(codeName))
                        {
                            if (!TryExtractMacroName(code, out codeName))
                            {
                                MessageBox.Show("输入的代码不符合VBA程序规范，请检查代码。");
                                return;
                            }
                        }

                        selectedCode.Name = codeName;
                        selectedCode.Code = code;
                        selectedCode.Remark = remark;
                        selectedCode.Shortcut = shortcut;

                        selectedItem.SubItems[0].Text = codeName;
                        selectedItem.SubItems[1].Text = remark;
                        selectedItem.SubItems[2].Text = shortcut;
                        SaveVbaCodes();
                        SyncToolListFromAllItems();
                    }
                }
            }
        }
        public static bool IsValidVbaCode(string code)
        {
            string dummy;
            if (!TryExtractMacroName(code, out dummy))
            {
                return false;
            }

            var lines = code.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            string lastNonCommentLine = lines
                .Reverse()
                .FirstOrDefault(line =>
                {
                    string trimmed = line.Trim();
                    return !string.IsNullOrWhiteSpace(trimmed) && !trimmed.StartsWith("'");
                });

            return string.Equals(lastNonCommentLine, "End Sub", StringComparison.OrdinalIgnoreCase);
        }
        public static bool TryExtractMacroName(string code, out string macroName)
        {
            macroName = string.Empty;
            if (string.IsNullOrWhiteSpace(code))
            {
                return false;
            }

            var lines = code.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            string firstNonCommentLine = lines.FirstOrDefault(line =>
            {
                string trimmed = line.Trim();
                return !string.IsNullOrWhiteSpace(trimmed) && !trimmed.StartsWith("'");
            });

            if (string.IsNullOrWhiteSpace(firstNonCommentLine))
            {
                return false;
            }

            Match match = Regex.Match(
                firstNonCommentLine.Trim(),
                @"^(?:(?:Public|Private|Friend)\s+)?Sub\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\(|$)",
                RegexOptions.IgnoreCase);

            if (!match.Success)
            {
                return false;
            }

            macroName = match.Groups[1].Value;
            return !string.IsNullOrWhiteSpace(macroName);
        }
        private void btnSaveResult_Click(object sender, EventArgs e)
        {
            SaveVbaCodes();
            //MessageBox.Show("保存成功！");
        }
        private void btnExport_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "JSON Files|*.json";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                SaveVbaCodes(saveFileDialog.FileName);
                MessageBox.Show("导出成功！");
            }
        }
        private void btnImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "JSON Files|*.json";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                LoadVbaCodes(openFileDialog.FileName, btnToolList);
                MessageBox.Show("导入成功！");
            }
        }
        public void SaveVbaCodes(string filePath = null)
        {
            if (filePath == null)
            {
                filePath = SaveFilePath;
            }

            List<VbaCode> vbaCodes = new List<VbaCode>();

            foreach (ListViewItem item in allItems)
            {
                VbaCode code = (VbaCode)item.Tag;
                vbaCodes.Add(code);
            }

            string directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            string json = JsonConvert.SerializeObject(vbaCodes, Formatting.Indented);
            File.WriteAllText(filePath, json);
        }
        public void LoadVbaCodes(string filePath = null, RibbonDropDown toolList = null)
        {
            if (filePath == null)
            {
                filePath = SaveFilePath;
            }

            lstVbaCodes.Items.Clear();
            allItems.Clear();
            if (toolList != null)
            {
                toolList.Items.Clear();
            }

            if (File.Exists(filePath))
            {
                string json = File.ReadAllText(filePath);
                List<VbaCode> vbaCodes = JsonConvert.DeserializeObject<List<VbaCode>>(json);

                foreach (var code in vbaCodes)
                {
                    ListViewItem item = new ListViewItem(new[] { code.Name, code.Remark, code.Shortcut ?? string.Empty });
                    item.Tag = code;
                    lstVbaCodes.Items.Add(item);
                    allItems.Add(item);
                }

            }

            if (toolList != null)
            {
                SyncToolListFromAllItems();
            }
        }
        private void SyncToolListFromAllItems()
        {
            if (btnToolList == null)
            {
                return;
            }

            btnToolList.Items.Clear();
            foreach (ListViewItem item in allItems)
            {
                VbaCode code = item.Tag as VbaCode;
                if (code == null)
                {
                    continue;
                }

                RibbonDropDownItem ribbonItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                ribbonItem.Label = code.Name;
                ribbonItem.Tag = code;
                btnToolList.Items.Add(ribbonItem);
            }

            if (btnToolList.Items.Count > 0)
            {
                btnToolList.Enabled = true;
                btnToolList.SelectedItem = btnToolList.Items[0];
                if (btnExecuteVba != null)
                {
                    btnExecuteVba.Enabled = true;
                }
            }
            else
            {
                RibbonDropDownItem placeholder = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                placeholder.Label = EmptyMacroPlaceholder;
                placeholder.Tag = null;
                btnToolList.Items.Add(placeholder);
                btnToolList.SelectedItem = placeholder;
                btnToolList.Enabled = false;
                if (btnExecuteVba != null)
                {
                    btnExecuteVba.Enabled = false;
                }
            }

            Globals.ThisAddIn.SyncVbaShortcutBindings();
        }
        private void VbaManagerForm_Load(object sender, EventArgs e)
        {
            LoadVbaCodes(null, btnToolList);
            lstVbaCodes_SizeChanged(null, null); // 初始化列宽比例
        }
        private void VbaManagerForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveVbaCodes();
        }
        public class VbaCode
        {
            public string Name { get; set; }
            public string Code { get; set; }
            public string Remark { get; set; }
            public string Shortcut { get; set; }
        }
        public class VbaCodeDialog : Form
        {
            private TextBox txtCodeName;
            private TextBox txtCode;
            private Button btnOk;
            private Button btnCancel;

            public string CodeName { get; private set; }
            public string Code { get; private set; }

            public VbaCodeDialog(string title, string codeName = "", string code = "")
            {
                this.Text = title;
                this.Width = 400;
                this.Height = 300;

                Label lblCodeName = new Label() { Left = 10, Top = 20, Text = "代码名称" };
                txtCodeName = new TextBox() { Left = 100, Top = 20, Width = 260, Text = codeName };

                Label lblCode = new Label() { Left = 10, Top = 60, Text = "代码" };
                txtCode = new TextBox() { Left = 100, Top = 60, Width = 260, Height = 150, Multiline = true, Text = code };

                btnOk = new Button() { Text = "确定", Left = 200, Width = 70, Top = 220, DialogResult = DialogResult.OK };
                btnOk.Click += new EventHandler(this.btnOk_Click);

                btnCancel = new Button() { Text = "取消", Left = 290, Width = 70, Top = 220, DialogResult = DialogResult.Cancel };

                this.Controls.Add(lblCodeName);
                this.Controls.Add(txtCodeName);
                this.Controls.Add(lblCode);
                this.Controls.Add(txtCode);
                this.Controls.Add(btnOk);
                this.Controls.Add(btnCancel);

                this.AcceptButton = btnOk;
                this.CancelButton = btnCancel;
            }

            private void btnOk_Click(object sender, EventArgs e)
            {
                CodeName = txtCodeName.Text;
                Code = txtCode.Text;

                if (string.IsNullOrEmpty(CodeName))
                {
                    // 提取VBA代码的Sub名称作为代码名称
                    int startIndex = Code.IndexOf("Sub ") + 4;
                    int endIndex = Code.IndexOf("(", startIndex);
                    if (startIndex > 3 && endIndex > startIndex)
                    {
                        CodeName = Code.Substring(startIndex, endIndex - startIndex).Trim();
                    }
                }

                this.DialogResult = DialogResult.OK;
                this.Close();
            }

            //  public static string ShowDialog(string title, string codeName = "", string code = "")
            //  {
            //      VbaCodeDialog dialog = new VbaCodeDialog(title, codeName, code);
            //      return dialog.ShowDialog() == DialogResult.OK ? $"{dialog.CodeName}: {dialog.Code}" : null;
            //  }
        }
        private void SearchTextBox_TextChanged(object sender, EventArgs e)
        {
            string keyword = this.SearchTextBox.Text.ToLower();

            lstVbaCodes.BeginUpdate();
            lstVbaCodes.Items.Clear();

            foreach (ListViewItem item in allItems)
            {
                VbaCode code = (VbaCode)item.Tag;
                if (code.Name.ToLower().Contains(keyword) || code.Remark.ToLower().Contains(keyword))
                {
                    lstVbaCodes.Items.Add(item);
                }
            }

            lstVbaCodes.EndUpdate();
        }
        private void SearchTextBox_Enter(object sender, EventArgs e)
        {
            if (SearchTextBox.Text == "在此输入名称，进行模糊查找")
            {
                SearchTextBox.Text = "";
                SearchTextBox.ForeColor = SystemColors.WindowText;
            }
        }
        private void SearchTextBox_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(SearchTextBox.Text))
            {
                SearchTextBox.Text = "在此输入名称，进行模糊查找";
                SearchTextBox.ForeColor = SystemColors.GrayText;
                ResetListView();
            }
        }
        private void ResetListView()
        {
            lstVbaCodes.BeginUpdate();
            lstVbaCodes.Items.Clear();
            foreach (ListViewItem item in allItems)
            {
                lstVbaCodes.Items.Add(item);
            }
            lstVbaCodes.EndUpdate();
        }

        private bool isAdjustingColumn = false;

        private void lstVbaCodes_SizeChanged(object sender, EventArgs e)
        {
            if (isAdjustingColumn || lstVbaCodes.Columns.Count < 3) return;

            isAdjustingColumn = true;
            try
            {
                int totalWidth = lstVbaCodes.ClientSize.Width;
                int col3Width = lstVbaCodes.Columns[2].Width;
                
                int remainingWidth = totalWidth - col3Width;
                if (remainingWidth > 0)
                {
                    // 保持第一列不变，让第二列吸收所有窗口缩放的变化
                    int newCol2Width = remainingWidth - lstVbaCodes.Columns[0].Width;
                    if (newCol2Width < 50)
                    {
                        // 如果第二列太小，则压缩第一列
                        lstVbaCodes.Columns[1].Width = 50;
                        lstVbaCodes.Columns[0].Width = remainingWidth - 50;
                    }
                    else
                    {
                        lstVbaCodes.Columns[1].Width = newCol2Width;
                    }
                }
            }
            finally
            {
                isAdjustingColumn = false;
            }
        }

        private void lstVbaCodes_ColumnWidthChanging(object sender, ColumnWidthChangingEventArgs e)
        {
            if (isAdjustingColumn) return;

            // 锁定第三列的最右侧边界，禁止直接往外或往里拉出空白
            if (e.ColumnIndex == 2)
            {
                e.Cancel = true;
                e.NewWidth = lstVbaCodes.Columns[2].Width;
                return;
            }

            // 允许调整第二列（左侧边界），并让第三列反向补偿宽度
            if (e.ColumnIndex == 1)
            {
                int totalWidth = lstVbaCodes.ClientSize.Width;
                int col1Width = lstVbaCodes.Columns[0].Width;
                int remainingWidth = totalWidth - col1Width;

                int newCol2Width = e.NewWidth;

                // 限制第二列不能无限变大（留给第三列最小 50 的空间）
                if (newCol2Width > remainingWidth - 50)
                {
                    newCol2Width = remainingWidth - 50;
                    e.Cancel = true;
                    e.NewWidth = newCol2Width;
                }

                // 限制第二列不能太小
                if (newCol2Width < 50)
                {
                    newCol2Width = 50;
                    e.Cancel = true;
                    e.NewWidth = newCol2Width;
                }

                // 同步反向调整第三列
                isAdjustingColumn = true;
                lstVbaCodes.Columns[2].Width = remainingWidth - newCol2Width;
                isAdjustingColumn = false;
                return;
            }

            // 允许调整第一列，并让第二列反向补偿宽度
            if (e.ColumnIndex == 0)
            {
                int totalWidth = lstVbaCodes.ClientSize.Width;
                int col3Width = lstVbaCodes.Columns[2].Width;
                int remainingWidth = totalWidth - col3Width;

                int newCol1Width = e.NewWidth;

                // 限制第一列不能无限变大（留给第二列最小 50 的空间）
                if (newCol1Width > remainingWidth - 50)
                {
                    newCol1Width = remainingWidth - 50;
                    e.Cancel = true;
                    e.NewWidth = newCol1Width;
                }

                // 限制第一列不能太小
                if (newCol1Width < 50)
                {
                    newCol1Width = 50;
                    e.Cancel = true;
                    e.NewWidth = newCol1Width;
                }

                // 同步反向调整第二列
                isAdjustingColumn = true;
                lstVbaCodes.Columns[1].Width = remainingWidth - newCol1Width;
                isAdjustingColumn = false;
            }
        }

    }
}








