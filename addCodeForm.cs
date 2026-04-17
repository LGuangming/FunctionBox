using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FunctionBox
{
    public partial class addCodeForm : Form
    {
        public addCodeForm()
        {
            InitializeComponent();
        }
        private void AddCodeConfirm_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(CodeNameTextBox.Text) && string.IsNullOrWhiteSpace(CodeEditTextBox.Text) && string.IsNullOrWhiteSpace(RemarkTextBox.Text))
            {
                // 如果三个TextBox都为空，则关闭窗口而不设置DialogResult
                MessageBox.Show("未输入内容，窗口即将关闭");
                this.DialogResult = DialogResult.Cancel; // 设置DialogResult为Cancel
                this.Close();
            }
            else
            {
                string code = this.CodeEditTextBox.Text;

                if (!VbaManagerForm.IsValidVbaCode(code))
                {
                    MessageBox.Show("输入的代码不符合VBA程序规范，请检查代码�?);
                    this.DialogResult = DialogResult.None; // 阻止对话框关�?
                }
                else
                {
                    string codeName = this.CodeNameTextBox.Text;

                    // 如果 codeName 为空，尝试从 code 中提�?
                    if (string.IsNullOrEmpty(codeName))
                    {
                        if (!VbaManagerForm.TryExtractMacroName(code, out codeName))
                        {
                            MessageBox.Show("输入的代码不符合VBA程序规范，请检查代码�?);
                            this.DialogResult = DialogResult.None; // 阻止对话框关�?
                            return;
                        }
                    }

                    // 如果备注为空，则使用代码名称填充备注
                    if (string.IsNullOrWhiteSpace(this.RemarkTextBox.Text))
                    {
                        this.RemarkTextBox.Text = codeName;
                    }

                    this.ShortcutTextBox.Text = NormalizeShortcutText(this.ShortcutTextBox.Text);

                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
            }
        }
        private void CodeEditTextBox_TextChanged(object sender, EventArgs e)
        {
            // 获取当前光标位置
            int selectionStart = this.CodeEditTextBox.SelectionStart;
            int selectionLength = this.CodeEditTextBox.SelectionLength;

            // 替换独立�?"sub" �?"Sub" �?"end sub" �?"End Sub"
            string text = this.CodeEditTextBox.Text;
            string newText = System.Text.RegularExpressions.Regex.Replace(text, @"\bsub\b", "Sub", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            newText = System.Text.RegularExpressions.Regex.Replace(newText, @"\bend sub\b", "End Sub", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            if (text != newText)
            {
                this.CodeEditTextBox.TextChanged -= CodeEditTextBox_TextChanged; // 暂时移除事件处理程序
                this.CodeEditTextBox.Text = newText;
                this.CodeEditTextBox.TextChanged += CodeEditTextBox_TextChanged; // 恢复事件处理程序

                // 恢复光标位置
                this.CodeEditTextBox.SelectionStart = selectionStart;
                this.CodeEditTextBox.SelectionLength = selectionLength;
            }

            // 确保字体和格式保持一�?            this.CodeEditTextBox.Font = new Font("Consolas", 10);
        }
        private static string NormalizeShortcutText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            return text
                .Trim()
                .Replace("�?, "+")
                .Replace(" ", string.Empty)
                .ToUpperInvariant();
        }
        private void ShortcutTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            e.SuppressKeyPress = true;

            if (e.KeyCode == Keys.Back || e.KeyCode == Keys.Delete)
            {
                ShortcutTextBox.Text = string.Empty;
                return;
            }

            if (e.KeyCode == Keys.ControlKey || e.KeyCode == Keys.ShiftKey || e.KeyCode == Keys.Menu)
            {
                return;
            }

            List<string> parts = new List<string>();
            if (e.Control)
            {
                parts.Add("Ctrl");
            }
            if (e.Shift)
            {
                parts.Add("Shift");
            }
            if (e.Alt)
            {
                parts.Add("Alt");
            }

            parts.Add(GetKeyDisplayName(e.KeyCode));
            ShortcutTextBox.Text = string.Join("+", parts);
        }
        private static string GetKeyDisplayName(Keys key)
        {
            switch (key)
            {
                case Keys.Oemtilde:
                    return "`";
                case Keys.OemQuotes:
                    return "'";
                case Keys.OemMinus:
                    return "-";
                case Keys.Oemplus:
                    return "=";
                case Keys.OemOpenBrackets:
                    return "[";
                case Keys.OemCloseBrackets:
                    return "]";
                case Keys.OemPipe:
                case Keys.OemBackslash:
                    return "\\";
                case Keys.OemSemicolon:
                    return ";";
                case Keys.Oemcomma:
                    return ",";
                case Keys.OemPeriod:
                    return ".";
                case Keys.OemQuestion:
                    return "/";
            }

            if (key >= Keys.D0 && key <= Keys.D9)
            {
                return ((char)('0' + (key - Keys.D0))).ToString();
            }

            if (key >= Keys.NumPad0 && key <= Keys.NumPad9)
            {
                return "Num" + (key - Keys.NumPad0);
            }

            return key.ToString().ToUpperInvariant();
        }
    }
}

