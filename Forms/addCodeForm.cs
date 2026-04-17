using System;
using System.Collections.Generic;
using System.Drawing;
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
            if (string.IsNullOrWhiteSpace(CodeNameTextBox.Text) &&
                string.IsNullOrWhiteSpace(CodeEditTextBox.Text) &&
                string.IsNullOrWhiteSpace(RemarkTextBox.Text))
            {
                MessageBox.Show("未输入内容，窗口即将关闭");
                DialogResult = DialogResult.Cancel;
                Close();
                return;
            }

            string code = CodeEditTextBox.Text;
            if (!VbaManagerForm.IsValidVbaCode(code))
            {
                MessageBox.Show("输入的代码不符合 VBA 程序规范，请检查代码。");
                DialogResult = DialogResult.None;
                return;
            }

            string codeName = CodeNameTextBox.Text;
            if (string.IsNullOrEmpty(codeName) &&
                !VbaManagerForm.TryExtractMacroName(code, out codeName))
            {
                MessageBox.Show("输入的代码不符合 VBA 程序规范，请检查代码。");
                DialogResult = DialogResult.None;
                return;
            }

            if (string.IsNullOrWhiteSpace(RemarkTextBox.Text))
            {
                RemarkTextBox.Text = codeName;
            }

            ShortcutTextBox.Text = NormalizeShortcutText(ShortcutTextBox.Text);
            DialogResult = DialogResult.OK;
            Close();
        }

        private void CodeEditTextBox_TextChanged(object sender, EventArgs e)
        {
            int selectionStart = CodeEditTextBox.SelectionStart;
            int selectionLength = CodeEditTextBox.SelectionLength;

            string text = CodeEditTextBox.Text;
            string newText = System.Text.RegularExpressions.Regex.Replace(
                text,
                @"\bsub\b",
                "Sub",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            newText = System.Text.RegularExpressions.Regex.Replace(
                newText,
                @"\bend sub\b",
                "End Sub",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            if (text != newText)
            {
                CodeEditTextBox.TextChanged -= CodeEditTextBox_TextChanged;
                CodeEditTextBox.Text = newText;
                CodeEditTextBox.TextChanged += CodeEditTextBox_TextChanged;
                CodeEditTextBox.SelectionStart = selectionStart;
                CodeEditTextBox.SelectionLength = selectionLength;
            }

            CodeEditTextBox.Font = new Font("Consolas", 10);
        }

        private static string NormalizeShortcutText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            return text
                .Trim()
                .Replace("＋", "+")
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
