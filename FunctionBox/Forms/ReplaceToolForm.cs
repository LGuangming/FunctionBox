using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace FunctionBox.Forms
{
    public partial class ReplaceToolForm : Form
    {
        private List<RuleEntry> ruleEntries = new List<RuleEntry>();
        private TextBox lastFocusedTextBox = null;

        public ReplaceToolForm()
        {
            InitializeComponent();
            AddRule();
        }

        private class RuleEntry
        {
            public CheckBox EnabledCheckBox { get; set; }
            public TextBox OldTextBox { get; set; }
            public TextBox NewTextBox { get; set; }
            public Panel ContainerPanel { get; set; }
        }

        private void btnAddRule_Click(object sender, EventArgs e)
        {
            AddRule();
        }

        private void AddRule()
        {
            Panel pnl = new Panel() { Height = 35, Dock = DockStyle.Top };
            CheckBox chk = new CheckBox() { Checked = true, Text = "", Width = 20, Location = new System.Drawing.Point(5, 7) };
            TextBox txtOld = new TextBox() { Width = 170, Location = new System.Drawing.Point(30, 5) };
            Label lbl = new Label() { Text = "替换为", AutoSize = true, Location = new System.Drawing.Point(210, 8) };
            TextBox txtNew = new TextBox() { Width = 170, Location = new System.Drawing.Point(265, 5) };

            txtOld.Enter += (s, ev) => lastFocusedTextBox = txtOld;
            txtNew.Enter += (s, ev) => lastFocusedTextBox = txtNew;

            pnl.Controls.Add(chk);
            pnl.Controls.Add(txtOld);
            pnl.Controls.Add(lbl);
            pnl.Controls.Add(txtNew);

            rulesPanel.Controls.Add(pnl);
            pnl.BringToFront();

            ruleEntries.Add(new RuleEntry { EnabledCheckBox = chk, OldTextBox = txtOld, NewTextBox = txtNew, ContainerPanel = pnl });
        }

        private void btnDeleteRule_Click(object sender, EventArgs e)
        {
            bool allSelected = true;
            foreach (var rule in ruleEntries)
            {
                if (!rule.EnabledCheckBox.Checked) allSelected = false;
            }

            if (allSelected)
            {
                for (int i = 1; i < ruleEntries.Count; i++)
                {
                    rulesPanel.Controls.Remove(ruleEntries[i].ContainerPanel);
                    ruleEntries[i].ContainerPanel.Dispose();
                }
                ruleEntries.RemoveRange(1, ruleEntries.Count - 1);
            }
            else
            {
                for (int i = ruleEntries.Count - 1; i >= 0; i--)
                {
                    if (ruleEntries[i].EnabledCheckBox.Checked)
                    {
                        rulesPanel.Controls.Remove(ruleEntries[i].ContainerPanel);
                        ruleEntries[i].ContainerPanel.Dispose();
                        ruleEntries.RemoveAt(i);
                    }
                }
            }
        }

        private void btnClearOld_Click(object sender, EventArgs e)
        {
            foreach (var rule in ruleEntries) rule.OldTextBox.Text = "";
        }

        private void btnClearNew_Click(object sender, EventArgs e)
        {
            foreach (var rule in ruleEntries) rule.NewTextBox.Text = "";
        }

        private void btnInsertSpecialChar_Click(object sender, EventArgs e)
        {
            // 获取选定的特殊字符
            string specialChar = cmbSpecialChars.Text.Split(' ')[0]; // 获取如 "\n"
            
            if (lastFocusedTextBox != null && !lastFocusedTextBox.IsDisposed)
            {
                int cursorPosition = lastFocusedTextBox.SelectionStart;
                lastFocusedTextBox.Text = lastFocusedTextBox.Text.Insert(cursorPosition, specialChar);
                lastFocusedTextBox.SelectionStart = cursorPosition + specialChar.Length;
                lastFocusedTextBox.Focus();
            }
            else
            {
                MessageBox.Show("请先点击选中一个输入框，然后再插入特殊字符。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private List<Tuple<string, string>> GetRules()
        {
            List<Tuple<string, string>> rules = new List<Tuple<string, string>>();
            foreach (var rule in ruleEntries)
            {
                if (rule.EnabledCheckBox.Checked)
                {
                    // Word 中 ^p 代表段落标记(回车)，^t 代表制表符，^l 代表手动换行符(Shift+Enter)
                    string oldText = rule.OldTextBox.Text.Replace("\\n", "^l").Replace("\\t", "^t").Replace("\\r", "^p").Replace("\\p", "^p");
                    string newText = rule.NewTextBox.Text.Replace("\\n", "^l").Replace("\\t", "^t").Replace("\\r", "^p").Replace("\\p", "^p");
                    if (!string.IsNullOrEmpty(oldText))
                    {
                        rules.Add(new Tuple<string, string>(oldText, newText));
                    }
                }
            }
            return rules;
        }

        private void btnReplaceCurrent_Click(object sender, EventArgs e)
        {
            var rules = GetRules();
            if (rules.Count == 0)
            {
                MessageBox.Show("请输入至少一组替换规则！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                ReplaceInDocument(Globals.ThisAddIn.Application.ActiveDocument, rules);
                MessageBox.Show("当前文档替换完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("替换出错: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = true;
            }
        }

        private void btnReplaceOther_Click(object sender, EventArgs e)
        {
            var rules = GetRules();
            if (rules.Count == 0)
            {
                MessageBox.Show("请输入至少一组替换规则！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Word Documents (*.docx;*.doc)|*.docx;*.doc";
                ofd.Multiselect = true;
                ofd.Title = "选择要替换的Word文档";
                
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string[] files = ofd.FileNames;
                    // 在后台任务中执行以避免卡住 UI
                    Task.Run(() => ProcessMultipleFiles(files, rules));
                }
            }
        }

        private void ProcessMultipleFiles(string[] files, List<Tuple<string, string>> rules)
        {
            int success = 0;
            int fail = 0;
            var app = Globals.ThisAddIn.Application;

            foreach (string file in files)
            {
                try
                {
                    Word.Document doc = app.Documents.Open(file, Visible: false);
                    ReplaceInDocument(doc, rules);
                    doc.Save();
                    doc.Close();
                    success++;
                }
                catch
                {
                    fail++;
                }
            }

            this.Invoke((MethodInvoker)delegate
            {
                MessageBox.Show(string.Format("替换完成！\n成功: {0}\n失败: {1}", success, fail), "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            });
        }

        private void ReplaceInDocument(Word.Document doc, List<Tuple<string, string>> rules)
        {
            foreach (Word.Range storyRange in doc.StoryRanges)
            {
                Word.Range currentRange = storyRange;
                while (currentRange != null)
                {
                    foreach (var rule in rules)
                    {
                        Word.Find find = currentRange.Find;
                        find.ClearFormatting();
                        find.Replacement.ClearFormatting();
                        find.Text = rule.Item1;
                        find.Replacement.Text = rule.Item2;
                        find.Format = false;
                        find.MatchCase = false;
                        find.MatchWholeWord = false;
                        find.MatchWildcards = false;
                        find.MatchSoundsLike = false;
                        find.MatchAllWordForms = false;
                        find.Wrap = Word.WdFindWrap.wdFindContinue;
                        find.Execute(Replace: Word.WdReplace.wdReplaceAll);
                    }
                    currentRange = currentRange.NextStoryRange;
                }
            }
        }
    }
}
