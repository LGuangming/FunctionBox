using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace FunctionBox.Features
{
    public static class DecimalAlignTool
    {
        public static void Execute(Word.Application app)
        {
            if (app.Documents.Count == 0)
            {
                MessageBox.Show("当前没有打开的文档。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DialogResult userChoice = MessageBox.Show(
                "横杠\"-\"在文档中代表零，请选择对齐方式：\n" +
                "是(Y)：小数点对齐（视为零）\n" +
                "否(N)：最右侧对齐（视为普通文本）",
                "对齐方式选择",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            bool treatDashAsZero = (userChoice == DialogResult.Yes);

            Word.Selection selection = app.Selection;
            bool hasSelection = (selection.Type != Word.WdSelectionType.wdSelectionIP);

            app.ScreenUpdating = false;

            try
            {
                if (hasSelection)
                {
                    // 判断是否选中了单个表格单元格
                    bool singleCell = false;
                    try
                    {
                        if (selection.Cells != null && selection.Cells.Count == 1)
                            singleCell = true;
                    }
                    catch { }

                    if (singleCell)
                    {
                        // 单个单元格模式：只调整该单元格，参考同列其他单元格的缩进
                        AdjustSingleCell(selection.Cells[1], treatDashAsZero);
                    }
                    else
                    {
                        // 多单元格/范围选中：收集选中的列索引，只处理这些列
                        HashSet<int> selectedColumns = new HashSet<int>();
                        try
                        {
                            foreach (Word.Cell c in selection.Cells)
                            {
                                selectedColumns.Add(c.ColumnIndex);
                            }
                        }
                        catch { }
                        ProcessTablesInRange(app.ActiveDocument, selection.Range, treatDashAsZero, selectedColumns.Count > 0 ? selectedColumns : null);
                    }
                }
                else
                {
                    // 全局模式：支持批量文件
                    DialogResult batchChoice = MessageBox.Show(
                        "是否要批量处理多个文件？\n" +
                        "是(Y)：选择多个文件进行批量处理\n" +
                        "否(N)：仅处理当前打开的活动文档",
                        "处理范围",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Question);

                    if (batchChoice == DialogResult.Cancel) return;

                    if (batchChoice == DialogResult.Yes)
                    {
                        List<string> filesToProcess = new List<string>();
                        using (OpenFileDialog ofd = new OpenFileDialog())
                        {
                            ofd.Filter = "Word文件|*.docx;*.doc";
                            ofd.Title = "请选择报告文件(可多选)";
                            ofd.Multiselect = true;
                            if (ofd.ShowDialog() != DialogResult.OK || ofd.FileNames.Length == 0)
                            {
                                MessageBox.Show("未选择文件，操作已取消。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }
                            filesToProcess.AddRange(ofd.FileNames);
                        }

                        for (int k = 0; k < filesToProcess.Count; k++)
                        {
                            string filePath = filesToProcess[k];
                            app.StatusBar = $"正在处理: {System.IO.Path.GetFileName(filePath)} ({(k * 100 / filesToProcess.Count)}%)";
                            Word.Document doc = app.Documents.Open(filePath, ReadOnly: false);
                            ProcessTablesInRange(doc, null, treatDashAsZero);
                            doc.Save();
                            doc.Close(Word.WdSaveOptions.wdSaveChanges);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                        }
                        app.StatusBar = "处理完成";
                        MessageBox.Show("批量处理完成，小数点已对齐。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        ProcessTablesInRange(app.ActiveDocument, null, treatDashAsZero);
                        MessageBox.Show("当前文档处理完成，小数点已对齐。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"处理过程中出现错误。\n错误描述：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                app.ScreenUpdating = true;
                app.StatusBar = "";
            }
        }

        /// <summary>
        /// 单元格信息：记录单元格对象、是否是括号负数、当前右边距
        /// </summary>
        private class CellAlignInfo
        {
            public Word.Cell Cell;
            public bool IsBracketNegative;
            public float CurrentRightPadding;
        }

        /// <summary>
        /// 处理文档中的表格。若 selectionRange 不为 null，则只处理与选区相交的表格。
        /// 
        /// 对齐逻辑（使用 Cell RightPadding）：
        /// 1. 右对齐时，负数(括号)比正数多一个右括号")"，导致小数点左移一个括号宽度
        /// 2. 若正数单元格已有右边距 >= 括号宽度 → 负数的右边距减去括号宽度
        /// 3. 若正数单元格右边距较小 → 正数增加括号宽度的右边距
        /// </summary>
        private static void ProcessTablesInRange(Word.Document doc, Word.Range selectionRange, bool treatDashAsZero, HashSet<int> selectedColumns = null)
        {
            Regex fullNumberRegex = new Regex("^" + TextHelper.IndependentNumberRegex.ToString() + "$", RegexOptions.Compiled);

            foreach (Word.Table tbl in doc.Tables)
            {
                try
                {
                    // 如果指定了选区范围，跳过不在选区内的表格
                    if (selectionRange != null)
                    {
                        Word.Range tblRange = tbl.Range;
                        if (tblRange.End < selectionRange.Start || tblRange.Start > selectionRange.End)
                            continue;
                    }

                    if (tbl.Rows.Count < 2) continue;

                    int colCount = 0;
                    try
                    {
                        colCount = tbl.Columns.Count;
                    }
                    catch
                    {
                        foreach (Word.Cell c in tbl.Range.Cells)
                        {
                            if (c.ColumnIndex > colCount) colCount = c.ColumnIndex;
                        }
                    }

                    // 遍历所有列，从第 2 列开始
                    for (int i = 2; i <= colCount; i++)
                    {
                        // 选区模式下，只处理选中的列
                        if (selectedColumns != null && !selectedColumns.Contains(i))
                            continue;
                        List<CellAlignInfo> cellInfos = new List<CellAlignInfo>();
                        bool hasBracketNegative = false;
                        string bracketFontName = "Arial";
                        float bracketFontSize = 10.5f;

                        // 第一轮：收集数字单元格信息
                        for (int r = 1; r <= tbl.Rows.Count; r++)
                        {
                            Word.Cell cell = null;
                            try { cell = tbl.Cell(r, i); }
                            catch { continue; }
                            if (cell == null) continue;

                            // 选区模式下，跳过不在选区内的单元格
                            if (selectionRange != null)
                            {
                                try
                                {
                                    if (cell.Range.End <= selectionRange.Start || cell.Range.Start >= selectionRange.End)
                                        continue;
                                }
                                catch { continue; }
                            }

                            Word.Range rng = cell.Range;
                            rng.End = rng.End - 1;
                            string cellText = CleanCellText(rng.Text);

                            bool isValidNumber = fullNumberRegex.IsMatch(cellText);
                            if (!isValidNumber && treatDashAsZero && cellText == "-")
                            {
                                isValidNumber = true;
                            }

                            if (!isValidNumber) continue;

                            bool isBracket = cellText.StartsWith("(") && cellText.EndsWith(")");

                            float currentPadding = 0f;
                            try { currentPadding = cell.RightPadding; }
                            catch { }
                            if (currentPadding < 0) currentPadding = 0;

                            cellInfos.Add(new CellAlignInfo
                            {
                                Cell = cell,
                                IsBracketNegative = isBracket,
                                CurrentRightPadding = currentPadding
                            });

                            if (isBracket)
                            {
                                hasBracketNegative = true;
                                // 记录括号所使用的字体用于测量宽度
                                string fn = cell.Range.Font.NameAscii;
                                if (string.IsNullOrEmpty(fn)) fn = cell.Range.Font.Name;
                                if (!string.IsNullOrEmpty(fn)) bracketFontName = fn;
                                float fs = cell.Range.Font.Size;
                                if (fs > 0 && fs != (float)Word.WdConstants.wdUndefined)
                                    bracketFontSize = fs;
                            }
                        }

                        // 如果该列没有括号负数，无需对齐
                        if (!hasBracketNegative || cellInfos.Count < 2) continue;

                        // 测量右括号 ")" 的宽度
                        float bracketWidth = MeasureTextWidth(")", bracketFontName, bracketFontSize);

                        // 判断当前列的右边距状态：取正数单元格的代表性右边距
                        float representativePadding = 0f;
                        foreach (CellAlignInfo info in cellInfos)
                        {
                            if (!info.IsBracketNegative)
                            {
                                representativePadding = info.CurrentRightPadding;
                                break;
                            }
                        }

                        // 第二轮：通过 Cell RightPadding 施加对齐
                        // 负数永远设为0（最大化空间），正数设为括号宽度
                        float negNewPadding = 0f;
                        float posNewPadding = bracketWidth;

                        foreach (CellAlignInfo info in cellInfos)
                        {
                            Word.ParagraphFormat pf = info.Cell.Range.ParagraphFormat;
                            pf.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                            pf.RightIndent = 0f;

                            if (info.IsBracketNegative)
                            {
                                info.Cell.RightPadding = negNewPadding;
                            }
                            else
                            {
                                info.Cell.RightPadding = posNewPadding;
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    continue;
                }
            }
        }

        /// <summary>
        /// 单个单元格调整：用户选中一个负数单元格后，参考同列正数的右缩进来调整
        /// </summary>
        private static void AdjustSingleCell(Word.Cell selectedCell, bool treatDashAsZero)
        {
            Regex fullNumberRegex = new Regex("^" + TextHelper.IndependentNumberRegex.ToString() + "$", RegexOptions.Compiled);

            Word.Range rng = selectedCell.Range;
            rng.End = rng.End - 1;
            string cellText = CleanCellText(rng.Text);

            bool isValidNumber = fullNumberRegex.IsMatch(cellText);
            if (!isValidNumber && treatDashAsZero && cellText == "-")
                isValidNumber = true;

            if (!isValidNumber) return;

            bool isBracket = cellText.StartsWith("(") && cellText.EndsWith(")");

            // 获取括号字体信息
            string fontName = selectedCell.Range.Font.NameAscii;
            if (string.IsNullOrEmpty(fontName)) fontName = selectedCell.Range.Font.Name;
            if (string.IsNullOrEmpty(fontName)) fontName = "Arial";
            float fontSize = selectedCell.Range.Font.Size;
            if (fontSize <= 0 || fontSize == (float)Word.WdConstants.wdUndefined) fontSize = 10.5f;

            float bracketWidth = MeasureTextWidth(")", fontName, fontSize);



            // 通过 Cell RightPadding 应用对齐：负数永远0，正数永远括号宽度
            Word.ParagraphFormat pf = selectedCell.Range.ParagraphFormat;
            pf.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            pf.RightIndent = 0f;

            if (isBracket)
            {
                selectedCell.RightPadding = 0f;
            }
            else
            {
                selectedCell.RightPadding = bracketWidth;
            }
        }

        private static string CleanCellText(string rawText)
        {
            if (string.IsNullOrEmpty(rawText)) return string.Empty;
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            foreach (char ch in rawText)
            {
                int code = (int)ch;
                if ((code >= 32 && code <= 126) || code > 255)
                {
                    if (code != 160)
                    {
                        sb.Append(ch);
                    }
                }
            }
            return sb.ToString().Trim();
        }

        private static float MeasureTextWidth(string text, string fontName, float fontSize)
        {
            if (string.IsNullOrEmpty(text)) return 0f;

            using (Bitmap bmp = new Bitmap(1, 1))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                float dpiX = g.DpiX;
                using (Font font = new Font(fontName, fontSize, GraphicsUnit.Point))
                using (StringFormat format = new StringFormat(StringFormat.GenericTypographic))
                {
                    // 使用 MeasureCharacterRanges 获取更精确的字符宽度
                    CharacterRange[] ranges = { new CharacterRange(0, text.Length) };
                    format.SetMeasurableCharacterRanges(ranges);
                    Region[] regions = g.MeasureCharacterRanges(text, font, new RectangleF(0, 0, 1000, 100), format);
                    RectangleF rect = regions[0].GetBounds(g);
                    // 将像素值转换为磅值（1磅 = 1/72英寸，1像素 = 1/DPI英寸）
                    return rect.Width * 72f / dpiX;
                }
            }
        }
    }
}
