using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Newtonsoft.Json;

namespace FunctionBox
{
    public partial class ThisAddIn
    {
        private bool vbaTrustWarningShown;
        private bool sumCheckDebugModeEnabled;
        private static readonly string[] VbaTrustErrorTokens =
        {
            "Programmatic access to Visual Basic Project is not trusted",
            "对 Visual Basic Project 的编程访问不受信任"
        };
        private const string ShortcutModulePrefix = "FunctionBoxShortcut_";
        private static readonly string VbaTrustWarningStateFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "FunctionBox",
            "vba-trust-warning-suppressed.flag");
        private static readonly string VbaCodeSaveFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "FunctionBox",
            "vba_codes.json");
        private static readonly string SumCheckDebugLogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "FunctionBox",
            "sum-check-debug.log");
        private const double SumTolerance = 0.001d;
        private static readonly Regex IndependentNumberRegex = new Regex(
            @"(?:\(-?[0-9][0-9,]*(?:\.\d+)?\)|（-?[0-9][0-9,]*(?:\.\d+)?）|-?[0-9][0-9,]*(?:\.\d+)?)",
            RegexOptions.Compiled);

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            vbaTrustWarningShown = LoadVbaTrustWarningState();

            // 自动尝试在注册表中开启 VBA 信任选项
            EnableVbaTrustAutomatically();

            // 添加事件处理程序-删除临时VBA
            this.Application.DocumentBeforeClose += new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(Application_DocumentBeforeClose);
            SyncVbaShortcutBindings();
        }

        private void EnableVbaTrustAutomatically()
        {
            try
            {
                string version = this.Application.Version;
                string keyPath = @"Software\Microsoft\Office\" + version + @"\Word\Security";
                
                using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(keyPath, true))
                {
                    if (key != null)
                    {
                        object value = key.GetValue("AccessVBOM");
                        if (value == null || (int)value != 1)
                        {
                            key.SetValue("AccessVBOM", 1, Microsoft.Win32.RegistryValueKind.DWord);
                        }
                    }
                    else
                    {
                        using (Microsoft.Win32.RegistryKey newKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(keyPath))
                        {
                            if (newKey != null)
                            {
                                newKey.SetValue("AccessVBOM", 1, Microsoft.Win32.RegistryValueKind.DWord);
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                // 静默失败，如果没有权限修改注册表则退回到常规的手动提示
            }
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 取消事件处理程序
            this.Application.DocumentBeforeClose -= new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(Application_DocumentBeforeClose);
        }
        public Word.Selection GetWordSelection()
        {
            return this.Application.Selection;
        }
        public bool SumCheckDebugModeEnabled
        {
            get { return sumCheckDebugModeEnabled; }
            set { sumCheckDebugModeEnabled = value; }
        }
        public void ValidateSumsHorizontal()
        {
            Word.Application wordApp = this.Application;
            Word.Selection selection = GetWordSelection();
            List<string> debugLines = BeginSumCheckDebug("横向加总检查", selection);

            // 禁用屏幕更新
            wordApp.ScreenUpdating = false;

            try
            {
                if (!TryGetSelectedCells(selection, out Dictionary<int, List<SelectedCellInfo>> cellsByTable, debugLines))
                    return;

                foreach (KeyValuePair<int, List<SelectedCellInfo>> pair in cellsByTable)
                {
                    bool isRectangular = IsRectangularSelection(pair.Value);
                    AppendSumCheckDebug(debugLines, $"Table[{pair.Key}] 选中单元格={pair.Value.Count}, 是否规则矩形={isRectangular}");

                    int validatedGroups = 0;
                    var groupedByRow = pair.Value
                        .GroupBy(cell => cell.RowIndex)
                        .Select(group => group.OrderBy(cell => cell.ColumnIndex).ToList());

                    foreach (List<SelectedCellInfo> cells in groupedByRow)
                    {
                        if (cells.Count < 2)
                        {
                            continue;
                        }

                        ValidateCellGroup(cells, isLastCellAsTarget: true, $"按行分组(Row={cells[0].RowIndex})", debugLines);
                        validatedGroups++;
                    }

                    // 如果没有任何一行有 >= 2 个单元格，进行兜底
                    if (validatedGroups == 0 && pair.Value.Count >= 2)
                    {
                        ValidateAsSingleSequence(pair.Value, isLastCellAsTarget: true, "单序列兜底", debugLines);
                    }
                }
            }
            finally
            {
                // 启用屏幕更新
                wordApp.ScreenUpdating = true;
                EndSumCheckDebug(debugLines);
            }
        }
        public void ValidateSumsVerticalTop()
        {
            Word.Application wordApp = this.Application;
            Word.Selection selection = GetWordSelection();
            List<string> debugLines = BeginSumCheckDebug("竖向加总检查-自下向上", selection);

            // 禁用屏幕更新
            wordApp.ScreenUpdating = false;

            try
            {
                if (!TryGetSelectedCells(selection, out Dictionary<int, List<SelectedCellInfo>> cellsByTable, debugLines))
                    return;

                foreach (KeyValuePair<int, List<SelectedCellInfo>> pair in cellsByTable)
                {
                    bool isRectangular = IsRectangularSelection(pair.Value);
                    AppendSumCheckDebug(debugLines, $"Table[{pair.Key}] 选中单元格={pair.Value.Count}, 是否规则矩形={isRectangular}");

                    int validatedGroups = 0;
                    var groupedByColumn = pair.Value
                        .GroupBy(cell => cell.ColumnIndex)
                        .Select(group => group.OrderBy(cell => cell.RowIndex).ToList());

                    foreach (List<SelectedCellInfo> cells in groupedByColumn)
                    {
                        if (cells.Count < 2)
                        {
                            continue;
                        }

                        ValidateCellGroup(cells, isLastCellAsTarget: false, $"按列分组(Col={cells[0].ColumnIndex})", debugLines);
                        validatedGroups++;
                    }

                    if (validatedGroups == 0 && pair.Value.Count >= 2)
                    {
                        ValidateAsSingleSequence(pair.Value, isLastCellAsTarget: false, "单序列兜底", debugLines);
                    }
                }
            }
            finally
            {
                // 启用屏幕更新
                wordApp.ScreenUpdating = true;
                EndSumCheckDebug(debugLines);
            }
        }
        public void ValidateSumsVerticalDown()
        {
            Word.Application wordApp = this.Application;
            Word.Selection selection = GetWordSelection();
            List<string> debugLines = BeginSumCheckDebug("竖向加总检查-自上向下", selection);

            // 禁用屏幕更新
            wordApp.ScreenUpdating = false;

            try
            {
                if (!TryGetSelectedCells(selection, out Dictionary<int, List<SelectedCellInfo>> cellsByTable, debugLines))
                    return;

                foreach (KeyValuePair<int, List<SelectedCellInfo>> pair in cellsByTable)
                {
                    bool isRectangular = IsRectangularSelection(pair.Value);
                    AppendSumCheckDebug(debugLines, $"Table[{pair.Key}] 选中单元格={pair.Value.Count}, 是否规则矩形={isRectangular}");

                    int validatedGroups = 0;
                    var groupedByColumn = pair.Value
                        .GroupBy(cell => cell.ColumnIndex)
                        .Select(group => group.OrderBy(cell => cell.RowIndex).ToList());

                    foreach (List<SelectedCellInfo> cells in groupedByColumn)
                    {
                        if (cells.Count < 2)
                        {
                            continue;
                        }

                        ValidateCellGroup(cells, isLastCellAsTarget: true, $"按列分组(Col={cells[0].ColumnIndex})", debugLines);
                        validatedGroups++;
                    }

                    if (validatedGroups == 0 && pair.Value.Count >= 2)
                    {
                        ValidateAsSingleSequence(pair.Value, isLastCellAsTarget: true, "单序列兜底", debugLines);
                    }
                }
            }
            finally
            {
                // 启用屏幕更新
                wordApp.ScreenUpdating = true;
                EndSumCheckDebug(debugLines);
            }
        }
        private bool TryGetSelectedCells(
            Word.Selection selection,
            out Dictionary<int, List<SelectedCellInfo>> cellsByTable,
            List<string> debugLines)
        {
            cellsByTable = new Dictionary<int, List<SelectedCellInfo>>();

            if (selection == null || selection.Tables.Count == 0)
            {
                AppendSumCheckDebug(debugLines, "未检测到表格选区");
                MessageBox.Show("请选择表格区域");
                return false;
            }

            List<SelectedCellInfo> selectedCells = CollectSelectedCellsDirect(selection);
            AppendSumCheckDebug(debugLines, $"直接读取 Selection.Cells 得到 {selectedCells.Count} 个单元格");

            bool needsFallback = selectedCells.Count < 2;
            if (!needsFallback)
            {
                try
                {
                    if (selection.Range.Cells.Count > selectedCells.Count)
                    {
                        needsFallback = true;
                        AppendSumCheckDebug(debugLines, $"检测到选区外框包围的单元格数({selection.Range.Cells.Count})大于 Selection.Cells，确认为非连续选区");
                    }
                }
                catch
                {
                    needsFallback = true; // 如果报错说明表格复杂，稳妥起见执行反查
                    AppendSumCheckDebug(debugLines, $"检测外框单元格时引发异常，降级执行反查");
                }
            }

            if (needsFallback)
            {
                List<SelectedCellInfo> fallbackCells = CollectSelectedCellsByToggleMarker(selection, debugLines);
                if (fallbackCells.Count > 0)
                {
                    List<SelectedCellInfo> mergedCells = MergeSelectedCells(selectedCells, fallbackCells);
                    if (mergedCells.Count != selectedCells.Count ||
                        !HaveSameSelectedCellCoordinates(selectedCells, mergedCells))
                    {
                        selectedCells = mergedCells;
                        AppendSumCheckDebug(debugLines, $"已合并反查结果，共 {selectedCells.Count} 个单元格");
                    }
                }
            }

            for (int index = 0; index < selectedCells.Count; index++)
            {
                SelectedCellInfo selectedCell = selectedCells[index];
                Word.Cell cell = selectedCell.Cell;
                Word.Table table = cell.Range.Tables[1];
                int tableKey = table.Range.Start;

                if (!cellsByTable.TryGetValue(tableKey, out List<SelectedCellInfo> tableCells))
                {
                    tableCells = new List<SelectedCellInfo>();
                    cellsByTable[tableKey] = tableCells;
                }

                selectedCell.SelectionOrder = index + 1;
                tableCells.Add(selectedCell);
            }

            if (cellsByTable.Count == 0)
            {
                AppendSumCheckDebug(debugLines, "有效选中单元格数为 0");
                MessageBox.Show("请选择表格区域");
                return false;
            }

            return true;
        }
        private static List<SelectedCellInfo> CollectSelectedCellsDirect(Word.Selection selection)
        {
            List<SelectedCellInfo> result = new List<SelectedCellInfo>();

            if (selection == null || selection.Cells.Count == 0)
            {
                return result;
            }

            for (int index = 1; index <= selection.Cells.Count; index++)
            {
                Word.Cell cell = selection.Cells[index];
                result.Add(new SelectedCellInfo
                {
                    Cell = cell,
                    RowIndex = cell.RowIndex,
                    ColumnIndex = cell.ColumnIndex,
                    SelectionOrder = index
                });
            }

            return result;
        }
        private static List<SelectedCellInfo> MergeSelectedCells(List<SelectedCellInfo> directCells, List<SelectedCellInfo> fallbackCells)
        {
            List<SelectedCellInfo> mergedCells = new List<SelectedCellInfo>();
            HashSet<string> directCoordinates = new HashSet<string>();

            if (directCells != null)
            {
                foreach (SelectedCellInfo directCell in directCells)
                {
                    directCoordinates.Add(GetCellCoordinateKey(directCell.RowIndex, directCell.ColumnIndex));
                }
            }

            if (fallbackCells != null)
            {
                foreach (SelectedCellInfo fallbackCell in fallbackCells)
                {
                    string coordinate = GetCellCoordinateKey(fallbackCell.RowIndex, fallbackCell.ColumnIndex);
                    if (!directCoordinates.Contains(coordinate))
                    {
                        mergedCells.Add(fallbackCell);
                    }
                }
            }

            if (directCells != null)
            {
                mergedCells.AddRange(directCells);
            }

            return mergedCells;
        }
        private static bool HaveSameSelectedCellCoordinates(List<SelectedCellInfo> leftCells, List<SelectedCellInfo> rightCells)
        {
            if (leftCells == null || rightCells == null)
            {
                return leftCells == rightCells;
            }

            if (leftCells.Count != rightCells.Count)
            {
                return false;
            }

            HashSet<string> coordinates = new HashSet<string>(
                leftCells.Select(cell => GetCellCoordinateKey(cell.RowIndex, cell.ColumnIndex)));

            foreach (SelectedCellInfo rightCell in rightCells)
            {
                if (!coordinates.Remove(GetCellCoordinateKey(rightCell.RowIndex, rightCell.ColumnIndex)))
                {
                    return false;
                }
            }

            return coordinates.Count == 0;
        }
        private List<SelectedCellInfo> CollectSelectedCellsByToggleMarker(Word.Selection selection, List<string> debugLines)
        {
            List<SelectedCellInfo> result = new List<SelectedCellInfo>();
            if (selection == null || selection.Range == null || selection.Range.Tables.Count == 0) return result;

            Word.Table table = selection.Range.Tables[1];
            // 使用一个极其罕见的字间距值作为标记
            float markerKerning = 9.87f;

            try
            {
                // 1. 给选区打上标记（一次 COM 调用）
                selection.Font.Kerning = markerKerning;

                // 2. 遍历表格所有单元格，检查哪些被标记了
                //    使用 table.Range.Cells 平铺集合，安全处理合并单元格
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    try
                    {
                        float k = cell.Range.Font.Kerning;
                        // 浮点比较用容差
                        if (Math.Abs(k - markerKerning) < 0.01f)
                        {
                            result.Add(new SelectedCellInfo
                            {
                                Cell = cell,
                                RowIndex = cell.RowIndex,
                                ColumnIndex = cell.ColumnIndex
                            });
                        }
                    }
                    catch { /* 跳过无法读取的合并残留格 */ }
                }
            }
            catch (Exception ex)
            {
                AppendSumCheckDebug(debugLines, "非连续选区反查失败: " + ex.Message);
            }
            finally
            {
                // 3. 使用撤销完美恢复所有格式（不会破坏任何字体属性）
                try { Application.ActiveDocument.Undo(); } catch { }
            }

            AppendSumCheckDebug(debugLines, $"非连续选区反查得到 {result.Count} 个单元格");

            // 去重并排序
            return result
                .GroupBy(c => c.RowIndex + "_" + c.ColumnIndex)
                .Select(g => g.First())
                .OrderBy(c => c.RowIndex)
                .ThenBy(c => c.ColumnIndex)
                .ToList();
        }
        private static string GetCellCoordinateKey(int rowIndex, int columnIndex)
        {
            return rowIndex + ":" + columnIndex;
        }
        private static bool IsRectangularSelection(List<SelectedCellInfo> cells)
        {
            if (cells == null || cells.Count == 0)
            {
                return false;
            }

            int minRow = cells.Min(cell => cell.RowIndex);
            int maxRow = cells.Max(cell => cell.RowIndex);
            int minColumn = cells.Min(cell => cell.ColumnIndex);
            int maxColumn = cells.Max(cell => cell.ColumnIndex);

            int expectedCount = (maxRow - minRow + 1) * (maxColumn - minColumn + 1);
            if (expectedCount != cells.Count)
            {
                return false;
            }

            HashSet<string> coordinates = new HashSet<string>();
            foreach (SelectedCellInfo cell in cells)
            {
                coordinates.Add(cell.RowIndex + ":" + cell.ColumnIndex);
            }

            return coordinates.Count == expectedCount;
        }
        private void ValidateAsSingleSequence(List<SelectedCellInfo> cells, bool isLastCellAsTarget, string context, List<string> debugLines)
        {
            if (cells == null || cells.Count < 2)
            {
                AppendSumCheckDebug(debugLines, $"{context}: 可用单元格不足2个，跳过");
                return;
            }

            List<SelectedCellInfo> orderedCells = cells
                .OrderBy(cell => cell.SelectionOrder)
                .ToList();

            int targetIndex = isLastCellAsTarget ? orderedCells.Count - 1 : 0;
            Word.Cell targetCell = orderedCells[targetIndex].Cell;
            double sum = 0;

            for (int index = 0; index < orderedCells.Count; index++)
            {
                if (index == targetIndex)
                {
                    continue;
                }

                sum += ParseCellValue(GetCellText(orderedCells[index].Cell));
            }

            double targetValue = ParseCellValue(GetCellText(targetCell));
            bool isMatch = Math.Abs(sum - targetValue) < SumTolerance;

            // 彻底去除文字高亮，并将所有可能的底纹（段落、文字、单元格）全部强制刷成目标颜色
            Word.WdColor color = isMatch ? Word.WdColor.wdColorLightGreen : Word.WdColor.wdColorYellow;
            targetCell.Range.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
            try { targetCell.Range.Font.Shading.BackgroundPatternColor = color; } catch { }
            try { targetCell.Range.ParagraphFormat.Shading.BackgroundPatternColor = color; } catch { }
            try { targetCell.Range.Shading.BackgroundPatternColor = color; } catch { }
            targetCell.Shading.BackgroundPatternColor = color;

            string resultIcon = isMatch ? "通过 ✅" : "失败 ❌";
            AppendSumCheckDebug(
                debugLines,
                $"\n  [{context}]\n" +
                $"    - 校验结果 : {resultIcon}\n" +
                $"    - 目标单元格: {FormatCellIdentity(orderedCells[targetIndex])}\n" +
                $"    - 参与求和格: {FormatCellSequence(orderedCells, targetIndex)}\n" +
                $"    - 实际计算和: {sum:N2}");
        }
        private void ValidateCellGroup(List<SelectedCellInfo> cells, bool isLastCellAsTarget, string context, List<string> debugLines)
        {
            if (cells == null || cells.Count < 2)
            {
                AppendSumCheckDebug(debugLines, $"{context}: 可用单元格不足2个，跳过");
                return;
            }

            int targetIndex = isLastCellAsTarget ? cells.Count - 1 : 0;
            Word.Cell targetCell = cells[targetIndex].Cell;
            double sum = 0;

            for (int index = 0; index < cells.Count; index++)
            {
                if (index == targetIndex)
                {
                    continue;
                }

                sum += ParseCellValue(GetCellText(cells[index].Cell));
            }

            double targetValue = ParseCellValue(GetCellText(targetCell));
            bool isMatch = Math.Abs(sum - targetValue) < SumTolerance;

            // 彻底去除文字高亮，并将所有可能的底纹（段落、文字、单元格）全部强制刷成目标颜色
            Word.WdColor color = isMatch ? Word.WdColor.wdColorLightGreen : Word.WdColor.wdColorYellow;
            targetCell.Range.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
            try { targetCell.Range.Font.Shading.BackgroundPatternColor = color; } catch { }
            try { targetCell.Range.ParagraphFormat.Shading.BackgroundPatternColor = color; } catch { }
            try { targetCell.Range.Shading.BackgroundPatternColor = color; } catch { }
            targetCell.Shading.BackgroundPatternColor = color;

            string resultIcon = isMatch ? "通过 ✅" : "失败 ❌";
            AppendSumCheckDebug(
                debugLines,
                $"\n  [{context}]\n" +
                $"    - 校验结果 : {resultIcon}\n" +
                $"    - 目标单元格: {FormatCellIdentity(cells[targetIndex])}\n" +
                $"    - 参与求和格: {FormatCellSequence(cells, targetIndex)}\n" +
                $"    - 实际计算和: {sum:N2}");
        }
        private List<string> BeginSumCheckDebug(string modeName, Word.Selection selection)
        {
            if (!sumCheckDebugModeEnabled)
            {
                return null;
            }

            return new List<string>
            {
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {modeName}",
                $"选区统计: Tables={selection.Tables.Count}, Cells={selection.Cells.Count}"
            };
        }
        private void EndSumCheckDebug(List<string> debugLines)
        {
            if (debugLines == null)
            {
                return;
            }

            try
            {
                string directory = Path.GetDirectoryName(SumCheckDebugLogPath);
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                string block = "\r\n============================================================\r\n" + 
                               string.Join(Environment.NewLine, debugLines) + 
                               "\r\n============================================================\r\n";
                File.AppendAllText(SumCheckDebugLogPath, block);
            }
            catch
            {
                // 调试日志失败不影响主流程
            }

            // 弹出显示格式化的校验结果摘要
            string preview = string.Join(Environment.NewLine, debugLines);
            if (preview.Length > 1500)
            {
                preview = preview.Substring(0, 1500) + "\n...\n(内容过长已截断，完整信息请查看日志文件)";
            }
            MessageBox.Show(
                preview,
                "加总检查调试报告",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        private static void AppendSumCheckDebug(List<string> debugLines, string line)
        {
            if (debugLines == null)
            {
                return;
            }

            debugLines.Add(line);
        }
        private static string FormatCellIdentity(SelectedCellInfo cell)
        {
            if (cell == null) return "null";
            return $"第{cell.RowIndex}行第{cell.ColumnIndex}列({GetCellText(cell.Cell)})";
        }
        private static string FormatCellSequence(List<SelectedCellInfo> cells, int excludeIndex = -1)
        {
            if (cells == null || cells.Count == 0) return "[]";
            var participating = excludeIndex >= 0 ? cells.Where((c, i) => i != excludeIndex).ToList() : cells;
            return string.Join(", ", participating.Select(c => $"第{c.RowIndex}行第{c.ColumnIndex}列({GetCellText(c.Cell)})"));
        }
        private static string GetCellText(Word.Cell cell)
        {
            return cell.Range.Text
                .Replace("\r\a", string.Empty)
                .Replace("\a", string.Empty)
                .Replace("\r", string.Empty)
                .Trim();
        }
        private static double ParseCellValue(string cellText)
        {
            if (string.IsNullOrWhiteSpace(cellText))
            {
                return 0;
            }

            string normalized = cellText
                .Replace("\u00A0", " ")
                .Replace("，", ",")
                .Replace("－", "-")
                .Replace("—", "-")
                .Replace("（", "(")
                .Replace("）", ")")
                .Trim();

            bool isPercent = normalized.EndsWith("%", StringComparison.Ordinal);
            if (isPercent)
            {
                normalized = normalized.Substring(0, normalized.Length - 1).Trim();
            }

            NumberStyles styles = NumberStyles.Number | NumberStyles.AllowParentheses | NumberStyles.AllowLeadingSign;
            double value;
            bool parsed = double.TryParse(normalized, styles, CultureInfo.CurrentCulture, out value) ||
                double.TryParse(normalized, styles, CultureInfo.InvariantCulture, out value);

            if (!parsed)
            {
                return 0;
            }

            if (isPercent)
            {
                value /= 100;
            }

            return value;
        }
        private sealed class SelectedCellInfo
        {
            public Word.Cell Cell { get; set; }
            public int RowIndex { get; set; }
            public int ColumnIndex { get; set; }
            public int SelectionOrder { get; set; }
        }
        public void ClearSelectionBackground()
        {
            Word.Application wordApp = this.Application;
            Word.Selection selection = GetWordSelection();

            // 禁用屏幕更新
            wordApp.ScreenUpdating = false;

            try
            {
                // 全面清除选中区域的所有背景色和高亮色
                Word.Range selectedRange = selection.Range;
                selectedRange.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
                try { selectedRange.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic; } catch { }
                try { selectedRange.ParagraphFormat.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic; } catch { }
                try { selectedRange.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic; } catch { }

                // 确保同时也清除了单元格本身的背景色
                if (selection.Cells.Count > 0)
                {
                    try { selection.Cells.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic; } catch { }
                }
            }
            finally
            {
                // 启用屏幕更新
                wordApp.ScreenUpdating = true;
            }
        }
        public void ClearDocumentBackground()
        {
            Word.Application wordApp = this.Application;
            Word.Document document = wordApp.ActiveDocument;

            // 禁用屏幕更新
            wordApp.ScreenUpdating = false;

            try
            {
                // 获取整个文档的范围
                Word.Range entireDocumentRange = document.Content;

                // 全面清除整个文档的背景色和高亮色
                entireDocumentRange.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
                try { entireDocumentRange.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic; } catch { }
                try { entireDocumentRange.ParagraphFormat.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic; } catch { }
                try { entireDocumentRange.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic; } catch { }

                // 由于 entireDocumentRange 已经处理了所有的 Font 和 ParagraphFormat 级别，
                // 我们只需要清理表格自身的 Shading，减少大量 COM 调用提升速度。
                foreach (Word.Table table in document.Tables)
                {
                    try { table.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic; } catch { }
                }

                // 清除形状（文本框等）的背景色
                foreach (Word.Shape shape in document.Shapes)
                {
                    if (shape.TextFrame.HasText != 0)
                    {
                        shape.TextFrame.TextRange.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
                        shape.TextFrame.TextRange.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
                    }
                }
            }
            finally
            {
                // 启用屏幕更新
                wordApp.ScreenUpdating = true;
            }
        }
        private void Application_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            try
            {
                if (Doc == null)
                {
                    return;
                }

                // 获取当前正在关闭文档的 VBA 项目
                var vbaProject = Doc.VBProject;

                // 删除该文档下的所有临时模块
                for (int i = vbaProject.VBComponents.Count; i >= 1; i--)
                {
                    var vbComponent = vbaProject.VBComponents.Item(i);
                    if (vbComponent.Name.StartsWith("TempMacro_"))
                    {
                        vbaProject.VBComponents.Remove(vbComponent);
                    }
                }
            }
            catch (Exception ex)
            {
                if (HandleVbaTrustErrorOnce(ex))
                {
                    return;
                }

                MessageBox.Show($"删除临时宏时出错: {ex.Message}");
            }
        }
        public void SyncVbaShortcutBindings()
        {
            try
            {
                List<StoredVbaCode> codes = new List<StoredVbaCode>();
                if (File.Exists(VbaCodeSaveFilePath))
                {
                    string json = File.ReadAllText(VbaCodeSaveFilePath);
                    codes = JsonConvert.DeserializeObject<List<StoredVbaCode>>(json) ?? new List<StoredVbaCode>();
                }

                ApplyVbaShortcutBindings(codes);
            }
            catch (Exception ex)
            {
                if (HandleVbaTrustErrorOnce(ex, "未启用“信任对 VBA 项目对象模型的访问”。后续将静默跳过快捷键自动绑定。"))
                {
                    return;
                }

                MessageBox.Show("同步 VBA 快捷键失败: " + ex.Message);
            }
        }
        private void ApplyVbaShortcutBindings(List<StoredVbaCode> codes)
        {
            Word.Application app = this.Application;
            if (app == null || app.NormalTemplate == null)
            {
                return;
            }

            ClearFunctionBoxShortcutBindings(app);
            ClearFunctionBoxShortcutModules(app);

            object previousContext = app.CustomizationContext;
            app.CustomizationContext = app.NormalTemplate;

            try
            {
                int moduleIndex = 1;
                foreach (StoredVbaCode code in codes)
                {
                    if (code == null ||
                        string.IsNullOrWhiteSpace(code.Shortcut) ||
                        string.IsNullOrWhiteSpace(code.Code) ||
                        !VbaManagerForm.TryExtractMacroName(code.Code, out string macroName))
                    {
                        continue;
                    }

                    if (!TryBuildWordKeyCode(app, code.Shortcut, out int wordKeyCode))
                    {
                        continue;
                    }

                    ClearKeyBindingByCode(app, wordKeyCode);

                    var module = app.NormalTemplate.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                    string moduleName = ShortcutModulePrefix + moduleIndex;
                    module.Name = moduleName;
                    module.CodeModule.AddFromString(code.Code);

                    string command = moduleName + "." + macroName;
                    dynamic keyBindings = app.KeyBindings;
                    keyBindings.Add(Word.WdKeyCategory.wdKeyCategoryMacro, command, wordKeyCode);

                    moduleIndex++;
                }
            }
            finally
            {
                app.CustomizationContext = previousContext;
            }
        }
        private static void ClearFunctionBoxShortcutBindings(Word.Application app)
        {
            for (int index = app.KeyBindings.Count; index >= 1; index--)
            {
                try
                {
                    Word.KeyBinding binding = app.KeyBindings[index];
                    string command = binding.Command;
                    if (!string.IsNullOrWhiteSpace(command) &&
                        command.IndexOf(ShortcutModulePrefix, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        binding.Clear();
                    }
                }
                catch
                {
                    // 单条绑定清理失败不影响其他清理
                }
            }
        }
        private static void ClearFunctionBoxShortcutModules(Word.Application app)
        {
            var vbaProject = app.NormalTemplate.VBProject;
            for (int index = vbaProject.VBComponents.Count; index >= 1; index--)
            {
                var component = vbaProject.VBComponents.Item(index);
                if (component.Name.StartsWith(ShortcutModulePrefix, StringComparison.OrdinalIgnoreCase))
                {
                    vbaProject.VBComponents.Remove(component);
                }
            }
        }
        private static void ClearKeyBindingByCode(Word.Application app, int keyCode)
        {
            for (int index = app.KeyBindings.Count; index >= 1; index--)
            {
                try
                {
                    Word.KeyBinding binding = app.KeyBindings[index];
                    if (binding.KeyCode == keyCode)
                    {
                        binding.Clear();
                    }
                }
                catch
                {
                    // 单条绑定清理失败不影响其他清理
                }
            }
        }
        private static bool TryBuildWordKeyCode(Word.Application app, string shortcut, out int keyCode)
        {
            keyCode = 0;
            if (app == null || string.IsNullOrWhiteSpace(shortcut))
            {
                return false;
            }

            string[] tokens = shortcut
                .Split(new[] { '+' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(token => token.Trim().ToUpperInvariant())
                .Where(token => !string.IsNullOrWhiteSpace(token))
                .ToArray();

            if (tokens.Length == 0 || tokens.Length > 4)
            {
                return false;
            }

            List<Word.WdKey> parts = new List<Word.WdKey>();
            bool hasCtrl = false;
            bool hasShift = false;
            bool hasAlt = false;
            Word.WdKey mainKey = Word.WdKey.wdNoKey;

            foreach (string token in tokens)
            {
                if ((token == "CTRL" || token == "CONTROL") && !hasCtrl)
                {
                    parts.Add(Word.WdKey.wdKeyControl);
                    hasCtrl = true;
                    continue;
                }

                if (token == "SHIFT" && !hasShift)
                {
                    parts.Add(Word.WdKey.wdKeyShift);
                    hasShift = true;
                    continue;
                }

                if (token == "ALT" && !hasAlt)
                {
                    parts.Add(Word.WdKey.wdKeyAlt);
                    hasAlt = true;
                    continue;
                }

                if (!TryParseMainWdKey(token, out mainKey))
                {
                    return false;
                }
            }

            if (mainKey == Word.WdKey.wdNoKey)
            {
                return false;
            }

            parts.Add(mainKey);

            object key1 = parts.Count > 0 ? (object)parts[0] : Type.Missing;
            object key2 = parts.Count > 1 ? (object)parts[1] : Type.Missing;
            object key3 = parts.Count > 2 ? (object)parts[2] : Type.Missing;
            object key4 = parts.Count > 3 ? (object)parts[3] : Type.Missing;
            dynamic appDynamic = app;
            keyCode = appDynamic.BuildKeyCode(key1, key2, key3, key4);
            return true;
        }
        private static bool TryParseMainWdKey(string token, out Word.WdKey key)
        {
            key = Word.WdKey.wdNoKey;
            if (string.IsNullOrWhiteSpace(token))
            {
                return false;
            }

            string normalized = token.Trim().ToUpperInvariant();
            if (TryParseSpecialWdKey(normalized, out key))
            {
                return true;
            }

            string enumName = null;

            if (normalized.Length == 1 && char.IsLetter(normalized[0]))
            {
                enumName = "wdKey" + normalized;
            }
            else if (normalized.Length == 1 && char.IsDigit(normalized[0]))
            {
                enumName = "wdKey" + normalized;
            }
            else if (normalized.StartsWith("F", StringComparison.Ordinal) &&
                int.TryParse(normalized.Substring(1), out int fNumber) &&
                fNumber >= 1 &&
                fNumber <= 12)
            {
                enumName = "wdKeyF" + fNumber;
            }
            else
            {
                return false;
            }

            if (Enum.TryParse(enumName, true, out Word.WdKey parsed))
            {
                key = parsed;
                return true;
            }

            return false;
        }
        private static bool TryParseSpecialWdKey(string normalizedToken, out Word.WdKey key)
        {
            key = Word.WdKey.wdNoKey;
            if (string.IsNullOrWhiteSpace(normalizedToken))
            {
                return false;
            }

            switch (normalizedToken)
            {
                case "`":
                case "OEMTILDE":
                case "OEM3":
                    key = Word.WdKey.wdKeyBackSingleQuote;
                    return true;
                case "'":
                case "OEMQUOTES":
                case "OEM7":
                    key = Word.WdKey.wdKeySingleQuote;
                    return true;
                case "-":
                case "OEMMINUS":
                    key = Word.WdKey.wdKeyHyphen;
                    return true;
                case "=":
                case "OEMPLUS":
                    key = Word.WdKey.wdKeyEquals;
                    return true;
                case "[":
                case "OEMOPENBRACKETS":
                case "OEM4":
                    key = Word.WdKey.wdKeyOpenSquareBrace;
                    return true;
                case "]":
                case "OEMCLOSEBRACKETS":
                case "OEM6":
                    key = Word.WdKey.wdKeyCloseSquareBrace;
                    return true;
                case "\\":
                case "OEMPIPE":
                case "OEMBACKSLASH":
                case "OEM5":
                    key = Word.WdKey.wdKeyBackSlash;
                    return true;
                case ";":
                case "OEMSEMICOLON":
                case "OEM1":
                    key = Word.WdKey.wdKeySemiColon;
                    return true;
                case ",":
                case "OEMCOMMA":
                    key = Word.WdKey.wdKeyComma;
                    return true;
                case ".":
                case "OEMPERIOD":
                    key = Word.WdKey.wdKeyPeriod;
                    return true;
                case "/":
                case "OEMQUESTION":
                case "OEM2":
                    key = Word.WdKey.wdKeySlash;
                    return true;
                default:
                    if (normalizedToken.StartsWith("NUM", StringComparison.Ordinal) &&
                        int.TryParse(normalizedToken.Substring(3), out int numValue) &&
                        numValue >= 0 &&
                        numValue <= 9)
                    {
                        string enumName = "wdKeyNumeric" + numValue;
                        if (Enum.TryParse(enumName, true, out Word.WdKey parsed))
                        {
                            key = parsed;
                            return true;
                        }
                    }
                    return false;
            }
        }

        private bool HandleVbaTrustErrorOnce(Exception ex)
        {
            return HandleVbaTrustErrorOnce(ex, "未启用“信任对 VBA 项目对象模型的访问”。后续将静默跳过临时宏清理。");
        }
        private bool HandleVbaTrustErrorOnce(Exception ex, string warningMessage)
        {
            if (!IsVbaTrustError(ex))
            {
                return false;
            }

            if (!vbaTrustWarningShown)
            {
                string helpText = warningMessage + "\n\n" +
                                  "要使这些功能生效，请在 Word 中手动开启权限：\n" +
                                  "1. 点击左上角【文件】->【选项】\n" +
                                  "2. 选择【信任中心】-> 点击【信任中心设置】\n" +
                                  "3. 选择【宏设置】\n" +
                                  "4. 勾选【信任对 VBA 项目对象模型的访问】并确定。";
                MessageBox.Show(helpText, "需要开启 VBA 信任权限", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                vbaTrustWarningShown = true;
                SaveVbaTrustWarningState();
            }

            return true;
        }

        private static bool IsVbaTrustError(Exception ex)
        {
            Exception current = ex;
            while (current != null)
            {
                foreach (string token in VbaTrustErrorTokens)
                {
                    if (!string.IsNullOrWhiteSpace(current.Message) &&
                        current.Message.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return true;
                    }
                }

                current = current.InnerException;
            }

            return false;
        }

        private static bool LoadVbaTrustWarningState()
        {
            try
            {
                return File.Exists(VbaTrustWarningStateFilePath);
            }
            catch
            {
                return false;
            }
        }
        private sealed class StoredVbaCode
        {
            public string Name { get; set; }
            public string Code { get; set; }
            public string Remark { get; set; }
            public string Shortcut { get; set; }
        }

        private static void SaveVbaTrustWarningState()
        {
            try
            {
                string directory = Path.GetDirectoryName(VbaTrustWarningStateFilePath);
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                File.WriteAllText(VbaTrustWarningStateFilePath, DateTime.UtcNow.ToString("O"));
            }
            catch
            {
                // 持久化失败不影响主流程
            }
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
        private void ApplyTextProcessing(string undoName, Action<Word.Range> processAction)
        {
            Word.Application wordApp = this.Application;
            if (wordApp.Documents.Count == 0) return;
            
            Word.Document doc = wordApp.ActiveDocument;
            Word.Selection selection = wordApp.Selection;

            wordApp.UndoRecord.StartCustomRecord(undoName);
            wordApp.ScreenUpdating = false;

            try
            {
                if (selection.Type == Word.WdSelectionType.wdSelectionIP)
                {
                    // 没有选中具体内容，应用到全文（遍历所有的 StoryRanges，包括正文、页眉页脚、表格等）
                    foreach (Word.Range storyRange in doc.StoryRanges)
                    {
                        Word.Range currentRange = storyRange;
                        while (currentRange != null)
                        {
                            processAction(currentRange);
                            currentRange = currentRange.NextStoryRange;
                        }
                    }
                }
                else
                {
                    // 只应用到选中的内容
                    Word.Range range = selection.Range;
                    processAction(range);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"处理时出错: {ex.Message}");
            }
            finally
            {
                wordApp.ScreenUpdating = true;
                wordApp.UndoRecord.EndCustomRecord();
            }
        }

        public void AddThousandSeparator()
        {
            ApplyTextProcessing("添加千分符", range =>
            {
                string content = range.Text;
                if (string.IsNullOrEmpty(content)) return;

                MatchCollection matches = IndependentNumberRegex.Matches(content);
                HashSet<string> processed = new HashSet<string>();

                for (int index = 0; index < matches.Count; index++)
                {
                    Match match = matches[index];
                    if (!ShouldFormatNumberString(match.Value))
                    {
                        continue;
                    }

                    if (!TryFormatIndependentNumber(match.Value, out string formattedText))
                    {
                        continue;
                    }

                    if (string.Equals(match.Value, formattedText, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    if (processed.Contains(match.Value))
                    {
                        continue;
                    }
                    processed.Add(match.Value);

                    Word.Range searchRange = range.Duplicate;
                    Word.Find find = searchRange.Find;
                    find.ClearFormatting();
                    find.Text = match.Value;
                    find.MatchWildcards = false;
                    find.MatchWholeWord = false;
                    find.Wrap = Word.WdFindWrap.wdFindStop;

                    while (find.Execute())
                    {
                        if (searchRange.InRange(range))
                        {
                            if (ShouldFormatRange(searchRange))
                            {
                                string originalAsciiFont = searchRange.Font.NameAscii;
                                searchRange.Text = formattedText;
                                if (!string.IsNullOrWhiteSpace(originalAsciiFont))
                                {
                                    searchRange.Font.NameAscii = originalAsciiFont;
                                }
                                searchRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                            }
                            else
                            {
                                searchRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                            }
                        }
                    }
                }
            });
        }
        private static bool ShouldFormatNumberString(string matchValue)
        {
            if (string.IsNullOrWhiteSpace(matchValue))
            {
                return false;
            }

            string numericText = matchValue;
            string openingBracket;
            string closingBracket;
            UnwrapNumericWrapper(ref numericText, out openingBracket, out closingBracket);

            int decimalPointIndex = numericText.IndexOf('.');
            string integerPart = decimalPointIndex >= 0
                ? numericText.Substring(0, decimalPointIndex)
                : numericText;

            integerPart = integerPart.Replace(",", string.Empty);
            bool isNegative = integerPart.StartsWith("-", StringComparison.Ordinal);
            if (isNegative)
            {
                integerPart = integerPart.Substring(1);
            }
            return integerPart.Length >= 4 && !integerPart.StartsWith("0", StringComparison.Ordinal);
        }

        private static bool ShouldFormatRange(Word.Range foundRange)
        {
            string prevChar = "";
            try
            {
                if (foundRange.Start > 0)
                {
                    Word.Range prevRange = foundRange.Document.Range(foundRange.Start - 1, foundRange.Start);
                    prevChar = prevRange.Text;
                }
            }
            catch { }

            string nextChar = "";
            try
            {
                if (foundRange.End < foundRange.Document.Content.End)
                {
                    Word.Range nextRange = foundRange.Document.Range(foundRange.End, foundRange.End + 1);
                    nextChar = nextRange.Text;
                }
            }
            catch { }

            if (!string.IsNullOrEmpty(prevChar))
            {
                char value = prevChar[0];
                if (IsAsciiIdentifierCharacter(value) || value == '.' || value == ',' || value == '/' || value == '\\')
                {
                    return false;
                }
                if (value == '第')
                {
                    return false;
                }
            }

            if (!string.IsNullOrEmpty(nextChar))
            {
                char value = nextChar[0];
                if (IsAsciiIdentifierCharacter(value) || value == '.' || value == ',' || value == '-' || value == '/' || value == '\\')
                {
                    return false;
                }
                if (value == '年' || value == '月' || value == '日' || value == '时' || value == '分' || value == '秒')
                {
                    return false;
                }
            }

            return true;
        }
        private static bool TryFormatIndependentNumber(string originalText, out string formattedText)
        {
            formattedText = originalText;
            if (string.IsNullOrWhiteSpace(originalText))
            {
                return false;
            }

            string numericText = originalText;
            string openingBracket;
            string closingBracket;
            UnwrapNumericWrapper(ref numericText, out openingBracket, out closingBracket);

            bool isNegative = numericText.StartsWith("-", StringComparison.Ordinal);
            if (isNegative)
            {
                numericText = numericText.Substring(1);
            }

            string[] parts = numericText.Split('.');
            if (parts.Length > 2)
            {
                return false;
            }

            string integerPart = parts[0].Replace(",", string.Empty);
            if (!Regex.IsMatch(integerPart, @"^\d+$"))
            {
                return false;
            }

            if (parts.Length == 2 && !Regex.IsMatch(parts[1], @"^\d+$"))
            {
                return false;
            }

            string formattedIntegerPart = Regex.Replace(integerPart, @"\B(?=(\d{3})+(?!\d))", ",");
            formattedText = isNegative
                ? "-" + formattedIntegerPart
                : formattedIntegerPart;

            if (parts.Length == 2)
            {
                formattedText += "." + parts[1];
            }

            if (!string.IsNullOrEmpty(openingBracket) && !string.IsNullOrEmpty(closingBracket))
            {
                formattedText = openingBracket + formattedText + closingBracket;
            }

            return true;
        }
        private static void UnwrapNumericWrapper(ref string numericText, out string openingBracket, out string closingBracket)
        {
            openingBracket = string.Empty;
            closingBracket = string.Empty;

            if (string.IsNullOrEmpty(numericText) || numericText.Length < 2)
            {
                return;
            }

            if (numericText.StartsWith("(", StringComparison.Ordinal) && numericText.EndsWith(")", StringComparison.Ordinal))
            {
                openingBracket = "(";
                closingBracket = ")";
                numericText = numericText.Substring(1, numericText.Length - 2);
                return;
            }

            if (numericText.StartsWith("（", StringComparison.Ordinal) && numericText.EndsWith("）", StringComparison.Ordinal))
            {
                openingBracket = "（";
                closingBracket = "）";
                numericText = numericText.Substring(1, numericText.Length - 2);
            }
        }
        private static bool IsAsciiIdentifierCharacter(char value)
        {
            return (value >= '0' && value <= '9') ||
                (value >= 'A' && value <= 'Z') ||
                (value >= 'a' && value <= 'z') ||
                value == '_';
        }

        public void ConvertBrackets()
        {
            ApplyTextProcessing("中英括号转换", range =>
            {
                string text = "";
                try
                {
                    Word.Range sampleRange = range.Duplicate;
                    // 取样前部分字符以判断转换方向，避免读取大文档全量Text导致内存溢出
                    if (sampleRange.End - sampleRange.Start > 5000)
                    {
                        sampleRange.End = sampleRange.Start + 5000;
                    }
                    text = sampleRange.Text ?? "";
                }
                catch
                {
                    text = "";
                }

                int halfCount = text.Count(c => c == '(' || c == ')');
                int fullCount = text.Count(c => c == '（' || c == '）');

                // 根据选区括号比例决定转换方向，实现更智能的双向转换
                if (halfCount > 0 && halfCount >= fullCount)
                {
                    ReplaceInRange(range.Duplicate, "(", "（", "宋体");
                    ReplaceInRange(range.Duplicate, ")", "）", "宋体");
                }
                else if (fullCount > 0)
                {
                    ReplaceInRange(range.Duplicate, "（", "(", "Arial");
                    ReplaceInRange(range.Duplicate, "）", ")", "Arial");
                }
            });
        }

        private void ReplaceInRange(Word.Range range, string findText, string replaceText, string fontName = null)
        {
            Word.Find find = range.Find;
            find.ClearFormatting();
            find.Replacement.ClearFormatting();
            if (!string.IsNullOrEmpty(fontName))
            {
                find.Replacement.Font.Name = fontName;
            }
            find.Text = findText;
            find.Replacement.Text = replaceText;
            find.Format = false;
            find.MatchCase = false;
            find.MatchWholeWord = false;
            find.MatchWildcards = false;
            find.MatchSoundsLike = false;
            find.MatchAllWordForms = false;
            find.Wrap = Word.WdFindWrap.wdFindStop;

            find.Execute(Replace: Word.WdReplace.wdReplaceAll);
        }

        public void ToggleNegativeFormat()
        {
            ApplyTextProcessing("负号格式转换", range =>
            {
                string content = range.Text;
                if (string.IsNullOrEmpty(content)) return;

                // 匹配负号数字和括号数字，支持千分位和小数点
                Regex negativeRegex = new Regex(@"(-[0-9][0-9,]*(?:\.\d+)?|\([0-9][0-9,]*(?:\.\d+)?\)|（[0-9][0-9,]*(?:\.\d+)?）)");
                MatchCollection matches = negativeRegex.Matches(content);
                HashSet<string> processed = new HashSet<string>();

                bool hasNegative = matches.Cast<Match>().Any(m => m.Value.StartsWith("-"));
                bool toBrackets = hasNegative; // 默认只要有负号，就全转成括号。如果没有负号但有括号，才转成负号

                for (int index = 0; index < matches.Count; index++)
                {
                    Match match = matches[index];
                    string val = match.Value;

                    if (toBrackets && !val.StartsWith("-")) continue;
                    if (!toBrackets && val.StartsWith("-")) continue;

                    if (processed.Contains(val)) continue;
                    processed.Add(val);

                    string formattedText = "";
                    if (val.StartsWith("-"))
                    {
                        formattedText = "(" + val.Substring(1) + ")";
                    }
                    else if (val.StartsWith("(") || val.StartsWith("（"))
                    {
                        formattedText = "-" + val.Substring(1, val.Length - 2);
                    }

                    Word.Range searchRange = range.Duplicate;
                    Word.Find find = searchRange.Find;
                    find.ClearFormatting();
                    find.Text = val;
                    find.MatchWildcards = false;
                    find.Wrap = Word.WdFindWrap.wdFindStop;

                    while (find.Execute())
                    {
                        if (searchRange.InRange(range))
                        {
                            string originalAsciiFont = searchRange.Font.NameAscii;
                            searchRange.Text = formattedText;
                            if (!string.IsNullOrWhiteSpace(originalAsciiFont))
                            {
                                searchRange.Font.NameAscii = originalAsciiFont;
                            }
                            searchRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        }
                    }
                }
            });
        }
    }
}
