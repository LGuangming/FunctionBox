using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
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

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            vbaTrustWarningShown = LoadVbaTrustWarningState();

            // 添加事件处理程序-删除临时VBA
            this.Application.DocumentBeforeClose += new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(Application_DocumentBeforeClose);
            SyncVbaShortcutBindings();
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

                    if (!isRectangular)
                    {
                        ValidateAsSingleSequence(pair.Value, isLastCellAsTarget: true, "非连续序列兜底", debugLines);
                        continue;
                    }

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

                    // 非连续跨单元格兜底：按单元格顺序校验“前面数据之和 == 最后一项”
                    if (validatedGroups == 0 && pair.Value.Count >= 2)
                    {
                        var orderedCells = pair.Value
                            .OrderBy(cell => cell.RowIndex)
                            .ThenBy(cell => cell.ColumnIndex)
                            .ToList();
                        ValidateCellGroup(orderedCells, isLastCellAsTarget: true, "按行分组无有效组兜底", debugLines);
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

                    if (!isRectangular)
                    {
                        ValidateAsSingleSequence(pair.Value, isLastCellAsTarget: false, "非连续序列兜底", debugLines);
                        continue;
                    }

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

                    // 非连续跨单元格兜底：按单元格顺序校验“前面数据之和 == 最后一项”
                    if (validatedGroups == 0 && pair.Value.Count >= 2)
                    {
                        var orderedCells = pair.Value
                            .OrderBy(cell => cell.RowIndex)
                            .ThenBy(cell => cell.ColumnIndex)
                            .ToList();
                        ValidateCellGroup(orderedCells, isLastCellAsTarget: false, "按列分组无有效组兜底", debugLines);
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

                    if (!isRectangular)
                    {
                        ValidateAsSingleSequence(pair.Value, isLastCellAsTarget: true, "非连续序列兜底", debugLines);
                        continue;
                    }

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

                    // 非连续跨单元格兜底：按单元格顺序校验“前面数据之和 == 最后一项”
                    if (validatedGroups == 0 && pair.Value.Count >= 2)
                    {
                        var orderedCells = pair.Value
                            .OrderBy(cell => cell.RowIndex)
                            .ThenBy(cell => cell.ColumnIndex)
                            .ToList();
                        ValidateCellGroup(orderedCells, isLastCellAsTarget: true, "按列分组无有效组兜底", debugLines);
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

            // Word 在非连续表格选区时可能只暴露最后一段，这里做一次兜底反查
            if (selectedCells.Count <= 1)
            {
                List<SelectedCellInfo> fallbackCells = CollectSelectedCellsByToggleMarker(selection, debugLines);
                if (fallbackCells.Count > selectedCells.Count)
                {
                    selectedCells = fallbackCells;
                    AppendSumCheckDebug(debugLines, $"已切换为反查结果，共 {selectedCells.Count} 个单元格");
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
        private List<SelectedCellInfo> CollectSelectedCellsByToggleMarker(Word.Selection selection, List<string> debugLines)
        {
            List<SelectedCellInfo> result = new List<SelectedCellInfo>();

            if (selection == null || selection.Range == null || selection.Range.Tables.Count == 0)
            {
                return result;
            }

            Word.Table table = selection.Range.Tables[1];
            Dictionary<string, int> originalBoldValues = new Dictionary<string, int>();

            try
            {
                for (int rowIndex = 1; rowIndex <= table.Rows.Count; rowIndex++)
                {
                    Word.Row row = table.Rows[rowIndex];
                    for (int columnIndex = 1; columnIndex <= row.Cells.Count; columnIndex++)
                    {
                        Word.Cell cell = row.Cells[columnIndex];
                        originalBoldValues[GetCellCoordinateKey(cell.RowIndex, cell.ColumnIndex)] = cell.Range.Bold;
                    }
                }

                selection.Font.Bold = (int)Word.WdConstants.wdToggle;

                for (int rowIndex = 1; rowIndex <= table.Rows.Count; rowIndex++)
                {
                    Word.Row row = table.Rows[rowIndex];
                    for (int columnIndex = 1; columnIndex <= row.Cells.Count; columnIndex++)
                    {
                        Word.Cell cell = row.Cells[columnIndex];
                        string key = GetCellCoordinateKey(cell.RowIndex, cell.ColumnIndex);

                        int originalBold;
                        if (!originalBoldValues.TryGetValue(key, out originalBold))
                        {
                            continue;
                        }

                        if (cell.Range.Bold != originalBold)
                        {
                            result.Add(new SelectedCellInfo
                            {
                                Cell = cell,
                                RowIndex = cell.RowIndex,
                                ColumnIndex = cell.ColumnIndex
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AppendSumCheckDebug(debugLines, "非连续选区反查失败: " + ex.Message);
            }
            finally
            {
                try
                {
                    for (int rowIndex = 1; rowIndex <= table.Rows.Count; rowIndex++)
                    {
                        Word.Row row = table.Rows[rowIndex];
                        for (int columnIndex = 1; columnIndex <= row.Cells.Count; columnIndex++)
                        {
                            Word.Cell cell = row.Cells[columnIndex];
                            string key = GetCellCoordinateKey(cell.RowIndex, cell.ColumnIndex);

                            int originalBold;
                            if (originalBoldValues.TryGetValue(key, out originalBold))
                            {
                                cell.Range.Bold = originalBold;
                            }
                        }
                    }
                }
                catch
                {
                    // 恢复失败不阻断主流程
                }
            }

            result = result
                .OrderBy(cell => cell.RowIndex)
                .ThenBy(cell => cell.ColumnIndex)
                .ToList();

            AppendSumCheckDebug(debugLines, $"非连续选区反查得到 {result.Count} 个单元格");
            return result;
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
            targetCell.Shading.BackgroundPatternColor = isMatch
                ? Word.WdColor.wdColorLightGreen
                : Word.WdColor.wdColorYellow;

            AppendSumCheckDebug(
                debugLines,
                $"{context}: 顺序单元格={FormatCellSequence(orderedCells)}, 目标={FormatCellIdentity(orderedCells[targetIndex])}, 求和={sum}, 目标值={targetValue}, 结果={(isMatch ? "通过" : "失败")}");
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

            targetCell.Shading.BackgroundPatternColor = isMatch
                ? Word.WdColor.wdColorLightGreen
                : Word.WdColor.wdColorYellow;

            AppendSumCheckDebug(
                debugLines,
                $"{context}: 单元格={FormatCellSequence(cells)}, 目标={FormatCellIdentity(cells[targetIndex])}, 求和={sum}, 目标值={targetValue}, 结果={(isMatch ? "通过" : "失败")}");
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

                string block = string.Join(Environment.NewLine, debugLines) +
                    Environment.NewLine + new string('-', 60) + Environment.NewLine;
                File.AppendAllText(SumCheckDebugLogPath, block);
            }
            catch
            {
                // 调试日志失败不影响主流程
            }

            string preview = string.Join(Environment.NewLine, debugLines.Take(12));
            MessageBox.Show(
                "加总调试输出已生成。\n日志文件: " + SumCheckDebugLogPath + "\n\n" + preview,
                "加总调试模式");
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
            if (cell == null)
            {
                return "null";
            }

            return $"R{cell.RowIndex}C{cell.ColumnIndex}(Order={cell.SelectionOrder},Text={GetCellText(cell.Cell)})";
        }
        private static string FormatCellSequence(List<SelectedCellInfo> cells)
        {
            if (cells == null || cells.Count == 0)
            {
                return "[]";
            }

            return "[" + string.Join(", ", cells.Select(FormatCellIdentity)) + "]";
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
                // 清除选中区域的背景色和高亮色
                Word.Range selectedRange = selection.Range;
                selectedRange.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
                selectedRange.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
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

                // 清除整个文档的背景色
                entireDocumentRange.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;

                // 清除整个文档的高亮色
                entireDocumentRange.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;

                // 清除表格的背景色
                foreach (Word.Table table in document.Tables)
                {
                    table.Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
                    table.Range.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
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
                MessageBox.Show(warningMessage);
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
    }
}










