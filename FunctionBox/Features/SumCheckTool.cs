using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace FunctionBox.Features
{
    public static class SumCheckTool
    {
        private static bool sumCheckDebugModeEnabled;
        private static readonly string SumCheckDebugLogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "FunctionBox",
            "sum-check-debug.log");
        private const double SumTolerance = 0.001d;

                public static bool SumCheckDebugModeEnabled
                {
                    get { return sumCheckDebugModeEnabled; }
                    set { sumCheckDebugModeEnabled = value; }
                }
                public static void ValidateSumsHorizontal(Word.Application app)
                {
                    Word.Selection selection = app.Selection;
                    List<string> debugLines = BeginSumCheckDebug("横向加总检查", selection);
        
                    // 禁用屏幕更新
                    app.ScreenUpdating = false;
        
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
                        app.ScreenUpdating = true;
                        EndSumCheckDebug(debugLines);
                    }
                }
                public static void ValidateSumsVerticalTop(Word.Application app)
                {
                    Word.Selection selection = app.Selection;
                    List<string> debugLines = BeginSumCheckDebug("竖向加总检查-自下向上", selection);
        
                    // 禁用屏幕更新
                    app.ScreenUpdating = false;
        
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
                        app.ScreenUpdating = true;
                        EndSumCheckDebug(debugLines);
                    }
                }
                public static void ValidateSumsVerticalDown(Word.Application app)
                {
                    Word.Selection selection = app.Selection;
                    List<string> debugLines = BeginSumCheckDebug("竖向加总检查-自上向下", selection);
        
                    // 禁用屏幕更新
                    app.ScreenUpdating = false;
        
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
                        app.ScreenUpdating = true;
                        EndSumCheckDebug(debugLines);
                    }
                }
                private static bool TryGetSelectedCells(
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
                private static List<SelectedCellInfo> CollectSelectedCellsByToggleMarker(Word.Selection selection, List<string> debugLines)
                {
                    List<SelectedCellInfo> result = new List<SelectedCellInfo>();
                    if (selection == null || selection.Range == null || selection.Range.Tables.Count == 0)
                    {
                        return result;
                    }
                    if (selection.Type == Word.WdSelectionType.wdSelectionIP)
                    {
                        // 仅仅是光标闪烁，不存在非连续多选
                        return result;
                    }
        
                    Word.Table table = selection.Range.Tables[1];
                    Dictionary<string, int> originalBoldValues = new Dictionary<string, int>();
                    bool formatChanged = false;
                    Word.UndoRecord undoRecord = null;
        
                    try
                    {
                        // 开启合并撤销记录
                        undoRecord = selection.Application.UndoRecord;
                        undoRecord.StartCustomRecord("FunctionBoxSumCheckMarker");
        
                        // 1. 记录初始状态
                        foreach (Word.Cell cell in table.Range.Cells)
                        {
                            try { originalBoldValues[GetCellCoordinateKey(cell.RowIndex, cell.ColumnIndex)] = cell.Range.Bold; }
                            catch { }
                        }
        
                        HashSet<string> foundKeys = new HashSet<string>();
        
                        // 2. 第一次原生地翻转选区加粗状态 (原生命令会作用于所有非连续区块)
                        selection.Application.CommandBars.ExecuteMso("Bold");
        
                        // 3. 收集第一次变化的单元格
                        foreach (Word.Cell cell in table.Range.Cells)
                        {
                            try
                            {
                                string key = GetCellCoordinateKey(cell.RowIndex, cell.ColumnIndex);
                                if (originalBoldValues.TryGetValue(key, out int orig) && cell.Range.Bold != orig)
                                {
                                    formatChanged = true;
                                    if (foundKeys.Add(key))
                                    {
                                        result.Add(new SelectedCellInfo { Cell = cell, RowIndex = cell.RowIndex, ColumnIndex = cell.ColumnIndex });
                                    }
                                }
                            }
                            catch { }
                        }
        
                        // 4. 第二次原生地翻转选区加粗状态 (确保原先就是混排粗体的格子也能发生变化)
                        selection.Application.CommandBars.ExecuteMso("Bold");
        
                        // 5. 收集第二次变化的单元格
                        foreach (Word.Cell cell in table.Range.Cells)
                        {
                            try
                            {
                                string key = GetCellCoordinateKey(cell.RowIndex, cell.ColumnIndex);
                                if (!foundKeys.Contains(key))
                                {
                                    if (originalBoldValues.TryGetValue(key, out int orig) && cell.Range.Bold != orig)
                                    {
                                        formatChanged = true;
                                        if (foundKeys.Add(key))
                                        {
                                            result.Add(new SelectedCellInfo { Cell = cell, RowIndex = cell.RowIndex, ColumnIndex = cell.ColumnIndex });
                                        }
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    catch (Exception ex)
                    {
                        AppendSumCheckDebug(debugLines, "非连续选区反查失败: " + ex.Message);
                    }
                    finally
                    {
                        if (undoRecord != null)
                        {
                            try { undoRecord.EndCustomRecord(); } catch { }
                        }
        
                        // 只有当确定发生过格式变化时，才进行撤销，防止误撤销用户的正常输入
                        if (formatChanged)
                        {
                            try { selection.Application.ActiveDocument.Undo(); } catch { }
                        }
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
                private static void ValidateAsSingleSequence(List<SelectedCellInfo> cells, bool isLastCellAsTarget, string context, List<string> debugLines)
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
                private static void ValidateCellGroup(List<SelectedCellInfo> cells, bool isLastCellAsTarget, string context, List<string> debugLines)
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
                private static List<string> BeginSumCheckDebug(string modeName, Word.Selection selection)
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
                private static void EndSumCheckDebug(List<string> debugLines)
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
    }
}
