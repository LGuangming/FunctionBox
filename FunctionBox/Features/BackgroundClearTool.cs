using System;
using Word = Microsoft.Office.Interop.Word;

namespace FunctionBox.Features
{
    public static class BackgroundClearTool
    {
        public static void ClearSelectionBackground(Word.Application app)
        {
            Word.Selection selection = app.Selection;

            // 禁用屏幕更新
            app.ScreenUpdating = false;

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
                app.ScreenUpdating = true;
            }
        }

        public static void ClearDocumentBackground(Word.Application app)
        {
            Word.Document document = app.ActiveDocument;

            // 禁用屏幕更新
            app.ScreenUpdating = false;

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
                app.ScreenUpdating = true;
            }
        }
    }
}
