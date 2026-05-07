using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace FunctionBox.Features
{
    public static class TextHelper
    {
        public static readonly Regex IndependentNumberRegex = new Regex(
            @"(?:\(-?[0-9][0-9,]*(?:\.\d+)?\)|（-?[0-9][0-9,]*(?:\.\d+)?）|-?[0-9][0-9,]*(?:\.\d+)?)",
            RegexOptions.Compiled);

        public static void ApplyTextProcessing(Word.Application wordApp, string undoName, Action<Word.Range> processAction)
        {
            if (wordApp.Documents.Count == 0) return;

            Word.Document doc = wordApp.ActiveDocument;
            Word.Selection selection = wordApp.Selection;

            wordApp.UndoRecord.StartCustomRecord(undoName);
            wordApp.ScreenUpdating = false;

            // 临时关闭修订追踪，避免 Find/Replace 在 Track Changes 下异常
            bool wasTrackingRevisions = doc.TrackRevisions;
            doc.TrackRevisions = false;

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
                doc.TrackRevisions = wasTrackingRevisions;
                wordApp.ScreenUpdating = true;
                wordApp.UndoRecord.EndCustomRecord();
            }
        }
    }
}
