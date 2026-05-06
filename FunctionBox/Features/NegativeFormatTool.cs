using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace FunctionBox.Features
{
    public static class NegativeFormatTool
    {
        private static readonly Regex NegativeRegex = new Regex(
            @"(-[0-9][0-9,]*(?:\.\d+)?|\([0-9][0-9,]*(?:\.\d+)?\)|（[0-9][0-9,]*(?:\.\d+)?）)",
            RegexOptions.Compiled);

        public static void Execute(Word.Application app)
        {
            TextHelper.ApplyTextProcessing(app, "负号格式转换", range =>
            {
                // 按段落逐个处理，避免表格标记符干扰正则匹配
                foreach (Word.Paragraph para in range.Paragraphs)
                {
                    try
                    {
                        Word.Range paraRange = para.Range;
                        ProcessParagraph(paraRange);
                    }
                    catch { }
                }
            });
        }

        private static void ProcessParagraph(Word.Range range)
        {
            string content = "";
            try { content = range.Text; } catch { return; }
            if (string.IsNullOrEmpty(content)) return;

            MatchCollection matches = NegativeRegex.Matches(content);
            if (matches.Count == 0) return;

            HashSet<string> processed = new HashSet<string>();
            bool hasNegative = matches.Cast<Match>().Any(m => m.Value.StartsWith("-"));
            bool toBrackets = hasNegative;

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
                    if (!searchRange.InRange(range)) break;

                    if (toBrackets)
                    {
                        string originalFont = searchRange.Font.Name;
                        searchRange.Text = formattedText;
                        int count = searchRange.Characters.Count;
                        if (count >= 2)
                        {
                            try
                            {
                                if (!string.IsNullOrEmpty(originalFont))
                                    searchRange.Font.Name = originalFont;
                                searchRange.Characters[1].Font.Name = "Arial";
                                searchRange.Characters[count].Font.Name = "Arial";
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        string originalFont = null;
                        try
                        {
                            if (searchRange.Characters.Count > 2)
                                originalFont = searchRange.Characters[2].Font.Name;
                        }
                        catch { }

                        searchRange.Text = formattedText;
                        if (!string.IsNullOrEmpty(originalFont))
                        {
                            try { searchRange.Font.Name = originalFont; } catch { }
                        }
                    }

                    searchRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }
            }
        }
    }
}
