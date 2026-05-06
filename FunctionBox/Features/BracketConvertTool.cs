using System;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace FunctionBox.Features
{
    public static class BracketConvertTool
    {
        public static void Execute(Word.Application app)
        {
            TextHelper.ApplyTextProcessing(app, "中英括号转换", range =>
            {
                string text = "";
                try { text = range.Text; } catch { return; }
                if (string.IsNullOrEmpty(text)) return;

                bool hasChinese = text.IndexOf('（') >= 0 || text.IndexOf('）') >= 0;

                if (hasChinese)
                {
                    ReplaceOne(range, "（", "(", "Arial");
                    ReplaceOne(range, "）", ")", "Arial");
                }
                else
                {
                    ReplaceNonNumericBrackets(range);
                }
            });
        }

        private static void ReplaceOne(Word.Range range, string findText, string replaceText, string fontName)
        {
            Word.Range searchRange = range.Duplicate;
            searchRange.Find.ClearFormatting();
            searchRange.Find.Text = findText;
            searchRange.Find.MatchWildcards = false;
            searchRange.Find.Wrap = Word.WdFindWrap.wdFindStop;

            while (searchRange.Find.Execute())
            {
                if (!searchRange.InRange(range)) break;
                searchRange.Text = replaceText;
                try
                {
                    if (!string.IsNullOrEmpty(fontName))
                        searchRange.Font.Name = fontName;
                }
                catch { }
                searchRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }
        }

        private static void ReplaceNonNumericBrackets(Word.Range range)
        {
            // 处理左括号
            Word.Range searchRange = range.Duplicate;
            searchRange.Find.ClearFormatting();
            searchRange.Find.Text = "(";
            searchRange.Find.MatchWildcards = false;
            searchRange.Find.Wrap = Word.WdFindWrap.wdFindStop;

            while (searchRange.Find.Execute())
            {
                if (!searchRange.InRange(range)) break;
                if (IsOpenBracketOfNumber(searchRange))
                {
                    searchRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    continue;
                }
                searchRange.Text = "（";
                try { searchRange.Font.Name = "宋体"; } catch { }
                searchRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }

            // 处理右括号
            searchRange = range.Duplicate;
            searchRange.Find.ClearFormatting();
            searchRange.Find.Text = ")";
            searchRange.Find.MatchWildcards = false;
            searchRange.Find.Wrap = Word.WdFindWrap.wdFindStop;

            while (searchRange.Find.Execute())
            {
                if (!searchRange.InRange(range)) break;
                if (IsCloseBracketOfNumber(searchRange))
                {
                    searchRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    continue;
                }
                searchRange.Text = "）";
                try { searchRange.Font.Name = "宋体"; } catch { }
                searchRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }
        }

        /// <summary>
        /// 判断找到的 "(" 是否是数字括号的左括号：向后读取文本，看是否匹配 (数字) 模式
        /// </summary>
        private static bool IsOpenBracketOfNumber(Word.Range bracketRange)
        {
            try
            {
                Word.Document doc = bracketRange.Document;
                int end = Math.Min(bracketRange.Start + 30, doc.Content.End);
                Word.Range lookAhead = doc.Range(bracketRange.Start, end);
                string text = lookAhead.Text;
                if (text != null && Regex.IsMatch(text, @"^\(-?[0-9][0-9,]*(?:\.\d+)?\)"))
                    return true;
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 判断找到的 ")" 是否是数字括号的右括号：向前读取文本，看是否匹配 (数字) 模式
        /// </summary>
        private static bool IsCloseBracketOfNumber(Word.Range bracketRange)
        {
            try
            {
                Word.Document doc = bracketRange.Document;
                int start = Math.Max(bracketRange.End - 30, doc.Content.Start);
                Word.Range lookBehind = doc.Range(start, bracketRange.End);
                string text = lookBehind.Text;
                if (text != null && Regex.IsMatch(text, @"\(-?[0-9][0-9,]*(?:\.\d+)?\)$"))
                    return true;
            }
            catch { }
            return false;
        }
    }
}
