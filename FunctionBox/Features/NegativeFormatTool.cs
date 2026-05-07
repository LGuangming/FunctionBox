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
            @"((?<![0-9a-zA-Z])-[0-9][0-9,]*(?:\.\d+)?|\([0-9][0-9,]*(?:\.\d+)?\)|（[0-9][0-9,]*(?:\.\d+)?）)",
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

                    // 跳过序号格式的括号数字，如 "(1) 文本内容"
                    // 判断条件：括号数字后面紧跟空格、字母或中文字符
                    if (!toBrackets && IsSequenceNumber(searchRange))
                    {
                        searchRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        continue;
                    }

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
        /// <summary>
        /// 判断括号数字是否是序号（如 "(1) 文本" 或段落开头的 "(1)"）。
        /// 判断条件：不含小数点和千分位的纯整数括号，且后面紧跟空格、字母或中文字符。
        /// </summary>
        private static bool IsSequenceNumber(Word.Range bracketRange)
        {
            try
            {
                string text = bracketRange.Text;
                if (string.IsNullOrEmpty(text)) return false;

                // 只检查括号格式 (数字) 或 （数字）
                if (!(text.StartsWith("(") || text.StartsWith("（"))) return false;

                // 如果包含小数点或千分位逗号，肯定是数字不是序号
                if (text.Contains(".") || text.Contains(",")) return false;

                // 检查后面的字符
                Word.Document doc = bracketRange.Document;
                if (bracketRange.End < doc.Content.End)
                {
                    Word.Range nextRange = doc.Range(bracketRange.End, Math.Min(bracketRange.End + 1, doc.Content.End));
                    string nextChar = nextRange.Text;
                    if (!string.IsNullOrEmpty(nextChar))
                    {
                        char c = nextChar[0];
                        // 后面是空格、字母、中文字符 → 序号
                        if (c == ' ' || c == '\t' || c == '.' ||
                            (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') ||
                            c > 255) // 中文等非ASCII字符
                        {
                            return true;
                        }
                    }
                }

                // 检查前面的字符：如果在段落/行首，也可能是序号
                if (bracketRange.Start > doc.Content.Start)
                {
                    Word.Range prevRange = doc.Range(bracketRange.Start - 1, bracketRange.Start);
                    string prevChar = prevRange.Text;
                    if (!string.IsNullOrEmpty(prevChar))
                    {
                        char c = prevChar[0];
                        // 前面是换行/段落标记 → 段落开头的序号
                        if (c == '\r' || c == '\n' || c == '\a')
                        {
                            return true;
                        }
                    }
                }
                else
                {
                    // 在文档最开头 → 序号
                    return true;
                }
            }
            catch { }
            return false;
        }
    }
}
