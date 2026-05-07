using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace FunctionBox.Features
{
    public static class ThousandSeparatorTool
    {
        public static void Execute(Word.Application app)
        {
            TextHelper.ApplyTextProcessing(app, "添加千分符", range =>
            {
                // 按段落逐个处理，避免表格标记符干扰正则匹配和 Word Find
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

            MatchCollection matches = TextHelper.IndependentNumberRegex.Matches(content);
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
    }
}
