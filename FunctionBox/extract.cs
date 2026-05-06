using System;
using System.IO;
using System.Text;
using System.Linq;

class Program
{
    static void Main()
    {
        string sourceFile = @"e:\files\CodeProject\VS Project\FunctionBox\FunctionBox\ThisAddIn.cs";
        string targetFile = @"e:\files\CodeProject\VS Project\FunctionBox\FunctionBox\Features\SumCheckTool.cs";
        
        // Use default encoding which reads UTF-8 or whatever is there properly
        string[] allLines = File.ReadAllLines(sourceFile, Encoding.UTF8);
        
        // SumCheck lines are from 116 to 748. (0-indexed: 115 to 747)
        // Wait, let's find the exact indices by looking for the start and end markers
        int startIndex = -1;
        int endIndex = -1;
        
        for (int i = 0; i < allLines.Length; i++)
        {
            if (allLines[i].Contains("public bool SumCheckDebugModeEnabled"))
            {
                startIndex = i;
            }
            if (allLines[i].Contains("public int SelectionOrder { get; set; }"))
            {
                endIndex = i + 2; // Include the closing braces of the struct
                break;
            }
        }
        
        if (startIndex == -1 || endIndex == -1)
        {
            Console.WriteLine("Could not find start or end markers.");
            return;
        }

        var sumCheckLines = allLines.Skip(startIndex).Take(endIndex - startIndex + 1).ToList();
        
        // We need to change "private void BeginSumCheckDebug", etc., to "public" if we want to call them, but wait!
        // The Execute method logic. SumCheckTool has 3 entry points: ValidateSumsHorizontal, ValidateSumsVerticalTop, ValidateSumsVerticalDown.
        // And they all need access to Application.
        // We will make them static and pass Application app.
        
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("using System;");
        sb.AppendLine("using System.Collections.Generic;");
        sb.AppendLine("using System.Globalization;");
        sb.AppendLine("using System.IO;");
        sb.AppendLine("using System.Linq;");
        sb.AppendLine("using Word = Microsoft.Office.Interop.Word;");
        sb.AppendLine("using System.Windows.Forms;");
        sb.AppendLine("using System.Text.RegularExpressions;");
        sb.AppendLine("");
        sb.AppendLine("namespace FunctionBox.Features");
        sb.AppendLine("{");
        sb.AppendLine("    public static class SumCheckTool");
        sb.AppendLine("    {");
        
        // The rest of the methods
        foreach (string line in sumCheckLines)
        {
            string modifiedLine = line;
            // Add static to everything
            if (modifiedLine.Contains("public bool SumCheckDebugModeEnabled")) modifiedLine = modifiedLine.Replace("public bool", "public static bool");
            if (modifiedLine.Contains("public void ValidateSums")) modifiedLine = modifiedLine.Replace("public void", "public static void");
            if (modifiedLine.Contains("private bool TryGetSelectedCells")) modifiedLine = modifiedLine.Replace("private bool", "private static bool");
            if (modifiedLine.Contains("private List<SelectedCellInfo> CollectSelectedCellsByToggleMarker")) modifiedLine = modifiedLine.Replace("private List", "private static List");
            if (modifiedLine.Contains("private void ValidateAsSingleSequence")) modifiedLine = modifiedLine.Replace("private void", "private static void");
            if (modifiedLine.Contains("private void ValidateCellGroup")) modifiedLine = modifiedLine.Replace("private void", "private static void");
            if (modifiedLine.Contains("private List<string> BeginSumCheckDebug")) modifiedLine = modifiedLine.Replace("private List", "private static List");
            if (modifiedLine.Contains("private void EndSumCheckDebug")) modifiedLine = modifiedLine.Replace("private void", "private static void");

            // Change GetWordSelection() to app.Selection and this.Application to app
            if (modifiedLine.Contains("Word.Application wordApp = this.Application;"))
            {
                // We will add 'Word.Application app' to the parameter list of ValidateSums...
                continue; 
            }
            if (modifiedLine.Contains("Word.Selection selection = GetWordSelection();"))
            {
                modifiedLine = "            Word.Selection selection = app.Selection;";
            }
            if (modifiedLine.Contains("public static void ValidateSumsHorizontal()")) modifiedLine = "        public static void ValidateSumsHorizontal(Word.Application app)";
            if (modifiedLine.Contains("public static void ValidateSumsVerticalTop()")) modifiedLine = "        public static void ValidateSumsVerticalTop(Word.Application app)";
            if (modifiedLine.Contains("public static void ValidateSumsVerticalDown()")) modifiedLine = "        public static void ValidateSumsVerticalDown(Word.Application app)";

            // Also we need to fix SumCheckDebugLogPath and SumTolerance, which are defined in ThisAddIn.cs
            
            sb.AppendLine("        " + modifiedLine);
        }
        
        sb.AppendLine("    }");
        sb.AppendLine("}");

        // Prepend SumTolerance and SumCheckDebugLogPath
        string finalCode = sb.ToString();
        string consts = @"
        private static bool sumCheckDebugModeEnabled;
        private static readonly string SumCheckDebugLogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            ""FunctionBox"",
            ""sum-check-debug.log"");
        private const double SumTolerance = 0.001d;
";
        finalCode = finalCode.Insert(finalCode.IndexOf("    {", finalCode.IndexOf("public static class SumCheckTool")) + 5, consts);
        // Clean up any occurrences of `wordApp.` to `app.`
        finalCode = finalCode.Replace("wordApp.ScreenUpdating", "app.ScreenUpdating");

        File.WriteAllText(targetFile, finalCode, new UTF8Encoding(true)); // UTF-8 with BOM
        Console.WriteLine("Extraction complete!");
    }
}
