using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;

class Program
{
    static void Main()
    {
        string sourceFile = @"e:\files\CodeProject\VS Project\FunctionBox\FunctionBox\ThisAddIn.cs";
        string[] allLines = File.ReadAllLines(sourceFile, Encoding.UTF8);
        
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

        List<string> newLines = new List<string>();
        newLines.AddRange(allLines.Take(startIndex));
        
        // Insert proxies
        newLines.Add("        public bool SumCheckDebugModeEnabled");
        newLines.Add("        {");
        newLines.Add("            get { return FunctionBox.Features.SumCheckTool.SumCheckDebugModeEnabled; }");
        newLines.Add("            set { FunctionBox.Features.SumCheckTool.SumCheckDebugModeEnabled = value; }");
        newLines.Add("        }");
        newLines.Add("        public void ValidateSumsHorizontal()");
        newLines.Add("        {");
        newLines.Add("            FunctionBox.Features.SumCheckTool.ValidateSumsHorizontal(this.Application);");
        newLines.Add("        }");
        newLines.Add("        public void ValidateSumsVerticalTop()");
        newLines.Add("        {");
        newLines.Add("            FunctionBox.Features.SumCheckTool.ValidateSumsVerticalTop(this.Application);");
        newLines.Add("        }");
        newLines.Add("        public void ValidateSumsVerticalDown()");
        newLines.Add("        {");
        newLines.Add("            FunctionBox.Features.SumCheckTool.ValidateSumsVerticalDown(this.Application);");
        newLines.Add("        }");
        
        newLines.AddRange(allLines.Skip(endIndex));
        
        File.WriteAllLines(sourceFile, newLines, new UTF8Encoding(true));
        Console.WriteLine("Replacement complete!");
    }
}
