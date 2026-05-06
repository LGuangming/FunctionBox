using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.Vbe.Interop;
using Newtonsoft.Json;

namespace FunctionBox
{
    public partial class ThisAddIn
    {
        private bool vbaTrustWarningShown;

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
        private static readonly Regex IndependentNumberRegex = new Regex(
            @"(?:\(-?[0-9][0-9,]*(?:\.\d+)?\)|（-?[0-9][0-9,]*(?:\.\d+)?）|-?[0-9][0-9,]*(?:\.\d+)?)",
            RegexOptions.Compiled);

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            vbaTrustWarningShown = LoadVbaTrustWarningState();

            // 自动尝试在注册表中开启 VBA 信任选项
            EnableVbaTrustAutomatically();

            // 添加事件处理程序-删除临时VBA
            this.Application.DocumentBeforeClose += new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(Application_DocumentBeforeClose);
            SyncVbaShortcutBindings();
        }

        public bool EnsureVbaTrustEnabled()
        {
            try
            {
                string version = this.Application.Version;
                string keyPath = @"Software\Microsoft\Office\" + version + @"\Word\Security";
                
                using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(keyPath, true))
                {
                    if (key != null)
                    {
                        object value = key.GetValue("AccessVBOM");
                        if (value == null || (int)value != 1)
                        {
                            key.SetValue("AccessVBOM", 1, Microsoft.Win32.RegistryValueKind.DWord);
                        }
                        return true;
                    }
                    else
                    {
                        using (Microsoft.Win32.RegistryKey newKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(keyPath))
                        {
                            if (newKey != null)
                            {
                                newKey.SetValue("AccessVBOM", 1, Microsoft.Win32.RegistryValueKind.DWord);
                                return true;
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                // 静默失败
            }
            return false;
        }
        private void EnableVbaTrustAutomatically()
        {
            EnsureVbaTrustEnabled();
        }
        public bool HandleVbaTrustErrorSilently(Exception ex)
        {
            if (!IsVbaTrustError(ex))
            {
                return false;
            }

            EnsureVbaTrustEnabled();
            vbaTrustWarningShown = true;
            SaveVbaTrustWarningState();
            return true;
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
            get { return FunctionBox.Features.SumCheckTool.SumCheckDebugModeEnabled; }
            set { FunctionBox.Features.SumCheckTool.SumCheckDebugModeEnabled = value; }
        }
        public void ValidateSumsHorizontal()
        {
            FunctionBox.Features.SumCheckTool.ValidateSumsHorizontal(this.Application);
        }
        public void ValidateSumsVerticalTop()
        {
            FunctionBox.Features.SumCheckTool.ValidateSumsVerticalTop(this.Application);
        }
        public void ValidateSumsVerticalDown()
        {
            FunctionBox.Features.SumCheckTool.ValidateSumsVerticalDown(this.Application);
        }
        public void ClearSelectionBackground()
        {
            FunctionBox.Features.BackgroundClearTool.ClearSelectionBackground(this.Application);
        }

        public void ClearDocumentBackground()
        {
            FunctionBox.Features.BackgroundClearTool.ClearDocumentBackground(this.Application);
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
            return HandleVbaTrustErrorSilently(ex);
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
        
        public void AddThousandSeparator()
        {
            FunctionBox.Features.ThousandSeparatorTool.Execute(this.Application);
        }

        public void ConvertBrackets()
        {
            FunctionBox.Features.BracketConvertTool.Execute(this.Application);
        }

        public void ToggleNegativeFormat()
        {
            FunctionBox.Features.NegativeFormatTool.Execute(this.Application);
        }

        public void DecimalAlign()
        {
            FunctionBox.Features.DecimalAlignTool.Execute(this.Application);
        }
    }
}
