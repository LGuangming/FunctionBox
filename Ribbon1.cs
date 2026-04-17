using System;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Vbe.Interop;
using static FunctionBox.VbaManagerForm;

namespace FunctionBox
{
    public partial class FunctionBoxRibbon
    {
        private VbaManagerForm vbaManagerForm;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            vbaManagerForm = new VbaManagerForm(btnToolList, btnExecuteVba);

            btnCheckSumDebug.Checked = Globals.ThisAddIn.SumCheckDebugModeEnabled;
        }
        private void btnValidateHorizontal_Click(object sender, RibbonControlEventArgs e)
        {
            // 调用ThisAddIn中的方法
            Globals.ThisAddIn.ValidateSumsHorizontal();
        }
        private void btnValidateVerticalTop_Click(object sender, RibbonControlEventArgs e)
        {
            // 调用ThisAddIn中的方法
            Globals.ThisAddIn.ValidateSumsVerticalTop();
        }
        private void btnValidateVerticalDown_Click(object sender, RibbonControlEventArgs e)
        {
            // 调用ThisAddIn中的方法
            Globals.ThisAddIn.ValidateSumsVerticalDown();
        }
        private void btnClearSelectionBackground_Click(object sender, RibbonControlEventArgs e)
        {
            // 调用ThisAddIn中的方法
            Globals.ThisAddIn.ClearSelectionBackground();
        }
        private void btnClearDocumentBackground_Click(object sender, RibbonControlEventArgs e)
        {
            // 调用ThisAddIn中的方法
            Globals.ThisAddIn.ClearDocumentBackground();
        }
        private void btnToolBox_Click(object sender, RibbonControlEventArgs e)
        {
            if (vbaManagerForm == null || vbaManagerForm.IsDisposed)
            {
                vbaManagerForm = new VbaManagerForm(this.btnToolList, this.btnExecuteVba);
            }

            vbaManagerForm.Show();
            vbaManagerForm.BringToFront();
        }
        private void btnExecuteVba_Click(object sender, RibbonControlEventArgs e)
        {
            if (btnToolList.SelectedItem?.Tag is VbaCode selectedCode)
            {
                ExecuteVbaCode(selectedCode.Code);
            }
            else
            {
                MessageBox.Show("请选择一个VBA代码来执行。");
            }
        }
        public void ExecuteVbaCode(string code)
        {
            try
            {
                if (!IsValidVbaCode(code) || !TryExtractMacroName(code, out string macroName))
                {
                    MessageBox.Show("VBA代码格式无效，无法执行。请检查是否以 Sub 开始并以 End Sub 结束。");
                    return;
                }

                // 获取当前的 Word 应用程序实例
                var wordApp = Globals.ThisAddIn.Application;

                // 获取当前文档的 VBA 项目
                var vbaProject = wordApp.VBE.ActiveVBProject;
                string tempModuleName = "TempMacro_" + macroName;

                // 尝试获取名为宏名称的模块，如果不存在则创建一个新的
                VBComponent module = null;
                foreach (VBComponent vbComponent in vbaProject.VBComponents)
                {
                    if (vbComponent.Name == tempModuleName)
                    {
                        module = vbComponent;
                        break;
                    }
                }

                if (module == null)
                {
                    module = vbaProject.VBComponents.Add(Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule);
                    module.Name = tempModuleName;
                }

                // 清空模块中的现有代码
                module.CodeModule.DeleteLines(1, module.CodeModule.CountOfLines);

                // 插入代码到模块
                module.CodeModule.AddFromString(code);

                // 执行代码
                wordApp.Run(tempModuleName + "." + macroName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"执行VBA代码时出错: {ex.Message}");
            }
        }

        private void btnQuestion_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("我也不知道", "帮助");
        }

        private async void btnUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            await AlistUpdater.CheckAndUpdateAsync();
        }
        private void btnCheckSumDebug_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.SumCheckDebugModeEnabled = btnCheckSumDebug.Checked;
            MessageBox.Show(
                btnCheckSumDebug.Checked
                    ? "加总调试模式已开启。执行检查后会弹出摘要并写入调试日志。"
                    : "加总调试模式已关闭。",
                "加总调试模式");
        }

    }
}

