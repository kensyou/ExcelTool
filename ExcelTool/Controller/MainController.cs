using ExcelDna.Integration;
using ExcelTool.Forms;
using ExcelTool.Helper;
using ExcelTool.Modules;
using ExcelTool.Properties;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelTool.Controller
{
    public class MainController : IDisposable
    {
        private readonly IExcelHelper _Helper;
        private readonly IZenrinModule _ZenrinModule;
        private readonly IIndustrialModule _IndustrialModule;
        public MainController(SampleController sample
            , WpfInteractionController wpfInteraction
            , IExcelHelper helper
            , IZenrinModule zenrin
            , IIndustrialModule industrial)
        {
            Sample = sample;
            WpfInteraction = wpfInteraction;
            _Helper = helper;
            _ZenrinModule = zenrin;
            _IndustrialModule = industrial;
        }

        public SampleController Sample { get; private set; }

        public WpfInteractionController WpfInteraction { get; private set; }

        public void Dispose()
        {
            Sample.Dispose();
            WpfInteraction.Dispose();
        }
        private async Task ExecuteAndCatch(Func<Task> func)
        {
            try
            {
                _Helper.ValidateWorkbookNotInProtectedView();
                await func();
            }
            catch (Exception ex)
            {
                DisplayMessageForAsync(ex.Message, Resources.Error, isError: true);
            }
        }
        private void DisplayMessageForAsync(string msg, string caption, bool isError = false)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                if (isError)
                {
                    var msgBox = new MessageBoxWithScroll(msg, caption, "Error: Please check detail below:", true);
                    msgBox.ShowDialog();
                    //MessageBox.Show(msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show(msg, caption);
                }
            });
        }
        public async Task ImportInterchangeData()
        {
            await ExecuteAndCatch(_ZenrinModule.ImportInterchangeData);
        }

        public async Task ExportInterchangeAsMergeSql()
        {
            await ExecuteAndCatch(_ZenrinModule.ExportInterchangeDataAsMergeSql);
        }

        public async Task ParseIndustrialStackingList()
        {
            await ExecuteAndCatch(_IndustrialModule.ParseIndustrialStackingList);
        }
    }
}
