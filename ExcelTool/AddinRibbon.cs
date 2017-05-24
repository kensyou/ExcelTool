using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AddinX.Ribbon.Contract;
using AddinX.Ribbon.Contract.Command;
using AddinX.Ribbon.ExcelDna;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelTool
{
    [ComVisible(true)]
    public class AddinRibbon : RibbonFluent
    {
        protected override void CreateFluentRibbon(IRibbonBuilder build)
        {
            build.CustomUi.Ribbon.Tabs(tab =>
            {
                tab.AddTab("Ken").SetId("KenTab")
                    .Groups(g =>
                    {
                        g.AddGroup("ゼンリン").SetId("ZenrinGroup").Items(d =>
                        {
                            d.AddButton("Import IC")
                            .SetId("btnImportIC")
                            .LargeSize().ImageMso("Repeat");
                            d.AddButton("Map")
                            .SetId("btnMap")
                            .LargeSize().ImageMso("Repeat");
                            d.AddButton("Export SQL")
                            .SetId("btnZenrinICExportSql")
                            .LargeSize().ImageMso("Export");
                        });
                        g.AddGroup("Industrial").SetId("IndustrialGroup").Items(d =>
                        {
                            d.AddButton("Parse Stacking")
                            .SetId("btnParseIndustrialStackingList")
                            .LargeSize().ImageMso("ResolveConflictOrError");
                        });
                    });
            });

        }

        protected override void CreateRibbonCommand(IRibbonCommands cmds)
        {
            cmds.AddButtonCommand("btnImportIC")
                .IsEnabled(() => AddinContext.ExcelApp.Worksheets.Count() > 2)
                .Action(() => AddinContext.MainController.ImportInterchangeData().Wait());
            cmds.AddButtonCommand("btnMap")
                .IsEnabled(() => AddinContext.ExcelApp.Worksheets.Any()).IsVisible(() => true)
                .Action(() => AddinContext.MainController.Sample.OpenForm());
            cmds.AddButtonCommand("btnZenrinICExportSql")
                .IsEnabled(() => AddinContext.ExcelApp.Worksheets.Any()).IsVisible(() => true)
                .Action(() => AddinContext.MainController.ExportInterchangeAsMergeSql().Wait());
            cmds.AddButtonCommand("btnParseIndustrialStackingList")
                .IsEnabled(() => AddinContext.ExcelApp.Worksheets.Any()).IsVisible(() => true)
                .Action(() => AddinContext.MainController.ParseIndustrialStackingList().Wait());
            //cmds.AddButtonCommand("TestCmd")
            //    .Action(() => AddinContext.MainController.Sample.ShowMessage());

            //cmds.AddButtonCommand("TableCmd")
            //    .Action(() => AddinContext.MainController.Report.CreateTable());
        }

        public override void OnClosing()
        {
            AddinContext.TokenCancellationSource.Cancel();

            AddinContext.ExcelApp.DisposeChildInstances(true);
            AddinContext.ExcelApp = null;

            AddinContext.Container.Dispose();
            AddinContext.Container = null;
        }

        public override void OnOpening()
        {
            // Register to events
            //AddinContext.ExcelApp.SheetSelectionChangeEvent += (a, e) => RefreshRibbon();
            AddinContext.ExcelApp.SheetActivateEvent += (e) => RefreshRibbon();
            AddinContext.ExcelApp.SheetChangeEvent += (a, e) => RefreshRibbon();
        }

        private void RefreshRibbon()
        {
            Ribbon?.Invalidate();
        }
    }
}
