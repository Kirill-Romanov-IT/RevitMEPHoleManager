using Autodesk.Revit.UI;
using System.Reflection;

namespace RevitMEPHoleManager
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            const string tabName = "MEP Hole Manager";
            try { application.CreateRibbonTab(tabName); } catch { }

            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Основное");

            string asmPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData btn = new PushButtonData(
                "HoleManagerBtn",
                "Запуск",
                asmPath,
                "RevitMEPHoleManager.ShowGuiCommand");

            panel.AddItem(btn);
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
