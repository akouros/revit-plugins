using System.Reflection;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Media.Imaging;
using Autodesk.Revit.Attributes;

namespace H2M
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class H2MPluginR23 : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            string tabName = "H2M";
            application.CreateRibbonTab(tabName);

            RibbonPanel ribbonPanel = application.CreateRibbonPanel(tabName, "Company Tools");

            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            //Renumber Views on Sheet Button
            PushButtonData ViewNumberingButtonData = new PushButtonData("cmdRenumberViews", "Renumber\r\nViews", thisAssemblyPath, "H2M.ViewNumberingCommand");
            PushButton ViewNumberingPushButton = ribbonPanel.AddItem(ViewNumberingButtonData) as PushButton;
            ViewNumberingPushButton.ToolTip = "Renumber all views placed on a sheet in the project per H2M standard sequential order";

            // Load renumber views embedded image
            BitmapImage ViewNumberingBitmapImage = LoadEmbeddedImage("H2M.icons.ViewNumberingIcon.png");
            ViewNumberingPushButton.LargeImage = ViewNumberingBitmapImage;

            //Sheet Count Button
            PushButtonData SheetCountButtonData = new PushButtonData("cmdSheetCount", "Sheet\r\nCount", thisAssemblyPath, "H2M.SheetCountCommand");
            PushButton SheetCountPushButton = ribbonPanel.AddItem(SheetCountButtonData) as PushButton;
            SheetCountPushButton.ToolTip = "Renumber all sheet count numbers and total sheet count in the project from Excel";

            // Load sheet count embedded image
            BitmapImage SheetCountBitmapImage = LoadEmbeddedImage("H2M.icons.SheetCountIcon.png");
            SheetCountPushButton.LargeImage = SheetCountBitmapImage;

            //Spell Check Button
            PushButtonData SpellCheckButtonData = new PushButtonData("cmdSpellCheck", "Spell\r\nCheck", thisAssemblyPath, "H2M.SpellCheckCommand");
            PushButton SpellCheckPushButton = ribbonPanel.AddItem(SpellCheckButtonData) as PushButton;
            SpellCheckPushButton.ToolTip = "Spell check all text on every sheet in the current project";

            // Load spell check embedded image
            BitmapImage SpellCheckBitmapImage = LoadEmbeddedImage("H2M.icons.SpellCheckIcon.png");
            SpellCheckPushButton.LargeImage = SpellCheckBitmapImage;

            return Result.Succeeded;
        }
        //Load jpg icon method for each button
        private BitmapImage LoadEmbeddedImage(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    // Handle the case where the resource is not found
                    return null;
                }

                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.StreamSource = stream;
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                bitmap.Freeze(); // For thread safety
                return bitmap;
            }
        }
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}