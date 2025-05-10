using Autodesk.Windows;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using System.IO;



namespace Civ3DECAT
{
    public class Application : IExtensionApplication
    {
        public void Initialize()
        {
            ComponentManager.ItemInitialized += ComponentManager_ItemInitialized;
        }

        public void Terminate()
        {
            ComponentManager.ItemInitialized -= ComponentManager_ItemInitialized;
        }

        private void ComponentManager_ItemInitialized(object sender, EventArgs e)
        {
            // Unsubscribe after the ribbon is initialized to avoid duplicate tabs
            ComponentManager.ItemInitialized -= ComponentManager_ItemInitialized;

            AddRibbonTab();

        }

        public void AddRibbonTab()
        {
            // Get the ribbon control
            RibbonControl ribbonControl = ComponentManager.Ribbon;

            // Create a new tab
            RibbonTab ribbonTab = new RibbonTab
            {
                Title = "Walsh",
                Id = "{810C076A-3F8F-4A5F-8E7A-18E7C07DBF39}"
            };

            ribbonControl.Tabs.Add(ribbonTab);

            // Create a panel
            RibbonPanelSource panelSource = new RibbonPanelSource
            {
                Title = "Walsh"
            };

            RibbonPanel ribbonPanel = new RibbonPanel
            {
                Source = panelSource
            };

            ribbonTab.Panels.Add(ribbonPanel);

            

            RibbonButton ribbonButton = new RibbonButton
            {
                Text = "ECAT",
                ShowText = true,
                ShowImage = true,
                Size = RibbonItemSize.Large,
                CommandHandler = new RibbonCommandHandler()
            };

            panelSource.Items.Add(ribbonButton);
        }

        public class RibbonCommandHandler : System.Windows.Input.ICommand
        {
            public event EventHandler CanExecuteChanged;

            public bool CanExecute(object parameter)
            {
                return true;
            }

            public void Execute(object parameter)
            {
                try
                {
                    // Get the current AutoCAD and Civil 3D documents
                    var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    var cdoc = CivilApplication.ActiveDocument;

                    if (cdoc == null)
                    {
                        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("No Civil 3D document is open.");
                        return;
                    }

                    // Start a transaction
                    using (var transaction = doc.TransactionManager.StartTransaction())
                    {
                        // Get the alignment names
                        var alignmentNames = cdoc.GetAlignmentIds()
                            .Cast<ObjectId>()
                            .Select(id => transaction.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Alignment)
                            .Where(alignment => alignment != null)
                            .Select(alignment => alignment.Name)
                            .ToList();

                        // Display the alignment names
                        string message = alignmentNames.Any()
                            ? "Alignments in the current document:\n" + string.Join("\n", alignmentNames)
                            : "No alignments found in the document.";

                        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(message);

                        transaction.Commit();
                    }
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog($"Error: {ex.Message}");
                }

            }
        }

    }
}
