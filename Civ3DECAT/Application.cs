using Autodesk.Windows;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using System.IO;
using System.Text;
using Autodesk.Civil.DatabaseServices.Styles;
using Autodesk.AutoCAD.Geometry;
using System.Linq;
using Autodesk.Civil.Settings;
using System.Reflection;

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

        private void ComponentManager_ItemInitialized(object? sender, RibbonItemEventArgs e)
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

            // Create Manhole Volumes button
            RibbonButton volumeButton = new RibbonButton
            {
                Text = "ECAT",
                ShowText = true,
                ShowImage = true,
                Size = RibbonItemSize.Large,
                CommandHandler = new ManholeVolumeCommandHandler()
            };

            panelSource.Items.Add(volumeButton);
        }

        public class ManholeVolumeCommandHandler : System.Windows.Input.ICommand
        {
            // Fix for CS9067: Add a handler implementation to suppress the warning
            // about the event never being used
            public event EventHandler? CanExecuteChanged
            {
                add { /* No need to do anything here */ }
                remove { /* No need to do anything here */ }
            }

            // Structure to hold manhole dimensions
            private class ManholeDimensions
            {
                public double BarrelInnerDiameter { get; set; } = 1.2;  // Default 1.2m
                public double BarrelWallThickness { get; set; } = 0.15; // Default 15cm
                public double CoverDiameter { get; set; } = 0.6;        // Default 60cm
                public double CoverThickness { get; set; } = 0.1;       // Default 10cm
                public double ConeHeight { get; set; } = 0.6;           // Default 60cm
                public double ConeTopDiameter { get; set; } = 0.6;      // Default 60cm
                public double BottomSlab { get; set; } = 0.25;          // Default 25cm
                public bool HasCone { get; set; } = true;
            }

            public bool CanExecute(object? parameter)
            {
                return true;
            }

            public void Execute(object? parameter)
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
                        // Get all pipe networks in the document
                        var pipeNetworkIds = cdoc.GetPipeNetworkIds();

                        if (pipeNetworkIds.Count == 0)
                        {
                            Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("No pipe networks found in the document.");
                            transaction.Commit();
                            return;
                        }

                        StringBuilder resultBuilder = new StringBuilder();
                        resultBuilder.AppendLine("MANHOLE VOLUME REPORT");
                        resultBuilder.AppendLine("=====================");
                        resultBuilder.AppendLine();

                        int totalManholes = 0;
                        double totalWallVolume = 0;
                        double totalCoverVolume = 0;
                        double totalConeVolume = 0;
                        double totalBottomSlabVolume = 0;
                        double totalVoidVolume = 0;
                        double totalConcreteVolume = 0;

                        // Process each pipe network
                        foreach (ObjectId networkId in pipeNetworkIds)
                        {
                            // Fix for CS8600: Use 'as' with null check instead of casting
                            var network = transaction.GetObject(networkId, OpenMode.ForRead) as Network;
                            if (network == null) continue;

                            resultBuilder.AppendLine($"Network: {network.Name}");
                            resultBuilder.AppendLine("------------------");

                            var structureIds = network.GetStructureIds();

                            foreach (ObjectId structureId in structureIds)
                            {
                                // Fix for CS8600: Use 'as' with null check instead of casting
                                var structure = transaction.GetObject(structureId, OpenMode.ForRead) as Structure;
                                if (structure == null) continue;

                                if (IsManholeStructure(structure, transaction))
                                {
                                    totalManholes++;

                                    // Try to get the volume directly from the 3D solid if available
                                    double solidVolume = GetStructureSolidVolume(structure, transaction);

                                    if (solidVolume > 0)
                                    {
                                        // We have a direct solid volume, use it
                                        totalConcreteVolume += solidVolume;

                                        resultBuilder.AppendLine($"Manhole: {structure.Name}");
                                        resultBuilder.AppendLine($"  Height: {(structure.RimElevation - structure.SumpElevation):F2} m");
                                        resultBuilder.AppendLine($"  Volume (from 3D solid): {solidVolume:F2} m³");
                                        resultBuilder.AppendLine();
                                    }
                                    else
                                    {
                                        // Get dimensions from the structure
                                        ManholeDimensions dimensions = GetManholeDimensions(structure, transaction);

                                        // Calculate manhole height
                                        double height = structure.RimElevation - structure.SumpElevation;
                                        double barrelHeight = height;

                                        if (dimensions.HasCone)
                                        {
                                            barrelHeight = height - dimensions.ConeHeight;
                                        }

                                        // Calculate volumes
                                        double barrelOuterRadius = dimensions.BarrelInnerDiameter / 2 + dimensions.BarrelWallThickness;
                                        double barrelInnerRadius = dimensions.BarrelInnerDiameter / 2;

                                        // Wall volume (barrel section)
                                        double wallVolume = Math.PI * (Math.Pow(barrelOuterRadius, 2) - Math.Pow(barrelInnerRadius, 2)) * barrelHeight;

                                        // Bottom slab volume
                                        double bottomSlabVolume = Math.PI * Math.Pow(barrelOuterRadius, 2) * dimensions.BottomSlab;

                                        // Cover volume
                                        double coverRadius = dimensions.CoverDiameter / 2;
                                        double coverVolume = Math.PI * Math.Pow(coverRadius, 2) * dimensions.CoverThickness;

                                        // Cone volume (if present)
                                        double coneVolume = 0;
                                        if (dimensions.HasCone)
                                        {
                                            double coneBottomRadius = barrelOuterRadius;
                                            double coneTopRadius = dimensions.ConeTopDiameter / 2;
                                            coneVolume = (1.0 / 3.0) * Math.PI * dimensions.ConeHeight *
                                                        (Math.Pow(coneBottomRadius, 2) + coneBottomRadius * coneTopRadius + Math.Pow(coneTopRadius, 2));

                                            // Subtract inner cone volume
                                            double innerConeBottomRadius = barrelInnerRadius;
                                            double innerConeTopRadius = dimensions.ConeTopDiameter / 2 - dimensions.BarrelWallThickness;
                                            if (innerConeTopRadius < 0) innerConeTopRadius = 0;

                                            double innerConeVolume = (1.0 / 3.0) * Math.PI * dimensions.ConeHeight *
                                                                   (Math.Pow(innerConeBottomRadius, 2) + innerConeBottomRadius * innerConeTopRadius + Math.Pow(innerConeTopRadius, 2));

                                            coneVolume -= innerConeVolume;
                                        }

                                        // Inner void volume
                                        double voidVolume = Math.PI * Math.Pow(barrelInnerRadius, 2) * barrelHeight;
                                        if (dimensions.HasCone)
                                        {
                                            double innerConeBottomRadius = barrelInnerRadius;
                                            double innerConeTopRadius = dimensions.ConeTopDiameter / 2 - dimensions.BarrelWallThickness;
                                            if (innerConeTopRadius < 0) innerConeTopRadius = 0;

                                            double innerConeVolume = (1.0 / 3.0) * Math.PI * dimensions.ConeHeight *
                                                                   (Math.Pow(innerConeBottomRadius, 2) + innerConeBottomRadius * innerConeTopRadius + Math.Pow(innerConeTopRadius, 2));

                                            voidVolume += innerConeVolume;
                                        }

                                        // Total concrete volume
                                        double totalConcrete = wallVolume + bottomSlabVolume + coverVolume + coneVolume;

                                        // Add to totals
                                        totalWallVolume += wallVolume;
                                        totalBottomSlabVolume += bottomSlabVolume;
                                        totalCoverVolume += coverVolume;
                                        totalConeVolume += coneVolume;
                                        totalVoidVolume += voidVolume;
                                        totalConcreteVolume += totalConcrete;

                                        // Add details to report
                                        resultBuilder.AppendLine($"Manhole: {structure.Name}");
                                        resultBuilder.AppendLine($"  Height: {height:F2} m");
                                        resultBuilder.AppendLine($"  Inner Diameter: {dimensions.BarrelInnerDiameter:F2} m");
                                        resultBuilder.AppendLine($"  Wall Thickness: {dimensions.BarrelWallThickness:F2} m");
                                        resultBuilder.AppendLine($"  Wall Volume: {wallVolume:F2} m³");
                                        resultBuilder.AppendLine($"  Bottom Slab Volume: {bottomSlabVolume:F2} m³");
                                        resultBuilder.AppendLine($"  Cover Volume: {coverVolume:F2} m³");

                                        if (dimensions.HasCone)
                                        {
                                            resultBuilder.AppendLine($"  Cone Volume: {coneVolume:F2} m³");
                                        }

                                        resultBuilder.AppendLine($"  Inner Void Volume: {voidVolume:F2} m³");
                                        resultBuilder.AppendLine($"  Total Concrete Volume: {totalConcrete:F2} m³");
                                        resultBuilder.AppendLine();
                                    }
                                }
                            }
                        }

                        // Add summary
                        resultBuilder.AppendLine("SUMMARY");
                        resultBuilder.AppendLine("-------");
                        resultBuilder.AppendLine($"Total Manholes: {totalManholes}");

                        if (totalWallVolume > 0)
                        {
                            // Only show component volumes if we calculated them
                            resultBuilder.AppendLine($"Total Wall Volume: {totalWallVolume:F2} m³");
                            resultBuilder.AppendLine($"Total Bottom Slab Volume: {totalBottomSlabVolume:F2} m³");
                            resultBuilder.AppendLine($"Total Cover Volume: {totalCoverVolume:F2} m³");
                            resultBuilder.AppendLine($"Total Cone Volume: {totalConeVolume:F2} m³");
                            resultBuilder.AppendLine($"Total Void Volume: {totalVoidVolume:F2} m³");
                        }

                        resultBuilder.AppendLine($"Total Concrete Volume: {totalConcreteVolume:F2} m³");

                        // Display the results
                        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(resultBuilder.ToString());

                        transaction.Commit();
                    }
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog($"Error: {ex.Message}");
                }
                catch (System.Exception ex)
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog($"System Error: {ex.Message}\n{ex.StackTrace}");
                }
            }

            private bool IsManholeStructure(Structure structure, Transaction transaction)
            {
                try
                {
                    // Try to determine if it's a manhole from the structure properties
                    // Fix for CS8600: Use string.Empty instead of null
                    string structureType = GetStructureProperty(structure, "PartType")?.ToLower() ?? string.Empty;
                    string structureFamily = GetStructureProperty(structure, "PartFamily")?.ToLower() ?? string.Empty;
                    string structureName = structure.Name.ToLower();

                    return structureType.Contains("manhole") ||
                           structureType.Contains("mh") ||
                           structureFamily.Contains("manhole") ||
                           structureFamily.Contains("mh") ||
                           structureName.Contains("manhole") ||
                           structureName.Contains("mh");
                }
                catch
                {
                    // Fallback to checking the name if there's an error
                    return structure.Name.ToLower().Contains("manhole") ||
                           structure.Name.ToLower().Contains("mh");
                }
            }

            private double GetStructureSolidVolume(Structure structure, Transaction transaction)
            {
                try
                {
                    // Try to get the volume property directly
                    string? volumeStr = GetStructureProperty(structure, "Volume");
                    if (!string.IsNullOrEmpty(volumeStr) && double.TryParse(volumeStr, out double volume))
                    {
                        return volume;
                    }

                    // Try to get the extents of the structure to compute an approximate volume
                    Extents3d? extents = structure.GeometricExtents;
                    if (extents.HasValue)
                    {
                        // Calculate the bounding box volume - this is rough but better than nothing
                        double width = Math.Abs(extents.Value.MaxPoint.X - extents.Value.MinPoint.X);
                        double length = Math.Abs(extents.Value.MaxPoint.Y - extents.Value.MinPoint.Y);
                        double height = Math.Abs(extents.Value.MaxPoint.Z - extents.Value.MinPoint.Z);

                        // A cylinder's volume would be approximately 78.5% of the bounding box
                        // Pi/4 ≈ 0.785
                        return 0.785 * width * length * height;
                    }

                    return 0;
                }
                catch
                {
                    return 0;
                }
            }

            private ManholeDimensions GetManholeDimensions(Structure structure, Transaction transaction)
            {
                ManholeDimensions dimensions = new ManholeDimensions();

                try
                {
                    // Try to get inner diameter
                    string? innerDiameterStr = GetStructureProperty(structure, "InnerDiameter");
                    if (!string.IsNullOrEmpty(innerDiameterStr) && double.TryParse(innerDiameterStr, out double innerDiameter))
                    {
                        dimensions.BarrelInnerDiameter = innerDiameter;
                    }
                    else
                    {
                        // Try to get from part size
                        // Fix for CS8600: Use string.Empty instead of null
                        string partSize = GetStructureProperty(structure, "PartSizeName") ?? string.Empty;
                        if (!string.IsNullOrEmpty(partSize))
                        {
                            ParseDiameterFromPartSize(partSize, ref dimensions);
                        }
                    }

                    // Try to get wall thickness
                    string? wallThicknessStr = GetStructureProperty(structure, "WallThickness");
                    if (!string.IsNullOrEmpty(wallThicknessStr) && double.TryParse(wallThicknessStr, out double wallThickness))
                    {
                        dimensions.BarrelWallThickness = wallThickness;
                    }

                    // Try to determine if it has a cone and its dimensions
                    // Fix for CS8600: Use string.Empty instead of null
                    string structureType = GetStructureProperty(structure, "PartType") ?? string.Empty;
                    dimensions.HasCone = structureType.ToLower().Contains("cone") ||
                                        structureType.ToLower().Contains("eccentric");

                    if (dimensions.HasCone)
                    {
                        string? coneHeightStr = GetStructureProperty(structure, "ConeHeight");
                        if (!string.IsNullOrEmpty(coneHeightStr) && double.TryParse(coneHeightStr, out double coneHeight))
                        {
                            dimensions.ConeHeight = coneHeight;
                        }

                        string? coneTopDiameterStr = GetStructureProperty(structure, "ConeTopDiameter");
                        if (!string.IsNullOrEmpty(coneTopDiameterStr) && double.TryParse(coneTopDiameterStr, out double coneTopDiameter))
                        {
                            dimensions.ConeTopDiameter = coneTopDiameter;
                        }
                    }

                    // Try to get bottom slab thickness
                    string? bottomThicknessStr = GetStructureProperty(structure, "BottomThickness");
                    if (!string.IsNullOrEmpty(bottomThicknessStr) && double.TryParse(bottomThicknessStr, out double bottomThickness))
                    {
                        dimensions.BottomSlab = bottomThickness;
                    }

                    // Try to get cover dimensions
                    string? frameWidthStr = GetStructureProperty(structure, "FrameWidth");
                    if (!string.IsNullOrEmpty(frameWidthStr) && double.TryParse(frameWidthStr, out double frameWidth))
                    {
                        dimensions.CoverDiameter = frameWidth;
                    }

                    string? frameHeightStr = GetStructureProperty(structure, "FrameHeight");
                    if (!string.IsNullOrEmpty(frameHeightStr) && double.TryParse(frameHeightStr, out double frameHeight))
                    {
                        dimensions.CoverThickness = frameHeight;
                    }
                }
                catch
                {
                    // If there's an error, we'll use default dimensions
                }

                return dimensions;
            }

            private string? GetStructureProperty(Structure structure, string propertyName)
            {
                try
                {
                    // Try to use reflection to access properties
                    PropertyInfo? propInfo = structure.GetType().GetProperty(propertyName);
                    if (propInfo != null)
                    {
                        object? value = propInfo.GetValue(structure);
                        return value?.ToString();
                    }

                    // Try alternative property access methods

                    // Method 1: Using extended entity data
                    try
                    {
                        using (ResultBuffer rb = structure.GetXDataForApplication("AeccStructure"))
                        {
                            if (rb != null)
                            {
                                foreach (TypedValue tv in rb)
                                {
                                    if (tv.TypeCode == 1 && tv.Value is string strValue &&
                                        strValue.StartsWith(propertyName + "="))
                                    {
                                        return strValue.Substring(propertyName.Length + 1);
                                    }
                                }
                            }
                        }
                    }
                    catch { /* Ignore errors and try next method */ }

                    // Method 2: Using dynamic cast (rarely works but worth trying)
                    try
                    {
                        dynamic dynamicStructure = structure;
                        var value = dynamicStructure[propertyName];
                        return value?.ToString();
                    }
                    catch { /* Ignore errors and try next method */ }

                    // Method 3: Check if there's a method that returns the property
                    MethodInfo? method = structure.GetType().GetMethod("Get" + propertyName);
                    if (method != null)
                    {
                        object? result = method.Invoke(structure, null);
                        return result?.ToString();
                    }
                }
                catch
                {
                    // Ignore errors
                }

                return null;
            }

            private void ParseDiameterFromPartSize(string partSize, ref ManholeDimensions dimensions)
            {
                // Parse diameter from part size (e.g., "1200mm" -> 1.2)
                if (partSize.ToLower().Contains("mm"))
                {
                    string sizeStr = partSize.ToLower().Replace("mm", "").Trim();
                    if (double.TryParse(sizeStr, out double sizeMm))
                    {
                        dimensions.BarrelInnerDiameter = sizeMm / 1000.0; // Convert mm to m
                    }
                }
                else if (partSize.ToLower().Contains("m"))
                {
                    string sizeStr = partSize.ToLower().Replace("m", "").Trim();
                    if (double.TryParse(sizeStr, out double sizeM))
                    {
                        dimensions.BarrelInnerDiameter = sizeM;
                    }
                }
                else if (double.TryParse(partSize, out double size))
                {
                    // Assume size is in millimeters if it's large number (>100)
                    if (size > 100)
                    {
                        dimensions.BarrelInnerDiameter = size / 1000.0;
                    }
                    else
                    {
                        dimensions.BarrelInnerDiameter = size;
                    }
                }
            }
        }
    }
}