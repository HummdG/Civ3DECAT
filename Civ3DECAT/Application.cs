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
using System.Collections.Generic;

// Add alias directives to resolve ambiguous references
using WinForms = System.Windows.Forms;
using DrawingFont = System.Drawing.Font;
using DrawingColor = System.Drawing.Color;
using DrawingSize = System.Drawing.Size;

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
                Text = "Manhole Volumes",
                ShowText = true,
                ShowImage = true,
                Size = RibbonItemSize.Large,
                CommandHandler = new ManholeVolumeCommandHandler()
            };

            panelSource.Items.Add(volumeButton);
        }

        // Model class for manhole volume data
        public class ManholeVolumeData
        {
            public string Name { get; set; } = string.Empty;
            public string Network { get; set; } = string.Empty;
            public double Height { get; set; }
            public double InnerDiameter { get; set; }
            public double WallThickness { get; set; }
            public double WallVolume { get; set; }
            public double BottomSlabVolume { get; set; }
            public double CoverVolume { get; set; }
            public double ConeVolume { get; set; }
            public double VoidVolume { get; set; }
            public double TotalConcreteVolume { get; set; }
            public bool IsTotal { get; set; } = false;
        }

        // Windows Forms dialog for displaying manhole volume report
        public class ManholeVolumeReportForm : WinForms.Form
        {
            private WinForms.DataGridView dataGridView;
            private WinForms.Label totalLabel;

            public ManholeVolumeReportForm()
            {
                Text = "Manhole Volume Report";
                Size = new DrawingSize(1000, 600);
                StartPosition = WinForms.FormStartPosition.CenterScreen;

                // Create layout
                var mainLayout = new WinForms.TableLayoutPanel
                {
                    Dock = WinForms.DockStyle.Fill,
                    RowCount = 2,
                    ColumnCount = 1
                };
                mainLayout.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 90F));
                mainLayout.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 10F));
                Controls.Add(mainLayout);

                // Create data grid
                dataGridView = new WinForms.DataGridView
                {
                    Dock = WinForms.DockStyle.Fill,
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    ReadOnly = true,
                    AutoSizeColumnsMode = WinForms.DataGridViewAutoSizeColumnsMode.Fill,
                    AlternatingRowsDefaultCellStyle = new WinForms.DataGridViewCellStyle { BackColor = DrawingColor.AliceBlue }
                };

                // Add columns
                dataGridView.Columns.Add("Name", "Manhole");
                dataGridView.Columns.Add("Network", "Network");
                dataGridView.Columns.Add("Height", "Height (m)");
                /*dataGridView.Columns.Add("InnerDiameter", "Inner Diameter (m)");
                dataGridView.Columns.Add("WallThickness", "Wall Thickness (m)");
                dataGridView.Columns.Add("WallVolume", "Wall Volume (m³)");
                dataGridView.Columns.Add("BottomSlabVolume", "Bottom Slab Volume (m³)");
                dataGridView.Columns.Add("CoverVolume", "Cover Volume (m³)");
                dataGridView.Columns.Add("ConeVolume", "Cone Volume (m³)");
                dataGridView.Columns.Add("VoidVolume", "Void Volume (m³)");*/
                dataGridView.Columns.Add("TotalConcreteVolume", "Total Concrete (m³)");

                mainLayout.Controls.Add(dataGridView, 0, 0);

                // Create bottom panel
                var bottomPanel = new WinForms.FlowLayoutPanel
                {
                    Dock = WinForms.DockStyle.Fill,
                    FlowDirection = WinForms.FlowDirection.RightToLeft
                };

                var closeButton = new WinForms.Button
                {
                    Text = "Close",
                    Padding = new WinForms.Padding(10, 5, 10, 5),
                    Margin = new WinForms.Padding(10)
                };
                closeButton.Click += (s, e) => Close();

                totalLabel = new WinForms.Label
                {
                    Font = new DrawingFont(Font, System.Drawing.FontStyle.Bold),
                    Margin = new WinForms.Padding(10),
                    TextAlign = System.Drawing.ContentAlignment.MiddleRight,
                    AutoSize = true
                };

                bottomPanel.Controls.Add(closeButton);
                bottomPanel.Controls.Add(totalLabel);
                mainLayout.Controls.Add(bottomPanel, 0, 1);
            }

            public void AddManholeData(ManholeVolumeData data)
            {
                int rowIndex = dataGridView.Rows.Add();
                var row = dataGridView.Rows[rowIndex];

                row.Cells["Name"].Value = data.Name;
                row.Cells["Network"].Value = data.Network;
                row.Cells["Height"].Value = data.Height.ToString("F2");
                /* row.Cells["InnerDiameter"].Value = data.InnerDiameter.ToString("F2");
                row.Cells["WallThickness"].Value = data.WallThickness.ToString("F2");
                row.Cells["WallVolume"].Value = data.WallVolume.ToString("F2");
                row.Cells["BottomSlabVolume"].Value = data.BottomSlabVolume.ToString("F2");
                row.Cells["CoverVolume"].Value = data.CoverVolume.ToString("F2");
                row.Cells["ConeVolume"].Value = data.ConeVolume.ToString("F2");
                row.Cells["VoidVolume"].Value = data.VoidVolume.ToString("F2"); */
                row.Cells["TotalConcreteVolume"].Value = data.TotalConcreteVolume.ToString("F2");

                if (data.IsTotal)
                {
                    row.DefaultCellStyle.Font = new DrawingFont(dataGridView.Font, System.Drawing.FontStyle.Bold);
                    row.DefaultCellStyle.BackColor = DrawingColor.LightGray;
                }

                UpdateTotalLabel();
            }

            private void UpdateTotalLabel()
            {
                double total = 0;

                foreach (WinForms.DataGridViewRow row in dataGridView.Rows)
                {
                    if (row.DefaultCellStyle.Font != null && row.DefaultCellStyle.Font.Bold)
                    {
                        string valueStr = row.Cells["TotalConcreteVolume"].Value?.ToString() ?? "0";
                        if (double.TryParse(valueStr, out double value))
                        {
                            total += value;
                        }
                    }
                }

                totalLabel.Text = $"Total Concrete Volume: {total:F2} m³";
            }
        }

        public class ManholeVolumeCommandHandler : System.Windows.Input.ICommand
        {
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
                    var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    var cdoc = CivilApplication.ActiveDocument;

                    if (cdoc == null)
                    {
                        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("No Civil 3D document is open.");
                        return;
                    }

                    // Create Windows Forms dialog to display results
                    var reportForm = new ManholeVolumeReportForm();

                    using (var transaction = doc.TransactionManager.StartTransaction())
                    {
                        var pipeNetworkIds = cdoc.GetPipeNetworkIds();

                        if (pipeNetworkIds.Count == 0)
                        {
                            Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("No pipe networks found in the document.");
                            transaction.Commit();
                            return;
                        }

                        foreach (ObjectId networkId in pipeNetworkIds)
                        {
                            var network = transaction.GetObject(networkId, OpenMode.ForRead) as Network;
                            if (network == null) continue;

                            var structureIds = network.GetStructureIds();

                            int networkManholeCount = 0;
                            double networkWallVolume = 0;
                            double networkBottomSlabVolume = 0;
                            double networkCoverVolume = 0;
                            double networkConeVolume = 0;
                            double networkVoidVolume = 0;
                            double networkConcreteVolume = 0;

                            foreach (ObjectId structureId in structureIds)
                            {
                                var structure = transaction.GetObject(structureId, OpenMode.ForRead) as Structure;
                                if (structure == null) continue;

                                if (IsManholeStructure(structure, transaction))
                                {
                                    networkManholeCount++;

                                    var manholeData = new ManholeVolumeData
                                    {
                                        Name = structure.Name,
                                        Network = network.Name,
                                        Height = structure.RimElevation - structure.SumpElevation
                                    };

                                    double solidVolume = GetStructureSolidVolume(structure, transaction);

                                    if (solidVolume > 0)
                                    {
                                        manholeData.TotalConcreteVolume = solidVolume;
                                        networkConcreteVolume += solidVolume;
                                        reportForm.AddManholeData(manholeData);
                                    }
                                    else
                                    {
                                        ManholeDimensions dimensions = GetManholeDimensions(structure, transaction);

                                        double height = structure.RimElevation - structure.SumpElevation;
                                        double barrelHeight = height;

                                        if (dimensions.HasCone)
                                        {
                                            barrelHeight = height - dimensions.ConeHeight;
                                        }

                                        manholeData.InnerDiameter = dimensions.BarrelInnerDiameter;
                                        manholeData.WallThickness = dimensions.BarrelWallThickness;

                                        double barrelOuterRadius = dimensions.BarrelInnerDiameter / 2 + dimensions.BarrelWallThickness;
                                        double barrelInnerRadius = dimensions.BarrelInnerDiameter / 2;

                                        double wallVolume = Math.PI * (Math.Pow(barrelOuterRadius, 2) - Math.Pow(barrelInnerRadius, 2)) * barrelHeight;
                                        manholeData.WallVolume = wallVolume;
                                        networkWallVolume += wallVolume;

                                        double bottomSlabVolume = Math.PI * Math.Pow(barrelOuterRadius, 2) * dimensions.BottomSlab;
                                        manholeData.BottomSlabVolume = bottomSlabVolume;
                                        networkBottomSlabVolume += bottomSlabVolume;

                                        double coverRadius = dimensions.CoverDiameter / 2;
                                        double coverVolume = Math.PI * Math.Pow(coverRadius, 2) * dimensions.CoverThickness;
                                        manholeData.CoverVolume = coverVolume;
                                        networkCoverVolume += coverVolume;

                                        double coneVolume = 0;
                                        if (dimensions.HasCone)
                                        {
                                            double coneBottomRadius = barrelOuterRadius;
                                            double coneTopRadius = dimensions.ConeTopDiameter / 2;
                                            coneVolume = (1.0 / 3.0) * Math.PI * dimensions.ConeHeight *
                                                        (Math.Pow(coneBottomRadius, 2) + coneBottomRadius * coneTopRadius + Math.Pow(coneTopRadius, 2));

                                            double innerConeBottomRadius = barrelInnerRadius;
                                            double innerConeTopRadius = dimensions.ConeTopDiameter / 2 - dimensions.BarrelWallThickness;
                                            if (innerConeTopRadius < 0) innerConeTopRadius = 0;

                                            double innerConeVolume = (1.0 / 3.0) * Math.PI * dimensions.ConeHeight *
                                                                   (Math.Pow(innerConeBottomRadius, 2) + innerConeBottomRadius * innerConeTopRadius + Math.Pow(innerConeTopRadius, 2));

                                            coneVolume -= innerConeVolume;
                                        }
                                        manholeData.ConeVolume = coneVolume;
                                        networkConeVolume += coneVolume;

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
                                        manholeData.VoidVolume = voidVolume;
                                        networkVoidVolume += voidVolume;

                                        double totalConcrete = wallVolume + bottomSlabVolume + coverVolume + coneVolume;
                                        manholeData.TotalConcreteVolume = totalConcrete;
                                        networkConcreteVolume += totalConcrete;

                                        reportForm.AddManholeData(manholeData);
                                    }
                                }
                            }

                            if (networkManholeCount > 0)
                            {
                                reportForm.AddManholeData(new ManholeVolumeData
                                {
                                    Name = $"TOTAL ({networkManholeCount})",
                                    Network = network.Name,
                                    WallVolume = networkWallVolume,
                                    BottomSlabVolume = networkBottomSlabVolume,
                                    CoverVolume = networkCoverVolume,
                                    ConeVolume = networkConeVolume,
                                    VoidVolume = networkVoidVolume,
                                    TotalConcreteVolume = networkConcreteVolume,
                                    IsTotal = true
                                });
                            }
                        }

                        transaction.Commit();
                    }

                    // Show the Windows Forms dialog
                    WinForms.Application.EnableVisualStyles();
                    reportForm.ShowDialog();

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