using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Accord.Imaging;
using Accord.Statistics;
using Microsoft.Win32;
using Excel=Microsoft.Office.Interop.Excel;
using System.Reflection;
using Rectangle = System.Drawing.Rectangle;
using Point = System.Drawing.Point;
using System.IO;
using Pen = System.Drawing.Pen;

namespace GLCM_Magic
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string ImagePath { get; set; }
        private int CropPointX { get; set; }
        private int CropPointY { get; set; }
        private int CropLineX { get; set; }
        private int CropLineY { get; set; }
        private readonly System.Drawing.Brush[] colorBrushes = {
            System.Drawing.Brushes.White,
            System.Drawing.Brushes.Green,
            System.Drawing.Brushes.GreenYellow,
            System.Drawing.Brushes.Yellow,
            System.Drawing.Brushes.Orange,
            System.Drawing.Brushes.OrangeRed,
            System.Drawing.Brushes.Red,
            System.Drawing.Brushes.DarkRed
        };
        private string Degree { get; set; }
        private int Distance { get; set; }
        private bool Normalize { get; set; }
        private bool excel { get; set; }

        /// <summary>
        /// Tuple (x, y, offsetX, offsetY)
        /// </summary>
        private Dictionary<Tuple<int, int, int, int>, double> EntropyValues { get; set; }

        public MainWindow()
        {
            InitializeComponent();
        }

        private void loadButton_Click(object sender, RoutedEventArgs e)
        {
            var op = new OpenFileDialog
            {
                Title = "Select a picture",
                Filter = "Image files (*.png;*.jpeg;*.jpg;*.bmp)|*.png;*.jpeg;*.jpg;*.bmp|All files (*.*)|*.*"

            };

            if (op.ShowDialog() == true)
            {
                imageSource.Source = new BitmapImage(new Uri(op.FileName));
                ImagePath = op.FileName;
                if (!string.IsNullOrWhiteSpace(ImagePath))
                {
                    startButton.IsEnabled = true;
                    generateHeatmapsButton.IsEnabled = true;
                }
            }
        }

        private void ReadInputParameters()
        {
            this.Normalize = false;
            this.excel = false;
            var comboBoxItem = degreeComboBox.SelectedItem as ComboBoxItem;
            this.Degree = (string)comboBoxItem.Tag;
            if (normalizeCheckBox.IsChecked.HasValue)
                Normalize = normalizeCheckBox.IsChecked.Value;
            if (excelCheckBox.IsChecked.HasValue)
                excel = excelCheckBox.IsChecked.Value;
            this.Distance = int.Parse(distanceTextBox.Text);
            //this.CropPointX = int.Parse(CropPointXText.Text);
            //this.CropPointY = int.Parse(CropPointYText.Text);
            this.CropLineX = int.Parse(CropLenXText.Text);
            this.CropLineY = int.Parse(CropLenYText.Text);
        }

        private void startButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(ImagePath))
            {
                Debug.WriteLine("ImagePath is empty");
                return;
            }
            
            if (normalizeCheckBox.IsChecked != null) 
                Normalize = normalizeCheckBox.IsChecked.Value;
            ReadInputParameters();
            //CalulateGLCM(true, this.normalize, this.degree, this.distance, true, this.excel);
            //generateHeatMap(CalulateGLCM(true, false, this.degree, this.distance, false, false));
        }

        private BitmapImage BitmapToImageSource(Bitmap bitmap)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Bmp);
                memory.Position = 0;
                BitmapImage bitmapimage = new BitmapImage();
                bitmapimage.BeginInit();
                bitmapimage.StreamSource = memory;
                bitmapimage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapimage.EndInit();

                return bitmapimage;
            }
        }

        private double[,] CalulateGLCM(Bitmap bitmap)
        {
            using (var unmanagedImage = UnmanagedImage.FromManagedImage(bitmap))
            {
                var glcm = new GrayLevelCooccurrenceMatrix
                {
                    AutoGray = false,
                    Normalize = this.Normalize,
                    Distance = this.Distance,
                    Degree = (CooccurrenceDegree) Enum.Parse(typeof(CooccurrenceDegree), Degree)
                };

                return glcm.Compute(unmanagedImage);
            }
        }

        /*private double[,] CalulateGLCM(bool fullBitmap, bool normalize, string degree, int distance, bool updateMetrics, bool excel)
        {
            var image = PrepareBitmap(fullBitmap);
            var unmanagedImage = UnmanagedImage.FromManagedImage(image);

            var glcm = new GrayLevelCooccurrenceMatrix
            {
                AutoGray = false,
                Normalize = normalize,
                Distance = distance,
                Degree = (CooccurrenceDegree)Enum.Parse(typeof(CooccurrenceDegree), degree)
            };

            var results = glcm.Compute(unmanagedImage);

            if (updateMetrics)
            {
                var haralick = new HaralickDescriptor(results);
                entropyLabel.Content = string.Format("Entropy: {0}", haralick.Entropy.ToString("N"));
                energyLabel.Content = string.Format("Energy: {0}", haralick.AngularSecondMomentum.ToString("N5"));
                correlationLabel.Content = string.Format("Correlation: {0}", haralick.Correlation.ToString("N"));
                invDiffMomentLabel.Content = string.Format("Inv Diff Moment: {0}", haralick.InverseDifferenceMoment.ToString("N"));
                contrast.Content = string.Format("Contrast: {0}", haralick.Contrast.ToString("N"));
            }

            if (excel)
                showResultsInExcel(results);
            return results;
        }*/

        private void showResultsInExcel(double[,] results)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;
            try
            {
                oXL = new Excel.Application();
                oXL.Visible = true;

                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                for (int i = 1; i < results.GetLength(0)+1; i++)
                {
                    oSheet.Cells[1, i] = i;
                    oSheet.Cells[i, 1] = i;
                }

                //bold, vertical alignment = center.
                oSheet.get_Range("A1", "IV1").Font.Bold = true;
                oSheet.get_Range("A1", "IV1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("A1", "A256").Font.Bold = true;
                oSheet.get_Range("A1", "A256").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                // Create an array to multiple values at once.
                string[,] excelValues = new string[256, 256];

                for (int i = 0; i < results.GetLength(0); i++)
                {
                    for (int j = 0; j < results.GetLength(0); j++)
                    {
                        excelValues[i, j] = results[i, j].ToString(); //TODO: Some prettier float formatting
                    }
                }

                oSheet.get_Range("B2", "IV256").Value2 = excelValues;

                oRng = oSheet.get_Range("B2", "IV256");
                oRng.EntireColumn.AutoFit();

                oXL.Visible = true;
                oXL.UserControl = true;
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);
                MessageBox.Show(errorMessage, "Error");
            }
        }

        private void generateHeatmapsButton_Click(object sender, RoutedEventArgs e)
        {
            ReadInputParameters();
            CalculateValuesForEachPartialBitmap();
            GenerateHeatmapEntropy();
            // TODO: Add other heatmaps
        }

        private void CalculateValuesForEachPartialBitmap()
        {
            using (var absentRectangleImage = (Bitmap)Bitmap.FromFile(ImagePath))
            {
                var imgWidth = absentRectangleImage.Width;
                var imgHeight = absentRectangleImage.Height;
                var stepX = CropLineX;
                var stepY = CropLineY;

                EntropyValues = new Dictionary<Tuple<int, int, int, int>, double>();

                for (var y = 0; y < imgHeight; y += stepY)
                {
                    for (var x = 0; x < imgWidth; x += stepX)
                    {
                        // clamp
                        var lenX = x + stepX;
                        if (lenX > imgWidth)
                        {
                            lenX = imgWidth;
                        }

                        var lenY = y + stepY;
                        if (lenY > imgHeight)
                        {
                            lenY = imgHeight;
                        }

                        using (var currentTile = new Bitmap(lenX, lenY))
                        {
                            currentTile.SetResolution(absentRectangleImage.HorizontalResolution, absentRectangleImage.VerticalResolution);

                            using (var currentTileGraphics = Graphics.FromImage(currentTile))
                            {
                                currentTileGraphics.Clear(System.Drawing.Color.Black);
                                var absentRectangleArea = new Rectangle(x, y, lenX, lenY);
                                currentTileGraphics.DrawImage(absentRectangleImage, 0, 0, absentRectangleArea, GraphicsUnit.Pixel);
                            }

                            var haralick = new HaralickDescriptor(CalulateGLCM(currentTile));
                            EntropyValues.Add(new Tuple<int, int, int, int>(x, y, lenX, lenY), haralick.Entropy);
                        }
                    }
                }
            }
        }

        private void GenerateHeatmapEntropy()
        {
            var imgWidth = (int)imageSource.Source.Width;
            var imgHeight = (int)imageSource.Source.Height;
            var heatMap = new Bitmap(imgWidth, imgHeight);

            var entropyValues = EntropyValues.Values.Cast<double>();
            var maxValue = entropyValues.Max();
            var pivot = maxValue / (colorBrushes.Length - 1);

            using (var gr = Graphics.FromImage(heatMap))
            {
                foreach (var entropyValue in EntropyValues)
                {
                    var brushIndex = Convert.ToInt32(entropyValue.Value / pivot);
                    gr.FillRectangle(colorBrushes[brushIndex], entropyValue.Key.Item1, entropyValue.Key.Item2, entropyValue.Key.Item3, entropyValue.Key.Item4);
                }
            }
            
            imageResult.Source = BitmapToImageSource(heatMap);
        }
    }
}
