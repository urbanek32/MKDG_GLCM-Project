using System;
using System.Collections.Generic;
using System.ComponentModel;
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
using System.Threading;
using System.Windows.Threading;
using Pen = System.Drawing.Pen;

namespace GLCM_Magic
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string SourceImagePath { get; set; }
        private double SourceImageWidth { get; set; }
        private double SourceImageHeight { get; set; }
        private int CropLineX { get; set; }
        private int CropLineY { get; set; }
        private readonly System.Drawing.Brush[] colorBrushes = {
            System.Drawing.Brushes.LightGreen,
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
        private bool ExportToExcel { get; set; }

        private readonly BackgroundWorker _glcmBackgroundWorker;
        private readonly BackgroundWorker _heatmapsBackgroundWorker;

        /// <summary>
        /// Tuple (x, y, width, height)
        /// </summary>
        private Dictionary<Tuple<int, int, int, int>, double> EntropyValues { get; set; }
        private Dictionary<Tuple<int, int, int, int>, double> EnergyValues { get; set; }
        private Dictionary<Tuple<int, int, int, int>, double> CorrelationValues { get; set; }

        public MainWindow()
        {
            InitializeComponent();

            _glcmBackgroundWorker = new BackgroundWorker
            {
                WorkerReportsProgress = true
            };
            _glcmBackgroundWorker.DoWork += OnGlcmBackgroundDoWork;
            _glcmBackgroundWorker.ProgressChanged += OnGlcmBackgroundProgressChanged;
            _glcmBackgroundWorker.RunWorkerCompleted += OnGlcmBackgroundWorkerCompleted;

            _heatmapsBackgroundWorker = new BackgroundWorker
            {
                WorkerReportsProgress = true
            };
            _heatmapsBackgroundWorker.DoWork += OnHeatmapsBackgroundDoWork;
            _heatmapsBackgroundWorker.ProgressChanged += OnGlcmBackgroundProgressChanged;
            _heatmapsBackgroundWorker.RunWorkerCompleted += OnHeatmapsBackgroundWorkerCompleted;
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
                SourceImagePath = op.FileName;
                SourceImageHeight = imageSource.Source.Height;
                SourceImageWidth = imageSource.Source.Width;
                if (!string.IsNullOrWhiteSpace(SourceImagePath))
                {
                    GenerateHeatmapsButton.IsEnabled = true;
                    StatusTextBlock.Text = "Ready to start";
                }
            }
        }

        private void ReadInputParameters()
        {
            this.Normalize = false;
            this.ExportToExcel = false;
            var comboBoxItem = degreeComboBox.SelectedItem as ComboBoxItem;
            this.Degree = (string)comboBoxItem.Tag;
            if (normalizeCheckBox.IsChecked.HasValue)
                Normalize = normalizeCheckBox.IsChecked.Value;
            if (excelCheckBox.IsChecked.HasValue)
                ExportToExcel = excelCheckBox.IsChecked.Value;
            this.Distance = int.Parse(distanceTextBox.Text);
            this.CropLineX = int.Parse(CropLenXText.Text);
            this.CropLineY = int.Parse(CropLenYText.Text);
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

        private void UpdateMetrics(HaralickDescriptor haralick)
        {
            InvokeAction(() =>
            {
                entropyLabel.Content = string.Format("Entropy: {0}", haralick.Entropy.ToString("N"));
                energyLabel.Content = string.Format("Energy: {0}", haralick.AngularSecondMomentum.ToString("N5"));
                correlationLabel.Content = string.Format("Correlation: {0}", haralick.Correlation.ToString("N"));
                invDiffMomentLabel.Content = string.Format("Inv Diff Moment: {0}", haralick.InverseDifferenceMoment.ToString("N"));
                contrast.Content = string.Format("Contrast: {0}", haralick.Contrast.ToString("N"));
            });
        }

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
            GenerateHeatmapsButton.IsEnabled = false;
            StatusTextBlock.Text = "Calculating GLCM matrix...";
            _glcmBackgroundWorker.RunWorkerAsync();
        }

        private void GenerateHeatmap(Dictionary<Tuple<int, int, int, int>, double> dict, System.Windows.Controls.Image imageControl)
        {
            var heatMap = new Bitmap((int)SourceImageWidth, (int)SourceImageHeight);

            var entropyValues = dict.Values.Cast<double>();
            var maxValue = entropyValues.Max();
            var pivot = maxValue / (colorBrushes.Length - 1);

            using (var gr = Graphics.FromImage(heatMap))
            {
                foreach (var entropyValue in dict)
                {
                    var brushIndex = Convert.ToInt32(entropyValue.Value / pivot);
                    gr.FillRectangle(colorBrushes[brushIndex], entropyValue.Key.Item1, entropyValue.Key.Item2, entropyValue.Key.Item3, entropyValue.Key.Item4);
                }
            }

            InvokeAction(() =>
            {
                imageControl.Source = BitmapToImageSource(heatMap);
            });
        }

        private void OnGlcmBackgroundDoWork(object sender, DoWorkEventArgs e)
        {
            using (var sourceBitmap = (Bitmap)Bitmap.FromFile(SourceImagePath))
            {
                var imgWidth = sourceBitmap.Width;
                var imgHeight = sourceBitmap.Height;
                var stepX = CropLineX;
                var stepY = CropLineY;

                var iterations = Math.Ceiling((float)imgWidth / stepX) * Math.Ceiling((float)imgHeight / stepY);
                var iterationCounter = 1;

                // Calculate GLCM for entire image
                UpdateMetrics(new HaralickDescriptor(CalulateGLCM(sourceBitmap)));

                EntropyValues = new Dictionary<Tuple<int, int, int, int>, double>();
                EnergyValues = new Dictionary<Tuple<int, int, int, int>, double>();
                CorrelationValues = new Dictionary<Tuple<int, int, int, int>, double>();

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
                            currentTile.SetResolution(sourceBitmap.HorizontalResolution, sourceBitmap.VerticalResolution);

                            using (var currentTileGraphics = Graphics.FromImage(currentTile))
                            {
                                currentTileGraphics.Clear(System.Drawing.Color.Black);
                                var absentRectangleArea = new Rectangle(x, y, lenX, lenY);
                                currentTileGraphics.DrawImage(sourceBitmap, 0, 0, absentRectangleArea, GraphicsUnit.Pixel);
                            }

                            var haralick = new HaralickDescriptor(CalulateGLCM(currentTile));
                            EntropyValues.Add(new Tuple<int, int, int, int>(x, y, lenX, lenY), haralick.Entropy);
                            EnergyValues.Add(new Tuple<int, int, int, int>(x, y, lenX, lenY), haralick.AngularSecondMomentum);
                            CorrelationValues.Add(new Tuple<int, int, int, int>(x, y, lenX, lenY), haralick.Correlation);
                        }
                        
                        _glcmBackgroundWorker.ReportProgress((int)(100 / (iterations) * iterationCounter++));
                    }
                }
            }
        }

        private void OnGlcmBackgroundProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // BGW facilitates dealing with UI-owned objects by executing this handler on the main thread.
            ProgressBar.Value = e.ProgressPercentage;
        }

        private void OnGlcmBackgroundWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MessageBox.Show("BackgroundWorker was cancelled.", "Operation Cancelled", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            else if (e.Error != null)
            {
                MessageBox.Show($"BackgroundWorker operation failed: \n{e.Error}", "Operation Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                _heatmapsBackgroundWorker.RunWorkerAsync();
                return;
            }

            ResetBackgroundWorker();
        }

        private void OnHeatmapsBackgroundDoWork(object sender, DoWorkEventArgs e)
        {
            InvokeAction(() =>
            {
                StatusTextBlock.Text = "Generating Entropy heatmap...";
            });
            GenerateHeatmap(EntropyValues, EntropyImageResult);
            _heatmapsBackgroundWorker.ReportProgress(33);

            InvokeAction(() =>
            {
                StatusTextBlock.Text = "Generating Energy heatmap...";
            });
            GenerateHeatmap(EnergyValues, EnergyImageResult);
            _heatmapsBackgroundWorker.ReportProgress(66);

            InvokeAction(() =>
            {
                StatusTextBlock.Text = "Generating Correlation heatmap...";
            });
            GenerateHeatmap(CorrelationValues, CorrelationImageResult);
            _heatmapsBackgroundWorker.ReportProgress(100);
        }

        private void InvokeAction(Action action)
        {
            if (Dispatcher.CheckAccess())
            {
                action.Invoke();
            }
            else
            {
                Dispatcher.Invoke(DispatcherPriority.Background, action);
            }
        }

        private void OnHeatmapsBackgroundWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MessageBox.Show("BackgroundWorker was cancelled.", "Operation Cancelled", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            else if (e.Error != null)
            {
                MessageBox.Show($"BackgroundWorker operation failed: \n{e.Error}", "Operation Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                StatusTextBlock.Text = "Heatmaps are ready";
            }

            ResetBackgroundWorker();
        }

        private void ResetBackgroundWorker()
        {
            ProgressBar.Value = 0;
            GenerateHeatmapsButton.IsEnabled = true;
        }
    }
}
