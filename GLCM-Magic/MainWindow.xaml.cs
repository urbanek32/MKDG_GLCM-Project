﻿using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
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
using System.Collections;

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
            System.Drawing.Brushes.Blue,
            System.Drawing.Brushes.LightSkyBlue,
            System.Drawing.Brushes.Green,
            System.Drawing.Brushes.GreenYellow,
            System.Drawing.Brushes.Khaki,
            System.Drawing.Brushes.Yellow,
            System.Drawing.Brushes.Orange,
            System.Drawing.Brushes.OrangeRed
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
        private OrderedDictionary EntropyValues { get; set; }
        private OrderedDictionary EnergyValues { get; set; }
        private OrderedDictionary CorrelationValues { get; set; }
        private OrderedDictionary InvDiffMomentValues { get; set; }
        private OrderedDictionary ContrastValues { get; set; }

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
                var image = new BitmapImage(new Uri(op.FileName));
                imageSource.Source = image;
                SourceImagePath = op.FileName;
                SourceImageHeight = image.PixelHeight;
                SourceImageWidth = image.PixelWidth;
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

        private void showGlcmInExcel(double [,]glcm, string sheetName)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;
            string len = (glcm.GetLength(0) + 1).ToString();
            try
            {
                oXL = new Excel.Application();
                oXL.Visible = true;

                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                oSheet.Name = sheetName;

                for (int i = 1; i < glcm.GetLength(0) + 1; i++)
                {
                    oSheet.Cells[1, i] = i;
                    oSheet.Cells[i, 1] = i;
                }

                //bold, vertical alignment = center.
                oSheet.get_Range("A1", "IV1").Font.Bold = true;
                oSheet.get_Range("A1", "IV1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("A1", "A" + len).Font.Bold = true;
                oSheet.get_Range("A1", "A" + len).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                // Create an array to multiple values at once.
                double[,] excelValues = new double[glcm.GetLength(0), glcm.GetLength(0)];

                for (int i = 0; i < glcm.GetLength(0); i++)
                {
                    for (int j = 0; j < glcm.GetLength(0); j++)
                    {
                        excelValues[i, j] = glcm[i, j]; //TODO: Some prettier float formatting
                    }
                }

                oSheet.get_Range("B2", "IV" + len).Value2 = excelValues;

                oRng = oSheet.get_Range("B2", "IV" + len);
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
            if (ExportToExcel)
            {
                showGlcmInExcel(CalulateGLCM((Bitmap)Bitmap.FromFile(SourceImagePath)),"GLCM");
            }
        }

        private void GenerateHeatmap(OrderedDictionary dict, System.Windows.Controls.Image imageControl)
        {
            var heatMap = new Bitmap((int)SourceImageWidth, (int)SourceImageHeight);

            var entropyValues = dict.Values.Cast<double>();
            var maxValue = entropyValues.Max();
            var pivot = maxValue / (colorBrushes.Length - 1);

            using (var gr = Graphics.FromImage(heatMap))
            {
                var enumerator = dict.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    var enumKey = enumerator.Key as Tuple<int, int, int, int>;
                    var enumValue = enumerator.Value as double?;
                    var brushIndex = Convert.ToInt32(enumValue / pivot);
                    gr.FillRectangle(colorBrushes[brushIndex], enumKey.Item1, enumKey.Item2, enumKey.Item3, enumKey.Item4);
                }
            }

            InvokeAction(() =>
            {
                imageControl.Source = BitmapToImageSource(heatMap);
            });
        }

        private void OnGlcmBackgroundDoWork(object sender, DoWorkEventArgs e)
        {
            using (var sourceBitmap = MakeGrayscale3((Bitmap)Bitmap.FromFile(SourceImagePath)))
            {
                var imgWidth = sourceBitmap.Width;
                var imgHeight = sourceBitmap.Height;
                var stepX = CropLineX;
                var stepY = CropLineY;

                var iterations = Math.Ceiling((float)imgWidth / stepX) * Math.Ceiling((float)imgHeight / stepY);
                var iterationCounter = 1;

                // Calculate GLCM for entire image
                UpdateMetrics(new HaralickDescriptor(CalulateGLCM(sourceBitmap)));

                EntropyValues = new OrderedDictionary();
                EnergyValues = new OrderedDictionary();
                CorrelationValues = new OrderedDictionary();
                InvDiffMomentValues = new OrderedDictionary();
                ContrastValues = new OrderedDictionary();

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

                        using (var currentTile = new Bitmap(stepX, stepY))
                        {
                            currentTile.SetResolution(sourceBitmap.HorizontalResolution, sourceBitmap.VerticalResolution);

                            using (var currentTileGraphics = Graphics.FromImage(currentTile))
                            {
                                currentTileGraphics.Clear(System.Drawing.Color.Black);
                                var absentRectangleArea = new Rectangle(x, y, lenX, lenY);
                                currentTileGraphics.DrawImage(sourceBitmap, 0, 0, absentRectangleArea, GraphicsUnit.Pixel);

                                var haralick = new HaralickDescriptor(CalulateGLCM(currentTile));
                                EntropyValues.Add(new Tuple<int, int, int, int>(x, y, lenX, lenY), haralick.Entropy);
                                EnergyValues.Add(new Tuple<int, int, int, int>(x, y, lenX, lenY), haralick.AngularSecondMomentum);
                                CorrelationValues.Add(new Tuple<int, int, int, int>(x, y, lenX, lenY), haralick.Correlation);
                                InvDiffMomentValues.Add(new Tuple<int, int, int, int>(x, y, lenX, lenY), haralick.InverseDifferenceMoment);
                                ContrastValues.Add(new Tuple<int, int, int, int>(x, y, lenX, lenY), haralick.Contrast);
                            }
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
                MessageBox.Show("BackgroundWorker operation failed: \n{e.Error}", "Operation Failed", MessageBoxButton.OK, MessageBoxImage.Error);
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
            
            _heatmapsBackgroundWorker.ReportProgress(20);
            metricsToDouble(EntropyValues, "Entropy", ExportToExcel);
            InvokeAction(() =>
            {
                StatusTextBlock.Text = "Generating Energy heatmap...";
            });
            GenerateHeatmap(EnergyValues, EnergyImageResult);
            metricsToDouble(EnergyValues, "Energy", ExportToExcel);
            _heatmapsBackgroundWorker.ReportProgress(40);

            InvokeAction(() =>
            {
                StatusTextBlock.Text = "Generating Correlation heatmap...";
            });
            GenerateHeatmap(CorrelationValues, CorrelationImageResult);
            metricsToDouble(CorrelationValues, "Correlation", ExportToExcel);
            _heatmapsBackgroundWorker.ReportProgress(60);

            InvokeAction(() =>
            {
                StatusTextBlock.Text = "Generating Inv Dif fMoment heatmap...";
            });
            GenerateHeatmap(InvDiffMomentValues, InvDiffMomentImageResult);
            metricsToDouble(InvDiffMomentValues, "InvDiffMoment", ExportToExcel);
            _heatmapsBackgroundWorker.ReportProgress(80);

            InvokeAction(() =>
            {
                StatusTextBlock.Text = "Generating Contrast heatmap...";
            });
            GenerateHeatmap(ContrastValues, ContrastImageResult);
            metricsToDouble(ContrastValues, "Contrast", ExportToExcel);
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
                MessageBox.Show("BackgroundWorker operation failed: \n{e.Error}", "Operation Failed", MessageBoxButton.OK, MessageBoxImage.Error);
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

        private void metricsToDouble(OrderedDictionary dict, string metricName, bool export)
        {
            if (!export) 
                return;
            int len = Convert.ToInt32(Math.Sqrt(dict.Count));
            int i =0, j = 0, c = 0;
            double[][] metricsArray = new double[len][];
            for (int k = 0; k < len; k++)
            {
                metricsArray[k] = new double[len];
            }
            foreach (DictionaryEntry de in dict)
            {
                if (j == len) { i++; j=0; }
                metricsArray[i][j] = Convert.ToDouble(de.Value);
                j++;
                c++;
                if (c == len * len)
                {
                    break;
                }
            }
            showGlcmInExcel(To2D(metricsArray), metricName);
        }

        static T[,] To2D<T>(T[][] source)
        {
            try
            {
                int FirstDim = source.Length;
                int SecondDim = source.GroupBy(row => row.Length).Single().Key; // throws InvalidOperationException if source is not rectangular

                var result = new T[FirstDim, SecondDim];
                for (int i = 0; i < FirstDim; ++i)
                    for (int j = 0; j < SecondDim; ++j)
                        result[i, j] = source[i][j];

                return result;
            }
            catch (InvalidOperationException)
            {
                throw new InvalidOperationException("The given jagged array is not rectangular.");
            }
        }

        public static Bitmap MakeGrayscale3(Bitmap original)
        {
            //create a blank bitmap the same size as original
            var newBitmap = new Bitmap(original.Width, original.Height);

            //get a graphics object from the new image
            var g = Graphics.FromImage(newBitmap);

            //create the grayscale ColorMatrix
            var colorMatrix = new ColorMatrix(
               new[]
               {
                 new float[] {.3f, .3f, .3f, 0, 0},
                 new float[] {.59f, .59f, .59f, 0, 0},
                 new float[] {.11f, .11f, .11f, 0, 0},
                 new float[] {0, 0, 0, 1, 0},
                 new float[] {0, 0, 0, 0, 1}
               });

            //create some image attributes
            var attributes = new ImageAttributes();

            //set the color matrix attribute
            attributes.SetColorMatrix(colorMatrix);

            //draw the original image on the new image
            //using the grayscale color matrix
            g.DrawImage(original, new Rectangle(0, 0, original.Width, original.Height), 0, 0, original.Width, original.Height, GraphicsUnit.Pixel, attributes);

            //dispose the Graphics object
            g.Dispose();
            return newBitmap;
        }
    }
}
