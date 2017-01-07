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
        private string[] ColorNames = { "White", "Green", "GreenYellow", "Yellow", "Orange", "OrangeRed", "Red", "DarkRed" };
        /*private WriteableBitmap writeableBitmap;
        private Int32Rect rect;
        private int stride;
        private int bytesPerPixel;*/

        public MainWindow()
        {
            InitializeComponent();
            /*writeableBitmap = new WriteableBitmap(256, 256, 96, 96, PixelFormats.Bgra32, null);
            rect = new Int32Rect(0, 0, writeableBitmap.PixelWidth, writeableBitmap.PixelHeight);
            bytesPerPixel = (writeableBitmap.Format.BitsPerPixel + 7)/8;
            stride = writeableBitmap.PixelWidth*bytesPerPixel;*/
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
                    croppButton.IsEnabled = true;
                }
            }
        }

        private void startButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(ImagePath))
            {
                Debug.WriteLine("ImagePath is empty");
                return;
            }

            generateHeatMap(CalulateGLCM(true));
        }

        public Bitmap CropImage(Bitmap source, Rectangle section)
        {   
            Bitmap bmp = new Bitmap(section.Width, section.Height);
            Graphics g = Graphics.FromImage(bmp);
            g.DrawImage(source, 0, 0, section, GraphicsUnit.Pixel);
            return bmp;
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

        private Bitmap prepareBitmap(bool mainCalculation) {
            if (mainCalculation)
                return new Bitmap(ImagePath);
            else
            {
                Bitmap source = new Bitmap(ImagePath);
                Rectangle section = new Rectangle(new System.Drawing.Point(CropPointX, CropPointY), new System.Drawing.Size(CropLineX, CropLineY));
                
                Bitmap CroppedImage = CropImage(source, section);
                imageResult.Source = BitmapToImageSource(CroppedImage);
                return CroppedImage;
            }
        }

        private void generateHeatMap(double[,] glcmArray)
        {
            var heatMap = new Bitmap(256, 256);
            IEnumerable<double> allValues = glcmArray.Cast<double>();
            int max = Convert.ToInt32(allValues.Max());
            int pivot = max / (ColorNames.Length - 1);
            for (int i = 0; i < glcmArray.GetLength(0); i++)
            {
                for (int j = 0; j < glcmArray.GetLength(1); j++)
                {   
                    int x = Convert.ToInt32(glcmArray[i,j] / pivot);
                    string colorName = ColorNames[x];
                    heatMap.SetPixel(i, j, System.Drawing.Color.FromName(colorName));
                }
            }
            heatMapImage.Source = BitmapToImageSource(heatMap);
        }

        private double[,] CalulateGLCM(bool mainCalculation)
        {
            var image = prepareBitmap(mainCalculation);
            var unmanagedImage = UnmanagedImage.FromManagedImage(image);

            var glcm = new GrayLevelCooccurrenceMatrix
            {
                AutoGray = false
            };

            if (normalizeCheckBox.IsChecked != null) glcm.Normalize = normalizeCheckBox.IsChecked.Value;
            var comboBoxItem = degreeComboBox.SelectedItem as ComboBoxItem;
            if (comboBoxItem != null)
                glcm.Degree = (CooccurrenceDegree)Enum.Parse(typeof(CooccurrenceDegree), (string)comboBoxItem.Tag);
            glcm.Distance = int.Parse(distanceTextBox.Text);

            var results = glcm.Compute(unmanagedImage);
            var haralick = new HaralickDescriptor(results);

            entropyLabel.Content = string.Format("Entropy: {0}", haralick.Entropy.ToString("N"));
            energyLabel.Content = string.Format("Energy: {0}", haralick.AngularSecondMomentum.ToString("N5"));
            correlationLabel.Content = string.Format("Correlation: {0}", haralick.Correlation.ToString("N"));
            invDiffMomentLabel.Content = string.Format("Inv Diff Moment: {0}", haralick.InverseDifferenceMoment.ToString("N"));
            contrast.Content = string.Format("Contrast: {0}", haralick.Contrast.ToString("N"));

            if (excelCheckBox.IsChecked != null)
            {
                if(excelCheckBox.IsChecked.Value == true)
                showResultsInExcel(results);
            }
            if (glcm.Normalize)
            { //because for heatMap we need non-normalized array
                glcm.Normalize = false;
                results = glcm.Compute(unmanagedImage);
            }
            return results;
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

        private void CroppImage(object sender, RoutedEventArgs e)
        {
            CropPointX = Int32.Parse(CropPointXText.Text);
            CropPointY = Int32.Parse(CropPointYText.Text);

            CropLineX = Int32.Parse(CropLenXText.Text);
            CropLineY = Int32.Parse(CropLenYText.Text);

            var matrix = CalulateGLCM(false);
            generateHeatMap(matrix);
        }
    }
}
