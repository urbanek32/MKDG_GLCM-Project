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

namespace GLCM_Magic
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string ImagePath { get; set; }

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

            CalulateGLCM();
        }

        private void CalulateGLCM()
        {
            var image = new Bitmap(ImagePath);
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

                for (int i = 1; i < results.GetLength(0); i++)
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

                oSheet.get_Range("A2", "IV256").Value2 = excelValues;

                oRng = oSheet.get_Range("A2", "IV256");
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
    }
}
