using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using OfficeOpenXml;

namespace SampleScripter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void RunScriptButton_Click(object sender, RoutedEventArgs e)
        {
            // Run your script here and capture the output
            string scriptOutput = RunScript();

            // Update the outputTextBox with the script output
            outputTextBox.Text = scriptOutput;

            // Create an Excel file and save the output
            string excelFilePath = CreateExcelFile(scriptOutput);

            // Open the Excel file
            OpenExcelFile(excelFilePath);
        }

        private string RunScript()
        {
            // Run your script here and return the output as a string
            return "Script output goes here.";
        }

        private string CreateExcelFile(string scriptOutput)
        {
            // Create a new Excel package
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                // Add a new worksheet to the Excel package
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Script Output");

                // Write the script output to the worksheet
                worksheet.Cells["A1"].Value = scriptOutput;

                // Save the Excel package to a file
                string excelFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output.xlsx");
                excelPackage.SaveAs(new FileInfo(excelFilePath));

                return excelFilePath;
            }
        }

        private void OpenExcelFile(string excelFilePath)
        {
            try
            {
                // Use the default associated application to open the Excel file
                Process.Start(new ProcessStartInfo(excelFilePath)
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                // Handle any errors that occur during the process
                MessageBox.Show($"An error occurred while opening the Excel file:\n\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
