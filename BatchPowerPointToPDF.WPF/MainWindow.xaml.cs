using System;
using System.Collections.Generic;
using System.Windows;
using Microsoft.Win32;
using PowerPointToPDFLibrary;

namespace BatchPowerPointToPDF.WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        LinkedList<String> _pptxFilenames;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Implement more robust checking of Office Installation.
            bool officeInstalled = PptxExporter.OfficeInstalled();

            officeInstalledLabel.Content += officeInstalled.ToString().ToUpper();

            if (!officeInstalled)
            {
                MessageBox.Show("Office is not installed. In order to continue, please install Office.");
                Application.Current.Shutdown();
            }
        }

        private void openPPTXBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenPdf();
        }

        /// <summary>
        /// Opens Windows dialog box and allows user to pick PowerPoint presentations (PPTX) that are to converted to PDFs.
        /// </summary>
        private void OpenPdf()
        {
            _pptxFilenames = new LinkedList<String>();

            // Initializing Dialog Box
            OpenFileDialog openPptxDialog = new OpenFileDialog
            {
                InitialDirectory = "%documents%",
                Filter = "PowerPoint Presentations (*.PPTX)|*.PPTX",
                Multiselect = true,
                Title = "Select PowerPoint presentation(s)"
            };

            if (openPptxDialog.ShowDialog() ?? false)
            {
                foreach (String file in openPptxDialog.FileNames)
                {
                    _pptxFilenames.AddFirst(file.ToString());
                }
            }
        }

        private void button_Copy_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Fix blocking of UI thread, implement ConvertToPDF as async function.
            // When implemented as a Task, sometimes the Task is not properly disposed, and therefore
            // PowerPoint is stuck and does not convert the file. 

            foreach (String file in _pptxFilenames)
            {
                PptxExporter.ConvertToPdf(file);
            }
        }
    }
}
