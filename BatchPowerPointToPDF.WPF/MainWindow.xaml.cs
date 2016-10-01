using System;
using System.Collections.Generic;
using System.Windows;
using Microsoft.Win32;
using PowerPointToPDFLibrary;
using System.Collections.ObjectModel;

namespace BatchPowerPointToPDF.WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        PptxExporter exporter;
        public ObservableCollection<String> _pptxFilenames;

        public MainWindow()
        {
            InitializeComponent();
            exporter = new PptxExporter();
            _pptxFilenames = new ObservableCollection<string>();
            filenamesListView.ItemsSource = _pptxFilenames;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Implement more robust checking of Office Installation.
            bool officeInstalled = exporter.OfficeInstalled();

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
                    // Making sure we don't add duplicate entries.
                    if (!_pptxFilenames.Contains(file))
                    {
                        _pptxFilenames.Add(file.ToString());
                    }
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
                exporter.ConvertToPdf(file);
            }
        }

        private void removePPTXBtn_Click(object sender, RoutedEventArgs e)
        {
            if (filenamesListView.HasItems)
            {
                var list = filenamesListView.SelectedItems;
                LinkedList<String> removedItems = new LinkedList<string>();

                // Can't actually remove files here, since .NET/ C# will complain about us modifying the list while iterating through it.
                foreach (String file in list)
                {
                    removedItems.AddFirst(file);
                }

                foreach (String file in removedItems)
                {
                    _pptxFilenames.Remove(file);
                }
            }

        }

        private void openInPPT_Click(object sender, RoutedEventArgs e)
        {
            if (filenamesListView.SelectedItems != null)
            {
                foreach (String file in filenamesListView.SelectedItems)
                {
                    exporter.OpenInPowerPoint(file);
                }    
            }
                
        }
    }
}
