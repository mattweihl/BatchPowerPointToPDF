using System;
using System.Collections.Generic;
using System.Windows;
using Microsoft.Win32;
using PowerPointToPDFLibrary;
using System.Collections.ObjectModel;
using System.Windows.Controls;

namespace BatchPowerPointToPDF.WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public PptxExporter Exporter = new PptxExporter();
        public ObservableCollection<ListViewItem> PptxFilenames = new ObservableCollection<ListViewItem>();

        public MainWindow()
        {
            InitializeComponent();
            filenamesListView.ItemsSource = PptxFilenames;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (!Exporter.OfficeInstalled())
            {
                MessageBox.Show("PowerPoint is required to use this tool. Please install PowerPoint.");
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
            var openPptxDialog = new OpenFileDialog
            {
                InitialDirectory = "%documents%",
                Filter = "PowerPoint Presentations (*.PPTX)|*.PPTX",
                Multiselect = true,
                Title = "Select PowerPoint presentation(s)"
            };

            if (openPptxDialog.ShowDialog() ?? false)
            {
                foreach (var file in openPptxDialog.FileNames)
                {
                    // Making sure we don't add duplicate entries.

                    // TODO: Probably not efficient, research other ways to check for duplicate items.
                    // In Windows Forms, there is an option to perform "ContainsKey" but it is not available in WPF.

                    var item = new ListViewItem {Content = file};
                    var contains = false;
                    foreach (var compareItem in PptxFilenames)
                    {
                        if (compareItem.Content.ToString() == file)
                        {
                            contains = true;
                        }
                    }

                    if (!contains)
                    {
                        PptxFilenames.Add(item);
                    }
                }
            }
        }

        private void button_Copy_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Fix blocking of UI thread, implement ConvertToPDF as async function.
            // When implemented as a Task, sometimes the Task is not properly disposed, and therefore
            // PowerPoint is stuck and does not convert the file. 

            foreach (var item in PptxFilenames)
            {
                Exporter.ConvertToPdf(item.Content.ToString());
            }
        }

        private void removePPTXBtn_Click(object sender, RoutedEventArgs e)
        {
            if (filenamesListView.HasItems)
            {
                var list = filenamesListView.SelectedItems;
                var removedItems = new LinkedList<ListViewItem>();

                // Can't actually remove files here, since .NET/ C# will complain about us modifying the list while iterating through it.
                foreach (var item in list)
                {
                    removedItems.AddFirst(item as ListViewItem);
                }

                foreach (var item in removedItems)
                {
                    PptxFilenames.Remove(item);
                }
            }
        }

        private void openInPPT_Click(object sender, RoutedEventArgs e)
        {
            if (filenamesListView.SelectedItems != null)
            {
                foreach (ListViewItem file in filenamesListView.SelectedItems)
                {
                    Exporter.OpenInPowerPoint(file.Content.ToString());
                }    
            }
                
        }
    }
}
