using PPTXExporterLibrary;
using System;
using System.Collections.Generic;
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
using Microsoft.Win32;

namespace BatchPowerPointToPDF.WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        LinkedList<String> pptxFilenames = new LinkedList<String>();
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            bool officeInstalled = PPTXExporter.OfficeInstalled();

            officeInstalledLabel.Content += officeInstalled.ToString().ToUpper();
        }

        private void openPPTXBtn_Click(object sender, RoutedEventArgs e)
        {
            pptxFilenames = OpenPDF();
        }

        private LinkedList<String> OpenPDF()
        {
            LinkedList<String> pptxFilenames = new LinkedList<string>();
            // Initializing Dialog Box
            OpenFileDialog openPPTXDialog = new OpenFileDialog();
            openPPTXDialog.InitialDirectory = "%documents%";
            openPPTXDialog.Filter = "PowerPoint Presentations (*.PPTX)|*.PPTX";
            openPPTXDialog.Multiselect = true;
            openPPTXDialog.Title = "Select PowerPoint presentation(s)";

            // Opening Dialog Box
            bool? dialogResult = openPPTXDialog.ShowDialog();

            if (dialogResult ?? false)
            {
                foreach (String file in openPPTXDialog.FileNames)
                {
                    pptxFilenames.AddFirst(file.ToString());
                }
            }


            return pptxFilenames;
        }

        private void button_Copy_Click(object sender, RoutedEventArgs e)
        {
            foreach (String file in pptxFilenames)
            {
                // So we do not block the UI thread.
                Task convert = Task.Factory.StartNew(() => PPTXExporterLibrary.PPTXExporter.ConvertToPDF(file));
                if (convert.IsCompleted)
                {
                    convert.Dispose();
                }
            }
        }
    }
}
