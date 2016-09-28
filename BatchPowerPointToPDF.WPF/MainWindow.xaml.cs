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
        LinkedList<String> pptxFilenames;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            bool officeInstalled = PPTXExporterLibrary.PPTXExporter.OfficeInstalled();

            officeInstalledLabel.Content += officeInstalled.ToString().ToUpper();

            if (!officeInstalled)
            {
                MessageBox.Show("Office is not installed. In order to continue, please install Office.");
                Application.Current.Shutdown();
            }
        }

        private void openPPTXBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenPDF();
        }

        private void OpenPDF()
        {
            pptxFilenames = new LinkedList<String>();

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
        }

        private void button_Copy_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Fix blocking of UI thread, implement ConvertToPDF as async function.
            // When implemented as a Task, sometimes the Task is not properly disposed, and therefore
            // PowerPoint is stuck and does not convert the file. 

            foreach (String file in pptxFilenames)
            {
                PPTXExporterLibrary.PPTXExporter.ConvertToPDF(file);
            }
        }
    }
}
