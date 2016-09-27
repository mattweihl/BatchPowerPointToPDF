using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BatchPowerPointToPDF.WindowsForms
{
    public partial class Form1 : Form
    {
        LinkedList<String> pptxFileNames = new LinkedList<String>();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            if (PPTXExporterLibrary.PPTXExporter.OfficeInstalled())
            {
                officeInstalledLabel.Text += "TRUE";
                
            }

            else
            {
                officeInstalledLabel.Text += "FALSE";

                // Office has not been detected on the computer. Therefore the application must exit.
                MessageBox.Show("Office Is Not Installed");
                Application.Exit();
            }
        }

        private void toolStripStatusLabel1_Click(object sender, EventArgs e)
        {

        }

        private void openPPTXBtn_Click(object sender, EventArgs e)
        {
             pptxFileNames = OpenPDF();
        }

        private LinkedList<String> OpenPDF()
        {
            LinkedList<String> PPTXFileNames = new LinkedList<string>();
            // Initializing Dialog Box
            OpenFileDialog openPPTXDialog = new OpenFileDialog();
            openPPTXDialog.InitialDirectory = "%documents%";
            openPPTXDialog.Filter = "PowerPoint Presentations (*.PPTX)|*.PPTX";
            openPPTXDialog.Multiselect = true;
            openPPTXDialog.Title = "Select PowerPoint presentation(s)";

            // Opening Dialog Box
            DialogResult dr = openPPTXDialog.ShowDialog();

            if (dr == DialogResult.OK)
            {
                // Adding user-selected filenames
                foreach (String file in openPPTXDialog.FileNames)
                {
                    PPTXFileNames.AddFirst(file.ToString());
                }
            }

            return PPTXFileNames;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            foreach(String x in pptxFileNames)
            {
                // So we do not block the UI thread.
                Task convert = Task.Factory.StartNew(() => PPTXExporterLibrary.PPTXExporter.ConvertToPDF(x));
                if (convert.IsCompleted)
                {
                    convert.Dispose();
                }
            }

        }
    }
}
