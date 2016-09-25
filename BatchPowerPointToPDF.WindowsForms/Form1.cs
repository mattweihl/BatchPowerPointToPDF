﻿using System;
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
        LinkedList<String> PPTXFileNames = new LinkedList<string>();
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

                MessageBox.Show("Office Is Not Installed");
                Application.Exit();

            }
        }

        private void toolStripStatusLabel1_Click(object sender, EventArgs e)
        {

        }

        private void openPPTXBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openPPTXDialog = new OpenFileDialog();

            openPPTXDialog.InitialDirectory = "%documents%";
            openPPTXDialog.Filter = "PowerPoint Presentations (*.PPTX)|*.PPTX";
            openPPTXDialog.Multiselect = true;
            openPPTXDialog.Title = "Select PowerPoint presentation(s)";

            DialogResult dr = openPPTXDialog.ShowDialog();

            if (dr == DialogResult.OK)
            {
                foreach (String file in openPPTXDialog.FileNames)
                {
                    PPTXFileNames.AddFirst(file.ToString());
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // TODO: Investigate non-blocking UI approach.
            foreach(String x in PPTXFileNames)
            {
                PPTXExporterLibrary.PPTXExporter.ConvertToPDF(x);
            }
        }
    }
}