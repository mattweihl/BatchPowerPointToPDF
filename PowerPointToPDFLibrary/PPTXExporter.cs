using System;
using Microsoft.Office.Core;
using Microsoft.Win32;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;

namespace PowerPointToPDFLibrary
{
    public class PptxExporter
    {
        PowerPoint.Application app;
        PowerPoint.Presentations presentation;
        public PptxExporter()
        {
            app = new PowerPoint.Application();
            presentation = app.Presentations;
        }

        /// <summary>
        /// Converts the given PowerPoint presentation to a PDF.
        /// </summary>
        /// <param name="pptxFilename"></param>
        public void ConvertToPdf(string pptxFilename)
        {
            try
            {
                if (!OfficeInstalled())
                {
                    throw new NotSupportedException("Office Installation Not Detected");
                }

                // Adding Escape Characters
                pptxFilename = pptxFilename.Replace(@"\\", @"\");

                // Opening PowerPoint
                PowerPoint.Presentation file = app.Presentations.Open(pptxFilename, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

                // Converting to PDF
                file.ExportAsFixedFormat(pptxFilename + ".pdf", PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
            }

            catch (Exception e)
            {
                WriteToLogFile("EXCEPTION: " + e.Message + " -- " + e.StackTrace);
            }
        }

        /// <summary>
        /// Returns whether Office is installed on the Local Machine.
        /// </summary>
        /// <returns></returns>
        public bool OfficeInstalled()
        {
            RegistryKey powerpointKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe");

            powerpointKey?.Close();

            return powerpointKey != null;
        }

        public void OpenInPowerPoint(String filename)
        {
            PowerPoint.Presentation file = app.Presentations.Open(filename, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
        }

        /// <summary>
        /// Writes specified String to log.txt in the program's current directory.
        /// </summary>
        /// <param name="line"></param>
        private void WriteToLogFile(String line)
        {
            String logPath = Directory.GetCurrentDirectory() + "\\log.txt";

            if (!File.Exists(logPath))
            {
                FileStream logFile = File.Create(logPath);
                logFile.Close();
            }

            TextWriter logWriter = new StreamWriter(logPath, true);
            logWriter.WriteLine(line);
            logWriter.Close();
        }
    }
}
