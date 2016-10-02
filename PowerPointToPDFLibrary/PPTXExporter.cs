using System;
using Microsoft.Office.Core;
using Microsoft.Win32;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;

namespace PowerPointToPDFLibrary
{
    public class PptxExporter
    {
        private PowerPoint.Application _app;
        private PowerPoint.Presentations _presentation;
        public PptxExporter()
        {
            _app = new PowerPoint.Application();
            _presentation = _app.Presentations;
        }

        /// <summary>
        /// Converts the given PowerPoint presentation to a PDF.
        /// </summary>
        /// <param name="pptxFilename"></param>
        public void ConvertToPdf(string pptxFilename)
        {
            try
            {
                // Adding Escape Characters
                pptxFilename = pptxFilename.Replace(@"\\", @"\");

                // Opening PowerPoint
                var file = _app.Presentations.Open(pptxFilename, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

                // Converting to PDF
                file.ExportAsFixedFormat(pptxFilename + ".pdf", PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
                WriteToLogFile("Converted to PDF: " + pptxFilename);
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
            var powerpointKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe");

            if (powerpointKey != null)
            {
                WriteToLogFile("Office installation detected.");
                return true;
            }
                WriteToLogFile("Office installation not detected.");
                return false;
        }

        public void OpenInPowerPoint(string filename)
        {
            var file = _app.Presentations.Open(filename, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
            WriteToLogFile("Opened in PowerPoint: " + filename);
        }

        /// <summary>
        /// Writes specified message to log.txt in the program's current directory.
        /// </summary>
        /// <param name="message"></param>
        private void WriteToLogFile(string message)
        {
            var logPath = Directory.GetCurrentDirectory() + "\\log.txt";

            if (!File.Exists(logPath))
            {
                var logFile = File.Create(logPath);
                logFile.Close();
            }

            TextWriter logWriter = new StreamWriter(logPath, true);
            logWriter.WriteLine(DateTimeOffset.Now.ToString() + ": " + message);
            logWriter.Close();
        }
    }
}
