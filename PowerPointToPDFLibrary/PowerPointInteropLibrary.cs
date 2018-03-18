using System;
using Microsoft.Office.Core;
using Microsoft.Win32;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OfficeInterop
{
    public class PowerPointInteropLibrary
    {
        private PowerPoint.Application _app;
        private PowerPoint.Presentations _presentation;
        public PowerPointInteropLibrary()
        {
            _app = new PowerPoint.Application();
            _presentation = _app.Presentations;
        }

        /// <summary>
        /// Converts the given <see cref="PowerPoint"/> presentation to a PDF.
        /// </summary>
        /// <param name="pptFilename"></param>
        public bool ConvertToPdf(string pptFilename)
        {
            try
            {
                // Adding Escape Characters
                pptFilename = pptFilename.Replace(@"\\", @"\");

                // Opening PowerPoint
                var file = _app.Presentations.Open(pptFilename, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

                // Converting to PDF
                file.ExportAsFixedFormat(pptFilename + ".pdf", PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);

                return true;
            }

            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Returns whether Office is installed on the Local Machine.
        /// </summary>
        /// <returns></returns>
        public bool OfficeInstalled() => (Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe") != null);

        public void OpenInPowerPoint(string filename)
        {
            var file = _app.Presentations.Open(filename, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
        }

    }
}
