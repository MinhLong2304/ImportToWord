using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace UserControlLib
{
    public class setuppaper
    {
        public static void SetupPaper(Word.Application wordApp, double WindowWidth, double WindowHeight, string filePath)
        {



            // Get the current width and height of the main window in pixels
            double widthInPixels = WindowWidth;
            double heightInPixels = WindowHeight;

            // Convert the pixel values to points
            double widthInPoints = widthInPixels * 72 / 96; // 96 DPI is the standard screen DPI, 72 DPI is the standard printing DPI
            double heightInPoints = heightInPixels * 72 / 96;



            // Open an existing Word document
            Document doc = wordApp.Documents.Open(filePath);

            // Set the paper size to custom
            doc.PageSetup.PaperSize = WdPaperSize.wdPaperCustom;

            // Set the custom paper width and height to the values obtained from the main window
            doc.PageSetup.PageWidth = (int)widthInPoints;
            doc.PageSetup.PageHeight = (int)heightInPoints;

            // Set the margin value to 0
            doc.PageSetup.LeftMargin = 0;
            doc.PageSetup.RightMargin = 0;
            doc.PageSetup.TopMargin = 0;
            doc.PageSetup.BottomMargin = 0;

            // Save and close the document
            doc.Save();
            /*doc.Close();*/


        }
    }
}
