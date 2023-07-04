using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace UserControlLib
{
    public  class insertimageelement
    {
        public static void InsertImageElements(Word.Application wordApp, string fileName, string imagePath, float positionleft, float positontop)
        {

            Document doc = new Document();

            // Open the Word document
            doc = wordApp.Documents.Open(fileName);
            // Insert the image
            Shape shape = doc.Shapes.AddPicture(imagePath);


            // Set the wrap format to be in front of text
            shape.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapSquare;

            //Set location
            shape.Left = positionleft * 72 / 96;
            shape.Top = positontop * 72 / 96;

            // Save the changes and close the document
            doc.Save();
            /*doc.Close();*/


        }
    }
}
