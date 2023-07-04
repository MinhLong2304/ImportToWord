using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace UserControlLib
{
    public class createworddoc
    {
        public static void CreateNewWordDoc(Word.Application wordApp, string worddocpath)
        {
            Document doc = wordApp.Documents.Add();
            doc.SaveAs2(worddocpath);
           /* doc.Close();*/
        }
    }
}
