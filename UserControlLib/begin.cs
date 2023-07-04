using System.Windows;
using System.Windows.Media;

namespace UserControlLib
{
    public class begin
    {
        public static void Begin(string docfilepath, Visual visual, FrameworkElement container)
        {


            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            createworddoc.CreateNewWordDoc(wordApp, docfilepath);
            wordApp.Visible = true;
            setuppaper.SetupPaper(wordApp, container.ActualWidth, container.ActualHeight, docfilepath);
            exportmediaelement.ExportMediaElement(wordApp,container, docfilepath, container);
            scanandinsertvisual.TraverseVisualTree2(wordApp, docfilepath, visual, container);

            addtextblock.AddManyTextBlocksToWordDocument(wordApp, docfilepath, container);
            addtextbox.AddManyTextBoxesToWordDocument(wordApp, docfilepath, container);

            /*wordApp.Quit();*/
            //DeleteExport.clearexport();
            MessageBox.Show("Done");
        }
    }
}
