using System.Windows;
using System.Windows.Media;

namespace UserControlLib
{
    public class begin
    {
        public static void Begin(string docfilepath, FrameworkElement container, Visual visual)
        {


            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            createworddoc.CreateNewWordDoc(wordApp, docfilepath);
            wordApp.Visible = true;
            setuppaper.SetupPaper(wordApp, container.ActualWidth, container.ActualHeight, docfilepath);

            scanandinsertvisual.TraverseVisualTree2(wordApp, docfilepath, visual, container);

            addtextblock.AddManyTextBlocksToWordDocument(wordApp, docfilepath, container);
            addtextbox.AddManyTextBoxesToWordDocument(wordApp, docfilepath, container);

            /*wordApp.Quit();*/
            DeleteExport.clearexport();
            MessageBox.Show("Done");
        }
    }
}
