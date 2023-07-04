using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Web.WebView2.Wpf;
using System.Windows.Documents;
using System.Threading;
using System.Xml.Linq;
using System.Windows.Controls.Primitives;
using System.Windows.Shapes;

namespace UserControlLib
{
    public class scanandinsertvisual
    {
        
        public static void TraverseVisualTree2(Word.Application wordApp, string worddocpath, Visual visual, FrameworkElement container)
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(visual); i++)
            {
                Visual childVisual = (Visual)VisualTreeHelper.GetChild(visual, i);
                FrameworkElement childFrameworkElement = childVisual as FrameworkElement;

                // Export the FrameworkElement
                if (childFrameworkElement != null)
                {
                    if(!(childFrameworkElement is Button|| childFrameworkElement is RadioButton || childFrameworkElement is TextBlock || childFrameworkElement is ContentPresenter || childFrameworkElement is Label || childFrameworkElement is System.Windows.Documents.AdornerLayer||childFrameworkElement is ItemsPresenter|| childFrameworkElement is WrapPanel||childFrameworkElement is Popup||childFrameworkElement is Path|| childFrameworkElement is VirtualizingStackPanel))
                    {
                       
                          /*  MessageBox.Show(childFrameworkElement?.GetType().ToString());*/
                        


                        exportmediaelement.ExportMediaElement(wordApp, childFrameworkElement, worddocpath, container);
                    }
                } 
              


                if (childFrameworkElement is TextBox || childFrameworkElement is Slider || childFrameworkElement is ComboBox || childFrameworkElement is DataGrid)
                {

                    continue;
                }

                

                if (childFrameworkElement is WebView2)
                {
                    WebView2 webView2 = (WebView2)childFrameworkElement;
                    Renderwebview.Render(wordApp, webView2, worddocpath, container);
                }
               




                // Traverse the visual tree recursively
                TraverseVisualTree2(wordApp, worddocpath, childVisual, container);
            }
        }
    }
}
