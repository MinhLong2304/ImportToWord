using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows;
using System.ComponentModel;
using Word = Microsoft.Office.Interop.Word;
using System.Xml.Linq;

namespace UserControlLib
{
    public class LogicaTree
    {
        public static void TraverseLogicalTree(Word.Application wordApp, string worddocpath, DependencyObject parent, FrameworkElement container)
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(parent, i);
                FrameworkElement childFrameworkElement = child as FrameworkElement;

                // Call MessageBox.Show for each element
                MessageBox.Show(childFrameworkElement?.GetType().ToString());
                exportmediaelement.ExportMediaElement(wordApp, childFrameworkElement, worddocpath, container);

                // Recursively traverse the logical tree
                TraverseLogicalTree(wordApp, worddocpath, childFrameworkElement,container);
            }
        }
    }
}
