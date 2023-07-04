using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows;
//abs
namespace UserControlLib
{
    public class findtextbox
    {
        public static List<TextBox> FindTextBoxes(DependencyObject parent)
        {
            List<TextBox> textBoxes = new List<TextBox>();
            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(parent, i);
                if (child is TextBox)
                {
                    textBoxes.Add(child as TextBox);
                }
                else
                {
                    textBoxes.AddRange(FindTextBoxes(child));
                }
            }
            return textBoxes;
        }

    }
}
