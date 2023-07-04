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
    public class findtextblock
    {
        public static List<TextBlock> FindTextBlocks(DependencyObject parent)
        {
            List<TextBlock> textBlocks = new List<TextBlock>();
            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(parent, i);
                if (child is TextBlock)
                {
                    textBlocks.Add(child as TextBlock);
                }
                else
                {
                    textBlocks.AddRange(FindTextBlocks(child));
                }
            }
            return textBlocks;
        }

    }
}
