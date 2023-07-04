using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace UserControlLib
{
    public class calculatemargin
    {
        public static Thickness CalculateElementMargin(UIElement element1, UIElement element2)
        {
            System.Windows.Window window = System.Windows.Window.GetWindow(element1);

            System.Windows.Point element1Location = element1.TranslatePoint(new System.Windows.Point(0, 0), window);
            System.Windows.Point element2Location = element2.TranslatePoint(new System.Windows.Point(0, 0), window);

            double leftMargin = element2Location.X - element1Location.X;
            double topMargin = element2Location.Y - element1Location.Y;

            return new Thickness(leftMargin, topMargin, 0, 0);
        }

    }
}
