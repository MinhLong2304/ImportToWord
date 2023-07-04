using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows;
using System.ComponentModel;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Controls;
using System.IO;
using Microsoft.Web.WebView2.Wpf;

namespace UserControlLib
{
    public class Renderwebview
    {
        public static void Render(Word.Application wordApp, WebView2 element, string worddocpath, FrameworkElement container)
        {
            string rootFolder = AppDomain.CurrentDomain.BaseDirectory;
            string exportFolder = System.IO.Path.Combine(rootFolder, "export");

            if (!Directory.Exists(exportFolder))
            {
                Directory.CreateDirectory(exportFolder);
            }

            string filePath = System.IO.Path.Combine(exportFolder, element.Name + ".png");
            if (element == null)
            {
                MessageBox.Show("Element is null.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Invalid file path.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            // Get the bounding box of the element, including any margins or padding
            Rect bounds = VisualTreeHelper.GetDescendantBounds(element);
            System.Windows.Point topLeft = element.PointToScreen(bounds.TopLeft);
            System.Windows.Point bottomRight = element.PointToScreen(bounds.BottomRight);
            int width = (int)(bottomRight.X - topLeft.X);
            int height = (int)(bottomRight.Y - topLeft.Y);

            // Create a bitmap of the Webview control
            var bmp = new Bitmap(width, height);
            var gfx = Graphics.FromImage(bmp);
            gfx.CopyFromScreen((int)topLeft.X, (int)topLeft.Y, 0, 0, bmp.Size, CopyPixelOperation.SourceCopy);

            // Save the bitmap to a file
            bmp.Save(filePath);


            float positionLeft = 0;
            float positionTop = 0;

            // Calculate margin
            if (container != element)
            {
                Thickness margin = calculatemargin.CalculateElementMargin(container, element);

                // Insert image
                positionLeft = (float)margin.Left;
                positionTop = (float)margin.Top;
            }

            else
            {
                positionLeft = (float)element.Margin.Left;
                positionTop = (float)element.Margin.Top;
            }
            insertimageelement.InsertImageElements(wordApp, worddocpath, filePath, positionLeft, positionTop);
        }
    }
}
