using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using System.Windows.Media;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace UserControlLib
{
    public class old_addimg
    {
        public static void ExportAndInsertUIElement(Word.Application wordApp, FrameworkElement element, string worddocpath, FrameworkElement container)
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

            // Hide child elements
            foreach (UIElement child in LogicalTreeHelper.GetChildren(element).OfType<UIElement>())
            {
                child.Visibility = Visibility.Hidden;
            }



            // Measure and arrange the element to determine its size
            element.Measure(new System.Windows.Size(double.PositiveInfinity, double.PositiveInfinity));
            element.Arrange(new Rect(new System.Windows.Point(0, 0), element.DesiredSize));

            // Create a VisualBrush to render the element
            VisualBrush visualBrush = new VisualBrush(element);

            // Create a DrawingVisual and draw the element to it
            DrawingVisual drawingVisual = new DrawingVisual();
            using (DrawingContext drawingContext = drawingVisual.RenderOpen())
            {
                drawingContext.DrawRectangle(visualBrush, null, new Rect(0, 0, element.ActualWidth, element.ActualHeight));
            }

            // Create a RenderTargetBitmap and render the DrawingVisual to it
            RenderTargetBitmap renderTargetBitmap = new RenderTargetBitmap(
                (int)element.ActualWidth, (int)element.ActualHeight, 96, 96, PixelFormats.Pbgra32);
            renderTargetBitmap.Render(drawingVisual);

            // Create a BitmapEncoder based on the file extension
            BitmapEncoder encoder = new PngBitmapEncoder();


            // Create a MemoryStream to hold the image data
            using (System.IO.MemoryStream stream = new System.IO.MemoryStream())
            {
                // Save the BitmapSource to the stream using the encoder
                encoder.Frames.Add(BitmapFrame.Create(renderTargetBitmap));
                encoder.Save(stream);

                // Write the stream to the file
                using (System.IO.FileStream file = new System.IO.FileStream(filePath, System.IO.FileMode.Create))
                {
                    stream.WriteTo(file);
                }
            }



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

            // Restore child element visibility
            foreach (UIElement child in LogicalTreeHelper.GetChildren(element).OfType<UIElement>())
            {
                child.Visibility = Visibility.Visible;
            }
        }
    }
}
