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
    public class exportmediaelement
    {
        public static void ExportMediaElement(Word.Application wordApp, FrameworkElement mediaElement, string worddocpath, FrameworkElement container)
        {
            if (mediaElement.ActualWidth <= 0 || mediaElement.ActualHeight <= 0)
            {

                return;
            }
            string rootFolder = AppDomain.CurrentDomain.BaseDirectory;
            string exportFolder = System.IO.Path.Combine(rootFolder, "export");

            if (!Directory.Exists(exportFolder))
            {
                Directory.CreateDirectory(exportFolder);
            }

            string filePath = System.IO.Path.Combine(exportFolder, mediaElement.Name + ".png");
            if (mediaElement == null)
            {
                MessageBox.Show("Element is null.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Invalid file path.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            foreach (UIElement child in LogicalTreeHelper.GetChildren(mediaElement).OfType<UIElement>())
            {
                child.Visibility = Visibility.Hidden;
            }
            // Create a VisualBrush to render the MediaElement
            VisualBrush visualBrush = new VisualBrush(mediaElement);

            if (mediaElement.ActualWidth <= 0 || mediaElement.ActualHeight <= 0)
            {

                return;
            }

            // Create a DrawingVisual and draw the VisualBrush to it
            DrawingVisual drawingVisual = new DrawingVisual();
            using (DrawingContext drawingContext = drawingVisual.RenderOpen())
            {
                drawingContext.DrawRectangle(visualBrush, null, new Rect(0, 0, mediaElement.ActualWidth, mediaElement.ActualHeight));
            }

            // Create a RenderTargetBitmap and render the DrawingVisual to it
            RenderTargetBitmap renderTargetBitmap = new RenderTargetBitmap(
                (int)mediaElement.ActualWidth, (int)mediaElement.ActualHeight, 96, 96, PixelFormats.Pbgra32);
            renderTargetBitmap.Render(drawingVisual);


            // Create a PngBitmapEncoder to encode the RenderTargetBitmap as a PNG image
            PngBitmapEncoder encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(renderTargetBitmap));


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
           /* if (container != mediaElement)
            {*/
                Thickness margin = calculatemargin.CalculateElementMargin(container, mediaElement);

               
                positionLeft = (float)margin.Left;
                positionTop = (float)margin.Top;
           /* }*/

           /* else
            {
                positionLeft = (float)mediaElement.Margin.Left;
                positionTop = (float)mediaElement.Margin.Top;
            }*/
            //insert element
            insertimageelement.InsertImageElements(wordApp, worddocpath, filePath, positionLeft, positionTop);

            // Restore child element visibility
            foreach (UIElement child in LogicalTreeHelper.GetChildren(mediaElement).OfType<UIElement>())
            {
                child.Visibility = Visibility.Visible;
            }
        }
    }
}
