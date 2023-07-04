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

namespace UserControlLib
{
    public class addtextbox
    {
        public static void AddManyTextBoxesToWordDocument(Word.Application wordApp, string filePath, FrameworkElement container)
        {
            List<TextBox> textBoxes = findtextbox.FindTextBoxes(container);
            // Open an existing document
            Document wordDoc = wordApp.Documents.Open(filePath);

            foreach (var textBox in textBoxes)
            {
                //Calculate margin
                Thickness margin = calculatemargin.CalculateElementMargin(container, textBox);
                // Add a new text box shape to the document
                Shape textBoxShape = wordDoc.Shapes.AddTextbox(
                   Orientation: Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                    Left: (float)margin.Left * 72 / 96,  // Set the left position
                    Top: (float)margin.Top * 72 / 96,  // Set the top position
                    Width: (float)textBox.ActualWidth,  // Set the width
                    Height: (float)textBox.ActualHeight);  // Set the height



                // Add text to the text box
                textBoxShape.TextFrame.TextRange.Text = textBox.Text;

                //hide outline
                textBoxShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;


                // align text
                textBoxShape.TextFrame.TextRange.ParagraphFormat.Alignment = (WdParagraphAlignment)textBox.TextAlignment;

                // Get font color
                SolidColorBrush brush = (SolidColorBrush)textBox.Foreground;
                byte r = brush.Color.R;
                byte g = brush.Color.G;
                byte b = brush.Color.B;

                // Set font color
                textBoxShape.TextFrame.TextRange.Font.Color = (WdColor)(r + 0x100 * g + 0x10000 * b);

                // Set font size
                textBoxShape.TextFrame.TextRange.Font.Size = (float)textBox.FontSize * 72 / 96;

                // Set font family
                textBoxShape.TextFrame.TextRange.Font.Name = textBox.FontFamily.Source;

                // Set font weight to bold if FontWeight is set to Bold
                if (textBox.FontWeight == FontWeights.Bold)
                {
                    textBoxShape.TextFrame.TextRange.Font.Bold = 1;
                }

                // Set font style to italic if FontStyle is set to Italic
                if (textBox.FontStyle == FontStyles.Italic)
                {
                    textBoxShape.TextFrame.TextRange.Font.Italic = 1;

                }

                // Set text decoration to underline if Underline TextDecoration is present in textBox
                if (textBox.TextDecorations != null && textBox.TextDecorations.Contains(TextDecorations.Underline[0]))
                {
                    textBoxShape.TextFrame.TextRange.Font.Underline = WdUnderline.wdUnderlineSingle;
                }

                // Set margin left and top to 0
                textBoxShape.TextFrame.MarginLeft = 0;
                textBoxShape.TextFrame.MarginTop = 0;
            }

            // Save the document
            wordDoc.Save();
            /*
                        // Close the document and Word application
                        wordDoc.Close();*/

        }
    }
}
