using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void changer_Click(object sender, RibbonControlEventArgs e)
        {
            work_with_Slides(Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange);


        }
        private void work_with_Slides(SlideRange SldRange)
        {
            PowerPoint.Presentation currentPresentation = Globals.ThisAddIn.Application.ActivePresentation;
            float slideWidth = currentPresentation.PageSetup.SlideWidth;

            foreach (Slide slide in SldRange)
            {
                foreach (Shape shape in slide.Shapes)
                {

                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                    {
                       
                        
                        if((shape.Left + float.Parse(editBox1.Text, CultureInfo.InvariantCulture) * 1.333f) < slideWidth)
                        {
                             shape.Left += float.Parse(editBox1.Text, CultureInfo.InvariantCulture) * 1.333f;
                        }
                       

                        shape.TextEffect.FontBold = Microsoft.Office.Core.MsoTriState.msoCTrue;
                        shape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.Red);
                    }
                }
            }
        }
    }
}
