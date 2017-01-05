using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing;

namespace PPT_Section_Indicator
{
    public partial class MainRibbon
    {
        bool includeSlideMarkers;
        IEnumerable<int> slideNumbers;

        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.ThisAddIn.Application.AfterPresentationOpen += new PowerPoint.EApplication_AfterPresentationOpenEventHandler(enableAddIn);
            Globals.ThisAddIn.Application.AfterNewPresentation += new PowerPoint.EApplication_AfterNewPresentationEventHandler(enableAddIn);
            Globals.ThisAddIn.Application.PresentationClose += new PowerPoint.EApplication_PresentationCloseEventHandler(disableAddIn);

            disableAddIn(null);
        }

        private void SlideMarkerCheckBox_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void StartButton_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            includeSlideMarkers = slideMarkerCheckBox.Checked;
            try
            {
                slideNumbers = Util.GetSlidesFromRangeExpr(slideRangeEditBox.Text);
            }
            catch (SlideRangeFormatException)
            {
                //TODO Display error message
            }

            PowerPoint.Slide firstSlide = presentation.Slides[slideNumbers.First()];
            firstSlide.Select();

            StepOneInsertFormatPlaceholders(firstSlide);

        }

        private void StepOneNextButton_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void StepOneInsertFormatPlaceholders(PowerPoint.Slide slide)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;

            PowerPoint.Shape textBox = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 10, 10, 100, 10);
            textBox.TextFrame.TextRange.InsertAfter("Active section");
            textBox.TextFrame.TextRange.Font.Color.RGB = Color.FromArgb(0, 0, 0).ToArgb();
            textBox.TextFrame.TextRange.Font.Size = 12;
            textBox.Name = "SectionIndicator_Format_ActiveSection";

            textBox = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 110, 10, 100, 10);
            textBox.TextFrame.TextRange.InsertAfter("Inactive section");
            textBox.TextFrame.TextRange.Font.Color.RGB = Color.FromArgb(190, 190, 190).ToArgb();
            textBox.TextFrame.TextRange.Font.Size = 12;

            if (includeSlideMarkers)
            {
                PowerPoint.Shape slideMarker = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, 18, 30, 8, 8);
                slideMarker.Fill.ForeColor.RGB = Color.FromArgb(0, 0, 0).ToArgb();
                slideMarker.Line.ForeColor.RGB = Color.FromArgb(0, 0, 0).ToArgb();

                slideMarker = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, 30, 30, 8, 8);
                slideMarker.Fill.ForeColor.RGB = Color.FromArgb(255, 255, 255).ToArgb();
                slideMarker.Line.ForeColor.RGB = Color.FromArgb(0, 0, 0).ToArgb();

                slideMarker = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, 118, 30, 8, 8);
                slideMarker.Fill.ForeColor.RGB = Color.FromArgb(255, 255, 255).ToArgb();
                slideMarker.Line.ForeColor.RGB = Color.FromArgb(190, 190, 190).ToArgb();
            }

        }

        public void enableAddIn(PowerPoint.Presentation presentation)
        {
            slideMarkerCheckBox.Enabled = true;
            slideRangeEditBox.Enabled = true;
            startButton.Enabled = true;
            stepOneNextButton.Enabled = true;
        }

        public void disableAddIn(PowerPoint.Presentation presentation)
        {
            slideMarkerCheckBox.Enabled = false;
            slideRangeEditBox.Enabled = false;
            startButton.Enabled = false;
            stepOneNextButton.Enabled = false;
        }


    }
}
