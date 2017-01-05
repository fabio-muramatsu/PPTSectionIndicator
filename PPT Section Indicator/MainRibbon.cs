using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing;
using System.Collections;

namespace PPT_Section_Indicator
{
    public partial class MainRibbon
    {
        private const string FORMAT_ACTIVE_SECTION_TEXT_BOX = "SectionIndicator_Format_ActiveSectionTextBox";
        private const string FORMAT_INACTIVE_SECTION_TEXT_BOX = "SectionIndicator_Format_InactiveSectionTextBox";
        private const string FORMAT_ACTIVE_SECTION_SLIDE_MARKER = "SectionIndicator_Format_ActiveSectionSlideMarker";
        private const string FORMAT_CURRENT_SLIDE_SLIDE_MARKER = "SectionIndicator_Format_CurrentSlideSlideMarker";
        private const string FORMAT_INACTIVE_SECTION_SLIDE_MARKER = "SectionIndicator_Format_InactiveSectionSlideMarker";

        private const string POSITION_TEXT_BOX = "SectionIndicator_Position_TextBox";
        private const string POSITION_SLIDE_MARKER = "SectionIndicator_Position_SlideMarker";

        private Dictionary<String, PowerPoint.Shape> formatShapes = new Dictionary<string, PowerPoint.Shape>();
        private Dictionary<int, PowerPoint.Shape> positionTextBoxes = new Dictionary<int, PowerPoint.Shape>();

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
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slide firstSlide = presentation.Slides[slideNumbers.First()];

            int currentSection = -1, previousSection = -1;
            try
            {
                previousSection = Util.GetSectionIndex(slideNumbers.First());
            }
            catch (NoSectionException)
            {
                //show error
            }
            foreach (int slideIndex in slideNumbers)
            {
                if (currentSection == -1)
                {
                    StepTwoInsertTextBoxes(previousSection);
                }
                currentSection = Util.GetSectionIndex(slideIndex);
                if(currentSection != previousSection)
                {
                    StepTwoInsertTextBoxes(currentSection);
                }
                previousSection = currentSection;
            }

            foreach(PowerPoint.Shape shape in formatShapes.Values)
            {
                shape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            }
        }

        private void StepOneInsertFormatPlaceholders(PowerPoint.Slide slide)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;

            PowerPoint.Shape textBox = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 10, 10, 100, 10);
            textBox.TextFrame.TextRange.InsertAfter("Active section");
            textBox.TextFrame.TextRange.Font.Color.RGB = Color.FromArgb(0, 0, 0).ToArgb();
            textBox.TextFrame.TextRange.Font.Size = 12;
            textBox.Name = FORMAT_ACTIVE_SECTION_TEXT_BOX;
            formatShapes.Add(FORMAT_ACTIVE_SECTION_TEXT_BOX, textBox);

            textBox = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 110, 10, 100, 10);
            textBox.TextFrame.TextRange.InsertAfter("Inactive section");
            textBox.TextFrame.TextRange.Font.Color.RGB = Color.FromArgb(190, 190, 190).ToArgb();
            textBox.TextFrame.TextRange.Font.Size = 12;
            textBox.Name = FORMAT_INACTIVE_SECTION_TEXT_BOX;
            formatShapes.Add(FORMAT_INACTIVE_SECTION_TEXT_BOX, textBox);


            if (includeSlideMarkers)
            {
                PowerPoint.Shape slideMarker = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, 18, 30, 8, 8);
                slideMarker.Fill.ForeColor.RGB = Color.FromArgb(0, 0, 0).ToArgb();
                slideMarker.Line.ForeColor.RGB = Color.FromArgb(0, 0, 0).ToArgb();
                slideMarker.Name = FORMAT_CURRENT_SLIDE_SLIDE_MARKER;
                formatShapes.Add(FORMAT_CURRENT_SLIDE_SLIDE_MARKER, slideMarker);

                slideMarker = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, 30, 30, 8, 8);
                slideMarker.Fill.ForeColor.RGB = Color.FromArgb(255, 255, 255).ToArgb();
                slideMarker.Line.ForeColor.RGB = Color.FromArgb(0, 0, 0).ToArgb();
                slideMarker.Name = FORMAT_ACTIVE_SECTION_SLIDE_MARKER;
                formatShapes.Add(FORMAT_ACTIVE_SECTION_SLIDE_MARKER, slideMarker);

                slideMarker = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, 118, 30, 8, 8);
                slideMarker.Fill.ForeColor.RGB = Color.FromArgb(255, 255, 255).ToArgb();
                slideMarker.Line.ForeColor.RGB = Color.FromArgb(190, 190, 190).ToArgb();
                slideMarker.Name = FORMAT_INACTIVE_SECTION_SLIDE_MARKER;
                formatShapes.Add(FORMAT_INACTIVE_SECTION_SLIDE_MARKER, slideMarker);
            }

        }

        private void StepTwoInsertTextBoxes(int section)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slide firstSlide = presentation.Slides[slideNumbers.First()];
            if (section == 1)
            {
                PowerPoint.Shape textBox;
                formatShapes.TryGetValue(FORMAT_ACTIVE_SECTION_TEXT_BOX, out textBox);
                textBox.Copy();
                IEnumerator enumerator = firstSlide.Shapes.Paste().GetEnumerator();
                enumerator.MoveNext();
                PowerPoint.Shape newTextBox = (PowerPoint.Shape)enumerator.Current;
                newTextBox.Left = 10;
                newTextBox.Top = 10;
                newTextBox.Name = POSITION_TEXT_BOX + "_" + section;
                newTextBox.TextFrame.TextRange.Text = presentation.SectionProperties.Name(section);
                positionTextBoxes.Add(section, newTextBox);
            }
            else
            {
                PowerPoint.Shape textBox;
                formatShapes.TryGetValue(FORMAT_INACTIVE_SECTION_TEXT_BOX, out textBox);
                textBox.Copy();
                IEnumerator enumerator = firstSlide.Shapes.Paste().GetEnumerator();
                enumerator.MoveNext();
                PowerPoint.Shape newTextBox = (PowerPoint.Shape)enumerator.Current;
                newTextBox.Left = 100 * (section - 1) + 10;
                newTextBox.Top = 10;
                newTextBox.Name = POSITION_TEXT_BOX + "_" + section;
                newTextBox.TextFrame.TextRange.Text = presentation.SectionProperties.Name(section);
                positionTextBoxes.Add(section, newTextBox);
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
