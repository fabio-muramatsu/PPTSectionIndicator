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

        private const int DEFAULT_SECTION_SPACING = 150;

        private Dictionary<String, PowerPoint.Shape> formatShapes = new Dictionary<string, PowerPoint.Shape>();
        private Dictionary<int, PowerPoint.Shape> positionTextBoxes = new Dictionary<int, PowerPoint.Shape>();
        private Dictionary<int, PowerPoint.Shape> positionMarkers = new Dictionary<int, PowerPoint.Shape>();

        bool includeSlideMarkers;
        IEnumerable<int> slideNumbers;
        IDictionary<int, IList<int>> slidesPerSection;

        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.ThisAddIn.Application.AfterPresentationOpen += new PowerPoint.EApplication_AfterPresentationOpenEventHandler(enableAddInStart);
            Globals.ThisAddIn.Application.AfterNewPresentation += new PowerPoint.EApplication_AfterNewPresentationEventHandler(enableAddInStart);
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
            catch (SlideRangeFormatException srfe)
            {
                Util.ShowErrorMessage(srfe.Message);
                return;
            }

            if(presentation.SectionProperties.Count == 0)
            {
                Util.ShowErrorMessage("There are no sections in the presentation");
                return;
            }

            PowerPoint.Slide firstSlide = presentation.Slides[slideNumbers.First()];
            firstSlide.Select();

            StepOneInsertFormatPlaceholders(firstSlide);
            enableAddInStepOne();
        }

        private void StepOneNextButton_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slide firstSlide = presentation.Slides[slideNumbers.First()];

            int currentSection = -1, previousSection = -1;
            try
            {
                previousSection = Util.GetSectionIndex(slideNumbers.First());
                slidesPerSection = Util.ClassifySlidesIntoSections(slideNumbers);
            }
            catch (NoSectionException nse)
            {
                Util.ShowErrorMessage(nse.Message);
                return;
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

                if (includeSlideMarkers)
                    StepTwoInsertSlideMarkers(currentSection, slideIndex);

                previousSection = currentSection;
            }

            foreach(PowerPoint.Shape shape in formatShapes.Values)
            {
                shape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            }
            enableAddInStepTwo();
        }

        private void StepTwoDoneButton_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void CleanupButton_Click(object sender, RibbonControlEventArgs e)
        {
            formatShapes.Clear();
            positionMarkers.Clear();
            positionTextBoxes.Clear();
            enableAddInStart(null);
            Util.CleanupShapes();
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
                newTextBox.Left = DEFAULT_SECTION_SPACING * (section - 1) + 10;
                newTextBox.Top = 10;
                newTextBox.Name = POSITION_TEXT_BOX + "_" + section;
                newTextBox.TextFrame.TextRange.Text = presentation.SectionProperties.Name(section);
                positionTextBoxes.Add(section, newTextBox);
            }
        }

        private void StepTwoInsertSlideMarkers(int section, int slideIndex)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slide firstSlide = presentation.Slides[slideNumbers.First()];

            PowerPoint.Shape marker;
            if(slideIndex == slideNumbers.First())
            {
                formatShapes.TryGetValue(FORMAT_CURRENT_SLIDE_SLIDE_MARKER, out marker);
            }
            else if (section == 1)
            {
                formatShapes.TryGetValue(FORMAT_ACTIVE_SECTION_SLIDE_MARKER, out marker);
            }
            else
            {
                formatShapes.TryGetValue(FORMAT_INACTIVE_SECTION_SLIDE_MARKER, out marker);
            }

            int maxNumberOfMarkers = (int)Math.Floor(DEFAULT_SECTION_SPACING / (marker.Width + 2));

            marker.Copy();
            IEnumerator enumerator = firstSlide.Shapes.Paste().GetEnumerator();
            enumerator.MoveNext();
            PowerPoint.Shape newMarker = (PowerPoint.Shape)enumerator.Current;

            int slideIndexWithinSection = Util.GetSlideIndexWithinSection(slidesPerSection, slideIndex);
            float left = 18 + (section - 1) * DEFAULT_SECTION_SPACING + ((slideIndexWithinSection - 1) % maxNumberOfMarkers) * (newMarker.Width + 2);
            float top = 10 + positionTextBoxes[section].Height + ((slideIndexWithinSection - 1)/maxNumberOfMarkers) * (marker.Height + 5);

            newMarker.Left = left;
            newMarker.Top = top;
            newMarker.Name = POSITION_SLIDE_MARKER + "_" + section + "_" + slideIndex;
            positionMarkers.Add(slideIndex, newMarker);
        }

        public void enableAddInStart(PowerPoint.Presentation presentation)
        {
            slideMarkerCheckBox.Enabled = true;
            slideRangeEditBox.Enabled = true;
            startButton.Enabled = true;
            stepOneNextButton.Enabled = false;
            stepTwoDoneButton.Enabled = false;
            cleanPresentationButton.Enabled = true;
        }

        public void enableAddInStepOne()
        {
            slideMarkerCheckBox.Enabled = false;
            slideRangeEditBox.Enabled = false;
            startButton.Enabled = false;
            stepOneNextButton.Enabled = true;
            stepTwoDoneButton.Enabled = false;
        }

        public void enableAddInStepTwo()
        {
            slideMarkerCheckBox.Enabled = false;
            slideRangeEditBox.Enabled = false;
            startButton.Enabled = false;
            stepOneNextButton.Enabled = false;
            stepTwoDoneButton.Enabled = true;
        }

        public void disableAddIn(PowerPoint.Presentation presentation)
        {
            slideMarkerCheckBox.Enabled = false;
            slideRangeEditBox.Enabled = false;
            startButton.Enabled = false;
            stepOneNextButton.Enabled = false;
            cleanPresentationButton.Enabled = false;
        }
    }
}
