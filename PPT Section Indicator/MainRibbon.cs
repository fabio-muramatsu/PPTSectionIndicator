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
using System.Threading.Tasks;
using System.Threading;

namespace PPT_Section_Indicator
{
    public partial class MainRibbon
    {
        private const string FORMAT_ACTIVE_SECTION_TEXT_BOX = "SectionIndicator_Format_ActiveSectionTextBox";
        private const string FORMAT_INACTIVE_SECTION_TEXT_BOX = "SectionIndicator_Format_InactiveSectionTextBox";
        private const string FORMAT_ACTIVE_SECTION_SLIDE_MARKER = "SectionIndicator_Format_ActiveSectionSlideMarker";
        private const string FORMAT_CURRENT_SLIDE_SLIDE_MARKER = "SectionIndicator_Format_CurrentSlideSlideMarker";
        private const string FORMAT_INACTIVE_SECTION_SLIDE_MARKER = "SectionIndicator_Format_InactiveSectionSlideMarker";

        public static readonly string POSITION_TEXT_BOX = "SectionIndicator_Position_TextBox";
        public static readonly string POSITION_SLIDE_MARKER = "SectionIndicator_Position_SlideMarker";

        private const string GROUPED_SHAPES = "SectionIndicator_GroupedItems";

        private const int DEFAULT_SECTION_SPACING = 150;

        private Dictionary<String, PowerPoint.Shape> formatShapes = new Dictionary<string, PowerPoint.Shape>();
        private Dictionary<int, PowerPoint.Shape> positionTextBoxes = new Dictionary<int, PowerPoint.Shape>();
        private Dictionary<int, PowerPoint.Shape> positionMarkers = new Dictionary<int, PowerPoint.Shape>();

        bool includeSlideMarkers;
        IList<int> slideNumbers;
        IDictionary<int, IList<int>> slidesPerSection;

        private ProgressDialogBox progressDialog;

        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.ThisAddIn.Application.AfterPresentationOpen += new PowerPoint.EApplication_AfterPresentationOpenEventHandler(EnableAddInStart);
            Globals.ThisAddIn.Application.AfterNewPresentation += new PowerPoint.EApplication_AfterNewPresentationEventHandler(EnableAddInStart);
            Globals.ThisAddIn.Application.PresentationClose += new PowerPoint.EApplication_PresentationCloseEventHandler(DisableAddIn);

            DisableAddIn(null);
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
            EnableAddInStepOne();
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
            EnableAddInStepTwo();
        }

        private void StepTwoDoneButton_Click(object sender, RibbonControlEventArgs e)
        {
            progressDialog = new ProgressDialogBox();
            progressDialog.SetDialogBoxShownCallback(StepThreePostDialogShown);
            progressDialog.Show();      
        }

        public async void StepThreePostDialogShown()
        {
            progressDialog.TopMost = true;
            PowerPoint.Shape groupedShapes = StepThreeGroupShapes();
            await Task.Run(() => StepThreePopulateSelectedSlides(groupedShapes, progressDialog));
            EnableAddInStart(null);
            progressDialog.Close();

        }

        private void CleanupButton_Click(object sender, RibbonControlEventArgs e)
        {
            formatShapes.Clear();
            positionMarkers.Clear();
            positionTextBoxes.Clear();
            EnableAddInStart(null);
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

        private PowerPoint.Shape StepThreeGroupShapes()
        {
            Microsoft.Office.Core.MsoTriState selectionMode = Microsoft.Office.Core.MsoTriState.msoTrue;
            foreach (PowerPoint.Shape s in positionTextBoxes.Values)
            {
                s.Select(selectionMode);

                //Change selection mode to keep previous selections
                if (selectionMode == Microsoft.Office.Core.MsoTriState.msoTrue)
                    selectionMode = Microsoft.Office.Core.MsoTriState.msoFalse;
            }
            foreach (PowerPoint.Shape s in positionMarkers.Values)
            {
                s.Select(selectionMode);
            }


            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.ShapeRange selectedShapes = presentation.Windows[1].Selection.ShapeRange;
            PowerPoint.Shape groupedShapes = selectedShapes.Group();
            groupedShapes.Name = GROUPED_SHAPES;
            return groupedShapes;
        }

        private void StepThreePopulateSelectedSlides(PowerPoint.Shape groupedShapes, ProgressDialogBox progressDialog)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;

            int prevSection = 1;
            bool skippedFirst = false;
            int slidesProcessed = 0, totalSlides = slideNumbers.Count;
            foreach (int section in slidesPerSection.Keys)
            {
                foreach(int slideIndex in slidesPerSection[section])
                {
                    Thread.Sleep(100);
                    ++slidesProcessed;
                    progressDialog.UpdateProgressMessage(slidesProcessed, totalSlides);
                    if (!skippedFirst)
                    {
                        skippedFirst = !skippedFirst;
                        continue; //First slide is already done
                    }
                    groupedShapes.Copy();
                    PowerPoint.Slide curSlide = presentation.Slides[slideIndex];
                    IEnumerator enumerator = curSlide.Shapes.Paste().GetEnumerator();
                    enumerator.MoveNext();
                    groupedShapes = (PowerPoint.Shape)enumerator.Current;

                    if (section != prevSection)
                    {
                        PowerPoint.Shape inactiveTextBox = formatShapes[FORMAT_INACTIVE_SECTION_TEXT_BOX];
                        inactiveTextBox.PickUp();
                        PowerPoint.Shape textBox = Util.FindTextBoxFromGroup(groupedShapes, prevSection);
                        textBox.Apply();

                        PowerPoint.Shape activeTextBox = formatShapes[FORMAT_ACTIVE_SECTION_TEXT_BOX];
                        activeTextBox.PickUp();
                        textBox = Util.FindTextBoxFromGroup(groupedShapes, section);
                        textBox.Apply();

                        prevSection = section;
                    }

                    if (includeSlideMarkers)
                    {
                        UpdateMarkers(groupedShapes, section, slideIndex);
                    }
                }
            }
        }

        public void UpdateMarkers(PowerPoint.Shape groupedShapes, int section, int slideIndex)
        {
            PowerPoint.Shape currentSlideMarker = formatShapes[FORMAT_CURRENT_SLIDE_SLIDE_MARKER];
            PowerPoint.Shape activeSlideMarker = formatShapes[FORMAT_ACTIVE_SECTION_SLIDE_MARKER];
            PowerPoint.Shape inactiveSlideMarker = formatShapes[FORMAT_INACTIVE_SECTION_SLIDE_MARKER];

            int slideIndexWithinSection = Util.GetSlideIndexWithinSection(slidesPerSection, slideIndex);
            int prevSlide = slideNumbers[slideNumbers.IndexOf(slideIndex) - 1];

            foreach (PowerPoint.Shape s in groupedShapes.GroupItems)
            {
                //Not interested in section text boxes
                if (s.Name.StartsWith(POSITION_TEXT_BOX))
                    continue;

                int markerSection, markerSlideIndex;
                if(!Util.TryGetSlideAndSectionIndexFromMarkerName(s.Name, out markerSection, out markerSlideIndex))
                {
                    throw new AddinException("Error updating markers - TryGetSlideAndSectionIndexFromMarkerName");
                }

                //Handle marker formatting when section changes
                if (slideIndexWithinSection == 1 && section > 1)
                {
                    if(markerSection == section - 1)
                    {
                        inactiveSlideMarker.PickUp();
                        s.Apply();
                    }
                    else if (markerSection == section)
                    {
                        activeSlideMarker.PickUp();
                        s.Apply();
                    }
                }

                //Format active slide marker
                if(section == markerSection && slideIndex == markerSlideIndex)
                {
                    currentSlideMarker.PickUp();
                    s.Apply();
                }

                //Remove active marker from previous slide
                if(slideIndexWithinSection > 1)
                {
                    if(markerSlideIndex == prevSlide)
                    {
                        activeSlideMarker.PickUp();
                        s.Apply();
                    }
                }
            }
        }

        public void EnableAddInStart(PowerPoint.Presentation presentation)
        {
            slideMarkerCheckBox.Enabled = true;
            slideRangeEditBox.Enabled = true;
            startButton.Enabled = true;
            stepOneNextButton.Enabled = false;
            stepTwoDoneButton.Enabled = false;
            cleanPresentationButton.Enabled = true;
        }

        public void EnableAddInStepOne()
        {
            slideMarkerCheckBox.Enabled = false;
            slideRangeEditBox.Enabled = false;
            startButton.Enabled = false;
            stepOneNextButton.Enabled = true;
            stepTwoDoneButton.Enabled = false;
        }

        public void EnableAddInStepTwo()
        {
            slideMarkerCheckBox.Enabled = false;
            slideRangeEditBox.Enabled = false;
            startButton.Enabled = false;
            stepOneNextButton.Enabled = false;
            stepTwoDoneButton.Enabled = true;
        }

        public void DisableAddIn(PowerPoint.Presentation presentation)
        {
            slideMarkerCheckBox.Enabled = false;
            slideRangeEditBox.Enabled = false;
            startButton.Enabled = false;
            stepOneNextButton.Enabled = false;
            cleanPresentationButton.Enabled = false;
            stepTwoDoneButton.Enabled = false;
        }
    }
}
