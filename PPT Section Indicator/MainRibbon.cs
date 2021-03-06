﻿using System;
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
using System.Runtime.InteropServices;

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

        private const string ABOUT_MESSAGE =
            "PPT Section Indicator v1.0.2\n\n" +
            "Written by Fábio Muramatsu and released under the MIT License";
        private const string CLEANUP_MESSAGE =
            "Your presentation contains elements that need to be cleaned before proceeding. Would you like to clean them and proceed?\n\n" +
            "If you've run this tool before, press YES to proceed. However, if this is the first time you run the tool on this presentation, it is likely that " +
            "it contains elements that will conflict with this tool. Press NO, and read the documentation to remove those conflicts.";
        private const string ONE_SECTION_MESSAGE = "PPT Section Indicator requires at least two sections if \"Include slide markers\" is not selected";
        private const string INDEXES_CHANGED_MESSAGE = "Your presentation has changed while PPT Section Indicator was working. Please, restart the process.";
        private const string COM_EXCEPTION_MESSAGE =
            "Unexpected error. Did you delete any element generated by PPT Section Indicator? Please, restart the process.\n\n" +
            "The following error message was produced: ";
        private const string ADDIN_EXCEPTION_MESSAGE =
            "The add-in found an error and can't continue. Please, restart the process.\n\n" +
            "The following error message was produced: ";
        private const string START_BUTTON_INSTRUCTIONS =
            "In this step, PPT Section Indicator creates sample text boxes (and markers, depending on your settings). You should format them to your liking " +
            "(e.g., font size, color, fill, etc.). Don't worry about placement yet: this will be done on the next step. When you're done, press the Next button in the toolbar.\n\n" +
            "The shapes will be inserted in slide number {0}. For more details, please check the documentation.";
        private const string STEP_ONE_BUTTON_INSTRUCTIONS =
            "In this step, PPT Section Indicator creates text boxes (and markers) for your presentation, based on the formatting you provided on the previous step. " +
            "Now, you can customize where each element should be placed on your slides. In each section, markers are ordered from left to right, spawning a new line if necessary. " +
            "When you're done, press the Done button.\n\n" +
            "For more details, please check the documentation.";

        private const string DOC_URL = "https://github.com/fabio-muramatsu/PPTSectionIndicator/blob/master/README.md";

        private const int DEFAULT_SECTION_SPACING = 150;

        private Dictionary<String, PowerPoint.Shape> formatShapes = new Dictionary<string, PowerPoint.Shape>();
        private Dictionary<int, PowerPoint.Shape> positionTextBoxes = new Dictionary<int, PowerPoint.Shape>();
        private Dictionary<int, PowerPoint.Shape> positionMarkers = new Dictionary<int, PowerPoint.Shape>();

        bool includeSlideMarkers, includeHyperlinks;
        IList<int> slideNumbers, sectionNumbers;
        IDictionary<int, IList<int>> slidesPerSection;

        private ProgressDialogBox progressDialog;

        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.ThisAddIn.Application.AfterPresentationOpen += new PowerPoint.EApplication_AfterPresentationOpenEventHandler(EnableAddInStart);
            Globals.ThisAddIn.Application.AfterNewPresentation += new PowerPoint.EApplication_AfterNewPresentationEventHandler(EnableAddInStart);
            Globals.ThisAddIn.Application.PresentationClose += new PowerPoint.EApplication_PresentationCloseEventHandler(PresentationCloseCallback);

            EnableAddInStart(null);
        }

        private void StartButton_Click(object sender, RibbonControlEventArgs e)
        {

            PowerPoint.Presentation presentation;

            try
            {
                presentation = Globals.ThisAddIn.Application.ActivePresentation;
            }
            catch (COMException)
            {
                Util.ShowErrorMessage("There is no open presentation");
                return;
            }

            if (presentation.SectionProperties.Count == 0)
            {
                Util.ShowErrorMessage("There are no sections in the presentation");
                return;
            }

            includeSlideMarkers = slideMarkerCheckBox.Checked;
            includeHyperlinks = hyperlinkCheckBox.Checked;

            if(Util.GetCleanupItems().Count > 0)
            {
                DialogResult result = Util.ShowWarningQuery(CLEANUP_MESSAGE);
                if (result == DialogResult.Yes)
                {
                    cleanupPresentation();
                }
                else return;
            }

            try
            {
                slideNumbers = Util.GetSlidesFromRangeExpr(slideRangeEditBox.Text);
                slidesPerSection = Util.ClassifySlidesIntoSections(slideNumbers);
                sectionNumbers = slidesPerSection.Keys.OrderBy(i => i).ToList();
            }
            catch (Exception exc) when(exc is SlideRangeFormatException || exc is SlideOutOfRangeException || exc is NoSectionException)
            {
                Util.ShowErrorMessage(exc.Message);
                return;
            }

            if(!includeSlideMarkers && slidesPerSection.Keys.Count < 2)
            {
                Util.ShowErrorMessage(ONE_SECTION_MESSAGE);
                return;
            }

            if (!Util.CheckSlideRange(slideNumbers))
            {
                Util.ShowErrorMessage("Specified range exceeds number of slides");
                return;
            }

            if (Properties.Settings.Default.ShowStartButtonInstructions)
            {
                bool dontShowMessage = new MessageCheckboxDialog(String.Format(START_BUTTON_INSTRUCTIONS, slideNumbers.First())).ShowDialogForResult();
                Properties.Settings.Default.ShowStartButtonInstructions = !dontShowMessage;
                Properties.Settings.Default.Save();
            }

            PowerPoint.Slide firstSlide = presentation.Slides[slideNumbers.First()];
            firstSlide.Select();

            StepOneInsertFormatPlaceholders(firstSlide);
            EnableAddInStepOne();
        }

        private void StepOneNextButton_Click(object sender, RibbonControlEventArgs e)
        {
            if(!Util.CheckPresentationIndexesUnchanged(slideNumbers, slidesPerSection))
            {
                Util.ShowErrorMessage(INDEXES_CHANGED_MESSAGE);
                cleanupPresentation();
                EnableAddInStart(null);
                return;
            }

            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slide firstSlide = presentation.Slides[slideNumbers.First()];

            int currentSection = -1, previousSection = -1;
            try
            {
                previousSection = Util.GetSectionIndex(slideNumbers.First());
            }
            catch (NoSectionException nse)
            {
                Util.ShowErrorMessage(nse.Message);
                return;
            }

            try
            {
                foreach (int slideIndex in slideNumbers)
                {
                    if (currentSection == -1)
                    {
                        StepTwoInsertTextBoxes(previousSection);
                    }
                    currentSection = Util.GetSectionIndex(slideIndex);
                    if (currentSection != previousSection)
                    {
                        StepTwoInsertTextBoxes(currentSection);
                    }

                    if (includeSlideMarkers)
                        StepTwoInsertSlideMarkers(currentSection, slideIndex);

                    previousSection = currentSection;
                }

                foreach (PowerPoint.Shape shape in formatShapes.Values)
                {
                    shape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                }
            }
            catch(COMException exc)
            {
                Util.ShowErrorMessage(COM_EXCEPTION_MESSAGE + exc.Message);
                cleanupPresentation();
                EnableAddInStart(null);
                return;
            }

            if (Properties.Settings.Default.ShowFirstStepButtonInstructions)
            {
                bool dontShowMessage = new MessageCheckboxDialog(STEP_ONE_BUTTON_INSTRUCTIONS).ShowDialogForResult();
                Properties.Settings.Default.ShowFirstStepButtonInstructions = !dontShowMessage;
                Properties.Settings.Default.Save();
            }

            EnableAddInStepTwo();
        }

        private void StepTwoDoneButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Util.CheckPresentationIndexesUnchanged(slideNumbers, slidesPerSection))
            {
                Util.ShowErrorMessage(INDEXES_CHANGED_MESSAGE);
                cleanupPresentation();
                EnableAddInStart(null);
                return;
            }

            progressDialog = new ProgressDialogBox();
            progressDialog.SetDialogBoxShownCallback(StepThreePostDialogShown);
            progressDialog.Show();      
        }

        private void StepOneAboutButton_Click(object sender, RibbonControlEventArgs e)
        {
            bool showInstructions = Properties.Settings.Default.ShowStartButtonInstructions;
            MessageCheckboxDialog dialog = new MessageCheckboxDialog(String.Format(START_BUTTON_INSTRUCTIONS, slideNumbers.First()));
            dialog.SetCheckBoxState(!showInstructions);
            Properties.Settings.Default.ShowStartButtonInstructions = !dialog.ShowDialogForResult();
            Properties.Settings.Default.Save();
        }

        private void StepTwoAboutButton_Click(object sender, RibbonControlEventArgs e)
        {
            bool showInstructions = Properties.Settings.Default.ShowFirstStepButtonInstructions;
            MessageCheckboxDialog dialog = new MessageCheckboxDialog(STEP_ONE_BUTTON_INSTRUCTIONS);
            dialog.SetCheckBoxState(!showInstructions);
            Properties.Settings.Default.ShowFirstStepButtonInstructions = !dialog.ShowDialogForResult();
            Properties.Settings.Default.Save();
        }


        private void CleanupButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                cleanupPresentation();
            }
            catch (NoActivePresentation exc)
            {
                Util.ShowErrorMessage(exc.Message);
            }
            EnableAddInStart(null);
        }

        private void ShowDocumentationButton_Click(object sender, RibbonControlEventArgs e)
        {
            Process.Start(DOC_URL);
        }

        private void AboutButton_Click(object sender, RibbonControlEventArgs e)
        {
            Util.ShowMessage(ABOUT_MESSAGE);
        }

        public async void StepThreePostDialogShown()
        {
            progressDialog.TopMost = true;
            try
            {
                PowerPoint.Shape groupedShapes = StepThreeGroupShapes();
                if (includeHyperlinks)
                {
                    StepThreeInsertHyperlinks(groupedShapes);
                }
                await Task.Run(() => StepThreePopulateSelectedSlides(groupedShapes, progressDialog));
                EnableAddInStart(null);

                foreach (PowerPoint.Shape s in formatShapes.Values)
                {
                    s.Delete();
                }
            }
            catch (COMException exc)
            {
                Util.ShowErrorMessage(COM_EXCEPTION_MESSAGE + exc.Message);
                cleanupPresentation();
                EnableAddInStart(null);
                return;
            }
            catch (AddinException exc)
            {
                Util.ShowErrorMessage(ADDIN_EXCEPTION_MESSAGE + exc.Message);
                cleanupPresentation();
                EnableAddInStart(null);
                return;
            }
            finally
            {
                progressDialog.Close();
            }
        }


        private void cleanupPresentation()
        {
            formatShapes.Clear();
            positionMarkers.Clear();
            positionTextBoxes.Clear();
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
            if (section == sectionNumbers.First())
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
                int effectiveSection = sectionNumbers.IndexOf(section) + 1;
                PowerPoint.Shape textBox;
                formatShapes.TryGetValue(FORMAT_INACTIVE_SECTION_TEXT_BOX, out textBox);
                textBox.Copy();
                IEnumerator enumerator = firstSlide.Shapes.Paste().GetEnumerator();
                enumerator.MoveNext();
                PowerPoint.Shape newTextBox = (PowerPoint.Shape)enumerator.Current;
                newTextBox.Left = DEFAULT_SECTION_SPACING * (effectiveSection - 1) + 10;
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
            else if (section == sectionNumbers.First())
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

            int effectiveSection = sectionNumbers.IndexOf(section) + 1;
            int slideIndexWithinSection = Util.GetSlideIndexWithinSection(slidesPerSection, slideIndex);
            float left = 18 + (effectiveSection - 1) * DEFAULT_SECTION_SPACING + ((slideIndexWithinSection - 1) % maxNumberOfMarkers) * (newMarker.Width + 2);
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

        private void StepThreeInsertHyperlinks(PowerPoint.Shape groupedElements)
        {
            int sectionIndex, slideIndex;
            foreach(PowerPoint.Shape s in groupedElements.GroupItems)
            {
                if (s.Name.StartsWith(POSITION_TEXT_BOX))
                {
                    if (!Util.TryGetSectionIndexFromTextboxName(s.Name, out sectionIndex))
                    {
                        throw new AddinException("TryGetSectionIndexFromTextboxName - Failed to parse text box name");
                    }
                    s.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Action = PowerPoint.PpActionType.ppActionHyperlink;
                    s.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.SubAddress = slidesPerSection[sectionIndex].First().ToString();
                }
                else if (s.Name.StartsWith(POSITION_SLIDE_MARKER))
                {
                    if (!Util.TryGetSlideAndSectionIndexFromMarkerName(s.Name, out sectionIndex, out slideIndex))
                    {
                        throw new AddinException("TryGetSlideAndSectionIndexFromMarkerName - Failed to parse marker name");
                    }
                    s.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Action = PowerPoint.PpActionType.ppActionHyperlink;
                    s.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.SubAddress = slideIndex.ToString();
                }
            }
        }

        private void StepThreePopulateSelectedSlides(PowerPoint.Shape groupedShapes, ProgressDialogBox progressDialog)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;

            int prevSection = sectionNumbers.First();
            bool skippedFirst = false;
            int slidesProcessed = 0, totalSlides = slideNumbers.Count;
            foreach (int section in slidesPerSection.Keys)
            {
                foreach(int slideIndex in slidesPerSection[section])
                {
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

        private void UpdateMarkers(PowerPoint.Shape groupedShapes, int section, int slideIndex)
        {
            PowerPoint.Shape currentSlideMarker = formatShapes[FORMAT_CURRENT_SLIDE_SLIDE_MARKER];
            PowerPoint.Shape activeSlideMarker = formatShapes[FORMAT_ACTIVE_SECTION_SLIDE_MARKER];
            PowerPoint.Shape inactiveSlideMarker = formatShapes[FORMAT_INACTIVE_SECTION_SLIDE_MARKER];

            int slideIndexWithinSection = Util.GetSlideIndexWithinSection(slidesPerSection, slideIndex);
            int prevSlide = slideNumbers[slideNumbers.IndexOf(slideIndex) - 1];

            int effectiveSection = sectionNumbers.IndexOf(section) - 1;
            int prevSection = effectiveSection == -1? -1 : sectionNumbers[effectiveSection];

            foreach (PowerPoint.Shape s in groupedShapes.GroupItems)
            {
                //Not interested in section text boxes
                if (s.Name.StartsWith(POSITION_TEXT_BOX))
                    continue;

                int markerSection, markerSlideIndex;
                if(!Util.TryGetSlideAndSectionIndexFromMarkerName(s.Name, out markerSection, out markerSlideIndex))
                {
                    throw new AddinException("TryGetSlideAndSectionIndexFromMarkerName - Error updating markers");
                }

                //Handle marker formatting when section changes
                if (slideIndexWithinSection == 1 && section > 1)
                {
                    if(markerSection == prevSection)
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
            hyperlinkCheckBox.Enabled = true;
            slideRangeEditBox.Enabled = true;
            startButton.Enabled = true;
            stepOneNextButton.Enabled = false;
            stepTwoDoneButton.Enabled = false;
            cleanPresentationButton.Enabled = true;
            stepOneAboutButton.Enabled = false;
            stepTwoAboutButton.Enabled = false;
        }

        public void EnableAddInStepOne()
        {
            slideMarkerCheckBox.Enabled = false;
            hyperlinkCheckBox.Enabled = false;
            slideRangeEditBox.Enabled = false;
            startButton.Enabled = false;
            stepOneNextButton.Enabled = true;
            stepTwoDoneButton.Enabled = false;
            stepOneAboutButton.Enabled = true;
            stepTwoAboutButton.Enabled = false;
        }

        public void EnableAddInStepTwo()
        {
            slideMarkerCheckBox.Enabled = false;
            hyperlinkCheckBox.Enabled = false;
            slideRangeEditBox.Enabled = false;
            startButton.Enabled = false;
            stepOneNextButton.Enabled = false;
            stepTwoDoneButton.Enabled = true;
            stepOneAboutButton.Enabled = false;
            stepTwoAboutButton.Enabled = true;
        }

        public void PresentationCloseCallback(PowerPoint.Presentation presentation)
        {
            EnableAddInStart(presentation);
            formatShapes.Clear();
            positionMarkers.Clear();
            positionTextBoxes.Clear();
        }
    }
}
