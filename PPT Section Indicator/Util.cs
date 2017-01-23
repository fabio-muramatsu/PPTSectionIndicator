using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPT_Section_Indicator
{
    public class Util
    {
        private const string SHAPE_NAME_PREFIX = "SectionIndicator";

        /// <summary>
        /// Checks if the slide range expression is valid.
        /// </summary>
        /// <param name="input">The slide range expression.</param>
        /// <returns>true if the expression is valid, false otherwise</returns>
        public static bool CheckPageRangeSyntax(string input)
        {
            bool isMatch = Regex.Match(input, @"^\s*\d+(\s*-\s*\d+)?(\s*;\s*\d+(\s*-\s*\d+)?)*\s*$").Success;
            Debug.WriteLine(isMatch ? input + "is valid" : input + "is not valid");
            return isMatch;
        }

        /// <summary>
        /// Returns an IEnumerable object contaning the slide numbers corresponding to the range expression. The output is sorted in ascending order.
        /// </summary>
        /// <param name="expression">The slide range expression.</param>
        /// <returns>An IEnumerable object containing the slide numbers sorted in ascending order.</returns>
        /// <exception cref="SlideRangeFormatException">Thrown when there is an error with the expression provided.</exception>
        public static IList<int> GetSlidesFromRangeExpr(string expression)
        {
            SortedSet<int> slides = new SortedSet<int>();
            if (CheckPageRangeSyntax(expression))
            {
                string[] slideRanges = expression.Trim().Split(';');
                foreach (string range in slideRanges)
                {
                    string[] slideNumbers = range.Trim().Split('-');
                    if (slideNumbers.Length == 1)
                        slides.Add(int.Parse(slideNumbers[0]));
                    else
                    {
                        int min = int.Parse(slideNumbers[0]);
                        int max = int.Parse(slideNumbers[1]);
                        if (max < min)
                        {
                            throw new SlideRangeFormatException("Wrong range format: left-hand side should be no grater than right-hand side");
                        }
                        else
                        {
                            slides.UnionWith(Enumerable.Range(min, max - min + 1));
                        }
                    }
                }

                return new List<int>(slides);
            }
            else
            {
                throw new SlideRangeFormatException("Invalid slide range input format");
            }
        }

        public static IDictionary<int, IList<int>> ClassifySlidesIntoSections(IEnumerable<int> slides)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.SectionProperties sections = presentation.SectionProperties;

            Dictionary<int, IList<int>> slidesPerIndex = new Dictionary<int, IList<int>>();

            if (slides.Last() > presentation.Slides.Count)
            {
                throw new SlideOutOfRangeException("Specified slide range exceeds the slide number in you presentation");
            }
            foreach (int slideIndex in slides)
            {
                int section = GetSectionIndex(slideIndex);
                if (slidesPerIndex.ContainsKey(section))
                    slidesPerIndex[section].Add(slideIndex);
                else
                    slidesPerIndex[section] = new List<int> { slideIndex };
            }
            return slidesPerIndex;
        }

        /// <summary>
        /// Gets section name for a given slide index.
        /// </summary>
        /// <param name="slideIndex">The index of the slide whose section name is to be obtained.</param>
        /// <returns>The section name where the slide is contained.</returns>
        public static string GetSectionName(int slideIndex)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.SectionProperties sections = presentation.SectionProperties;
            try
            {
                return sections.Name(GetSectionIndex(slideIndex));
            }
            catch (NoSectionException e)
            {
                throw e;
            }
        }

        public static int GetSectionIndex(int slideIndex)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.SectionProperties sections = presentation.SectionProperties;
            if (sections.Count == 0)
            {
                throw new NoSectionException("Presentation has no sections");
            }
            else
            {
                int sectionIndex = 1;
                for (; sectionIndex <= sections.Count; ++sectionIndex)
                {
                    if (sections.FirstSlide(sectionIndex) <= slideIndex &&
                        sections.FirstSlide(sectionIndex) + sections.SlidesCount(sectionIndex) - 1 >= slideIndex)
                        break;
                }
                return sectionIndex;
            }
        }

        public static int GetSlideIndexWithinSection(IDictionary<int, IList<int>> slidesPerSection, int slideIndex)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.SectionProperties sections = presentation.SectionProperties;

            foreach (int key in slidesPerSection.Keys)
            {
                IList<int> slides = slidesPerSection[key];
                if (slides.First() <= slideIndex && slides.Last() >= slideIndex)
                {
                    return slides.IndexOf(slideIndex) + 1;
                }

            }
            return -1;
        }

        public static void ShowErrorMessage(String message)
        {
            MessageBox.Show(message, "PPT Section Indicator", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static DialogResult ShowWarningQuery(String message)
        {
            return MessageBox.Show(message, "PPT Section Indicator", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
        }

        public static IEnumerable<PowerPoint.Shape> GetCleanupItems()
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            LinkedList<PowerPoint.Shape> matches = new LinkedList<PowerPoint.Shape>();
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Name.StartsWith(SHAPE_NAME_PREFIX))
                        matches.AddLast(shape);
                }
            }
            return matches;
        }

        public static void CleanupShapes()
        {
            foreach (PowerPoint.Shape shape in GetCleanupItems())
            {
                shape.Delete();
            }
        }

        public static PowerPoint.Shape FindTextBoxFromGroup(PowerPoint.Shape groupedShape, int section)
        {
            string name = MainRibbon.POSITION_TEXT_BOX + "_" + section;
            foreach (PowerPoint.Shape s in groupedShape.GroupItems)
            {
                if (s.Name.Equals(name))
                    return s;
            }
            return null;
        }

        /// <summary>
        /// Gets the slide and section indexes from the marker name.
        /// </summary>
        /// <param name="markerName">The string representing the marker name.</param>
        /// <param name="section">The section index corresponding to the marker name.</param>
        /// <param name="slideIndex">The slide index corresponding to the marker name.</param>
        /// <returns>True if the processing was successful, False otherwise.</returns>
        public static bool TryGetSlideAndSectionIndexFromMarkerName
            (string markerName, out int section, out int slideIndex)
        {
            section = -1;
            slideIndex = -1;
            if (!markerName.StartsWith(MainRibbon.POSITION_SLIDE_MARKER))
                return false;
            string[] parts = markerName.Split('_');

            try
            {
                section = int.Parse(parts[parts.Length - 2]);
                slideIndex = int.Parse(parts[parts.Length - 1]);
                return true;
            }
            catch (Exception e) when (e is ArgumentNullException || e is FormatException || e is OverflowException)
            {
                return false;
            }
        }

        /// <summary>
        /// Checks if the specified slide range is consistent with the number of slides in the presentation.
        /// </summary>
        /// <param name="slideNumbers">The array containing the sorted slide indexes to be included in the processing.</param>
        /// <returns>True if the range is consistent, False otherwise.</returns>
        public static bool CheckSlideRange(IList<int> slideNumbers)
        {
            int numberOfSlides = Globals.ThisAddIn.Application.ActivePresentation.Slides.Count;
            if (slideNumbers.Last() > numberOfSlides) return false;
            else return true;
        }

        public static bool IsPresentationClean()
        {
            if (GetCleanupItems().Count() > 0)
                return false;
            else return true;
        }

        /// <summary>
        /// Checks if the specified slide indexes are still valid in the presentation, and if their division within section are as specified in slidesPerSection.
        /// </summary>
        /// <param name="slideNumbers">The slide indexes to be checked.</param>
        /// <param name="slidesPerSection">The original division in sections to be compared.</param>
        /// <returns></returns>
        public static bool CheckPresentationIndexesUnchanged(IList<int> slideNumbers, IDictionary<int, IList<int>> slidesPerSection)
        {
            try
            {
                IDictionary<int, IList<int>> currentSlidesPerSection = ClassifySlidesIntoSections(slideNumbers);

                foreach (int key in slidesPerSection.Keys)
                {
                    if (!Enumerable.SequenceEqual(slidesPerSection[key], currentSlidesPerSection[key]))
                        return false;
                }
            }
            catch (Exception exc) when (exc is SlideRangeFormatException || exc is SlideOutOfRangeException || exc is KeyNotFoundException)
            {
                return false;
            }

            return true;
        }

    }

    class SlideRangeFormatException : Exception
    {
        public SlideRangeFormatException(string message) : base(message)
        {
        }
    }

    class NoSectionException : Exception
    {
        public NoSectionException(string message) : base(message)
        {
        }
    }

    class SlideOutOfRangeException : Exception
    {
        public SlideOutOfRangeException(string message) : base(message)
        {
        }
    }

    public class AddinException : Exception
    {
        public AddinException(string message) : base(message)
        {
        }
    }
}
