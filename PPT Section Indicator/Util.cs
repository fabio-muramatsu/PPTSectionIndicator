﻿using System;
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
        public static IEnumerable<int> GetSlidesFromRangeExpr(string expression)
        {
            SortedSet<int> slides = new SortedSet<int>();
            if (CheckPageRangeSyntax(expression))
            {
                string[] slideRanges = expression.Trim().Split(';');
                foreach(string range in slideRanges)
                {
                    string[] slideNumbers = range.Trim().Split('-');
                    if (slideNumbers.Length == 1)
                        slides.Add(int.Parse(slideNumbers[0]));
                    else
                    {
                        int min = int.Parse(slideNumbers[0]);
                        int max = int.Parse(slideNumbers[1]);
                        if(max < min)
                        {
                            throw new SlideRangeFormatException("Wrong range format: left-hand side should be no grater than right-hand side");
                        }
                        else
                        {
                            slides.UnionWith(Enumerable.Range(min, max - min + 1));
                        }
                    }
                }

                return slides;
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

            if(slides.Last() > presentation.Slides.Count)
            {
                throw new SlideOutOfRangeException("Specified slide range exceeds the slide number in you presentation");
            }
            foreach(int slideIndex in slides)
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
        public static string GetSectionName (int slideIndex)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.SectionProperties sections = presentation.SectionProperties;
            try
            {
                return sections.Name(GetSectionIndex(slideIndex));
            }
            catch(NoSectionException e)
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
            
            foreach(int key in slidesPerSection.Keys)
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

        public static IEnumerable<PowerPoint.Shape> GetCleanupItems()
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            LinkedList<PowerPoint.Shape> matches = new LinkedList<PowerPoint.Shape>();
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                foreach(PowerPoint.Shape shape in slide.Shapes)
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


    }

    //public static class CustomExtensions
    //{
    //    public static V GetValue<K,V>(this IDictionary<K,V> dict, K key)
    //    {
    //        V value;
    //        bool result = dict.TryGetValue(key, out value);
    //        if (result)
    //            return value;
    //        else

    //    }
    //}

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
}
