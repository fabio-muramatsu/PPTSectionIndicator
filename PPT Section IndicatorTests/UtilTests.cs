using Microsoft.VisualStudio.TestTools.UnitTesting;
using PPT_Section_Indicator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPT_Section_Indicator.Tests
{
    [TestClass()]
    public class UtilTests
    {
        [TestMethod()]
        public void CheckPageRangeSyntaxTest()
        {
            Assert.IsTrue(Util.CheckPageRangeSyntax("1"));
            Assert.IsTrue(Util.CheckPageRangeSyntax("1-3"));
            Assert.IsTrue(Util.CheckPageRangeSyntax("1;2"));
            Assert.IsTrue(Util.CheckPageRangeSyntax("1 ;2-4"));
            Assert.IsTrue(Util.CheckPageRangeSyntax("1- 2; 4  "));

            Assert.IsFalse(Util.CheckPageRangeSyntax("1-; 4  "));
            Assert.IsFalse(Util.CheckPageRangeSyntax("a"));
            Assert.IsFalse(Util.CheckPageRangeSyntax("-1-3; 4  "));
            Assert.IsFalse(Util.CheckPageRangeSyntax("-"));
            Assert.IsFalse(Util.CheckPageRangeSyntax(""));
            Assert.IsFalse(Util.CheckPageRangeSyntax("  "));
            Assert.IsFalse(Util.CheckPageRangeSyntax("-2"));
        }

        [TestMethod()]
        public void GetSlidesFromRangeExprTest()
        {
            IEnumerable<int> test1 = new List<int> { 1, 2, 3 };
            IEnumerable<int> test2 = new List<int> { 1, 2, 3, 6, 7, 8 };

            Assert.IsTrue(Enumerable.SequenceEqual<int>(test1, Util.GetSlidesFromRangeExpr("1-3")));
            Assert.IsTrue(Enumerable.SequenceEqual<int>(test1, Util.GetSlidesFromRangeExpr("1;2;3")));
            Assert.IsTrue(Enumerable.SequenceEqual<int>(test1, Util.GetSlidesFromRangeExpr("1;2-3")));
            Assert.IsTrue(Enumerable.SequenceEqual<int>(test1, Util.GetSlidesFromRangeExpr("1-1;2-3")));
            Assert.IsTrue(Enumerable.SequenceEqual<int>(test2, Util.GetSlidesFromRangeExpr("1-3;6-8")));
            Assert.IsTrue(Enumerable.SequenceEqual<int>(test2, Util.GetSlidesFromRangeExpr("1 ;2 ; 3; 2-3 ; 6 - 8")));
        }
    }
}