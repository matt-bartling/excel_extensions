using NUnit.Framework;

namespace Mb.ExcelExtensions.Tests
{
    [TestFixture]
    public class TestSorting
    {
        [Test]
        public void TestSortSimpleStrings()
        {
            var array = new[] {new object[]{"x"}, new object[]{"c"}, new object[]{"r"}};
            var sorted = Sorter.Sort(array, new[] { new SortParam { Col = 1 } });
            Assert.AreEqual("c", sorted[0][0]);
        }

        [Test]
        public void TestSortSimpleStringsDescending()
        {
            var array = new[] { new object[] { "x" }, new object[] { "c" }, new object[] { "r" } };
            var sorted = Sorter.Sort(array, new[] { new SortParam { Col = 1, SortDirection = SortDirection.Descending } });
            Assert.AreEqual("x", sorted[0][0]);
            Assert.AreEqual("r", sorted[1][0]);
            Assert.AreEqual("c", sorted[2][0]);
        }

        [Test]
        public void TestSortSimpleStringsDescendingWithNulls()
        {
            var array = new[] { new object[] { "x" }, new object[] { null }, new object[] { "r" } };
            var sorted = Sorter.Sort(array, new[] { new SortParam { Col = 1, SortDirection = SortDirection.Descending } });
            Assert.AreEqual("x", sorted[0][0]);
            Assert.AreEqual("r", sorted[1][0]);
            Assert.AreEqual(null, sorted[2][0]);
        }

        [Test]
        public void TestSortSimpleStringsWithNulls()
        {
            var array = new[] { new object[] { "x" }, new object[] { null }, new object[] { "r" } };
            var sorted = Sorter.Sort(array, new[] { new SortParam { Col = 1 } });
            Assert.AreEqual("r", sorted[0][0]);
            Assert.AreEqual("x", sorted[1][0]);
            Assert.AreEqual(null, sorted[2][0]);
        }

        [Test]
        public void TestSortMultipleStrings()
        {
            var array = new[]
            {
                new object[] {"a", "e", "x"},
                new object[] {"b", "q", "6"},
                new object[] {"a", "c", "y"}
            };

            var sortParams = new[]
            {
                new SortParam{Col=1},
                new SortParam{Col=2},
            };
            var sorted = Sorter.Sort(array, sortParams);
            Assert.AreEqual("a", sorted[0][0]);
            Assert.AreEqual("c", sorted[0][1]);
            Assert.AreEqual("a", sorted[1][0]);
            Assert.AreEqual("e", sorted[1][1]);
        }

        [Test]
        public void TestSortNumbers()
        {
            var array = new[]
            {
                new object[] {1, 2.3, 6},
                new object[] {700, 2, 1.1},
                new object[] {3, 1, 2.2},
            };
            var sortParams = new[]
            {
                new SortParam{Col=1},
            };
            var sorted = Sorter.Sort(array, sortParams);
            Assert.AreEqual(1, sorted[0][0]);
            Assert.AreEqual(3, sorted[1][0]);
            Assert.AreEqual(700, sorted[2][0]);
        }

        [Test]
        public void TestSortNumbersWithNulls()
        {
            var array = new[]
            {
                new object[] {(double)1, 2.3, 6},
                new object[] {null, 2, 1.1},
                new object[] {(double)3, 1, 2.2},
            };
            var sortParams = new[]
            {
                new SortParam{Col=1},
            };
            var sorted = Sorter.Sort(array, sortParams);
            Assert.AreEqual(1, sorted[0][0]);
            Assert.AreEqual(3, sorted[1][0]);
            Assert.AreEqual(null, sorted[2][0]);
        }

    }
}
