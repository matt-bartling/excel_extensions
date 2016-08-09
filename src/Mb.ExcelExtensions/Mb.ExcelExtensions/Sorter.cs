using System.Linq;

namespace Mb.ExcelExtensions
{
    public class Sorter
    {
        public static object[][] Sort(object[][] data, SortParam[] sortParams, bool emptyLast = true)
        {
            var sorted = data.OrderBy(x => 1);
            foreach (var sortParam in sortParams)
            {
                var p = sortParam;
                var defaultMin = data.All(x => IsNumberOrNull(x, p, emptyLast)) ? (object)double.MaxValue : "zzzzzzzzzzzzzzzzzzzzzzzzz";
                var defaultMax = data.All(x => IsNumberOrNull(x, p, emptyLast)) ? (object)double.MinValue : "a";
                if (p.SortDirection == SortDirection.Ascending)
                {
                    sorted = sorted.ThenBy(x => SortValue(x, p, defaultMin));
                }
                else if (p.SortDirection == SortDirection.Descending)
                {
                    sorted = sorted.ThenByDescending(x => SortValue(x, p, defaultMax));
                }
            }
            return sorted.ToArray();
        }

        private static object SortValue(object[] x, SortParam p, object defaultValue)
        {
            if (x[p.Col - 1] == null || (x[p.Col-1] is string && (string)x[p.Col-1] == string.Empty))
            {
                return defaultValue;
            }
            return x[p.Col - 1];
        }

        private static bool IsNumberOrNull(object[] x, SortParam p, bool emptyLast)
        {
            return x[p.Col - 1] == null ||
                   (emptyLast && x[p.Col - 1] is string && (string) x[p.Col - 1] == string.Empty) ||
                   x[p.Col - 1] is double ||
                   x[p.Col - 1] is int;
        }
    }

    public enum SortDirection
    {
        Ascending,
        Descending,
    }

    public class SortParam
    {
        public int Col;
        public SortDirection SortDirection = SortDirection.Ascending;   
    }
}
