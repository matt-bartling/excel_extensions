using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace Mb.ExcelExtensions
{
    public class SortFunctions
    {
        [ExcelFunction]
        public static object[,] VUnique(object[,] data)
        {
            var result = CreateArray(data.GetLength(0), data.GetLength(1));
            var unique = new Dictionary<string, string>();
            var rowCount = 0;
            for (var i = 0; i < data.GetLength(0); i++)
            {
                var key = ToKey(data, i);
                if (!string.IsNullOrEmpty(key) && !unique.ContainsKey(key))
                {
                    unique[key] = key;
                    for (var j = 0; j < data.GetLength(1); j++)
                    {
                        result[rowCount, j] = data[i, j];
                    }
                    rowCount++;
                }
            }

            return result;
        }
        [ExcelFunction]
        public static object[,] SubRange(object[,] data, int startCol, int startRow, int endCol, int endRow)
        {
            var endRowIndex = Math.Min(endRow - 1, data.GetLength(0) - 1);
            var endColIndex = Math.Min(endCol - 1, data.GetLength(1) - 1);
            var startRowIndex = startRow - 1;
            var startColIndex = startCol - 1;
            var array = new object[endRowIndex - startRowIndex + 1, endColIndex - startColIndex + 1];
            for (var i = startRowIndex; i <= endRowIndex; i++)
            {
                for (var j = startColIndex; j <= endColIndex; j++)
                {
                    array[i - startRowIndex, j - startColIndex] = data[i, j];
                }
            }
            return array;
        }
        [ExcelFunction]
        public static object[,] VSort(object[,] data, object[,] sortParams, bool emptyLast = true)
        {
            var r = ToRows(data);
            var sortParamArray = ToSortParams(sortParams);
            var sorted = Sorter.Sort(r, sortParamArray, emptyLast);
            return To2DArray(sorted);
        }

        private static SortParam[] ToSortParams(object[,] sortParams)
        {
            var sortParamList = new List<SortParam>();
            for (var i = 0; i < sortParams.GetLength(0); i++)
            {
                if (sortParams[i,0] is double)
                {
                    sortParamList.Add(new SortParam{Col = Convert.ToInt32(sortParams[i, 0])});
                }
            }
            return sortParamList.ToArray();
        }

        [ExcelFunction]
        public static object[,] VFilter(object[,] data, object[,] filterParams, bool negativeMatch = false)
        {
            var results = CreateArray(data.GetLength(0), data.GetLength(1));
            var resultCount = 0;
            for (var i = 0; i < data.GetLength(0); i++)
            {
                if (MatchesFilter(data, i, filterParams, negativeMatch))
                {
                    for (var j = 0; j < data.GetLength(1); j++)
                    {
                        results[resultCount, j] = data[i, j];
                    }
                    resultCount++;
                }
            }
            return results;
        }

        private static bool MatchesFilter(object[,] dataTable, int dataRow, object[,] filterTable, bool negativeMatch)
        {
            for (var j = 0; j < filterTable.GetLength(0); j++)
            {
                var matchesFilter = MatchesFilter(dataTable, dataRow, filterTable, j);
                if ((!matchesFilter && !negativeMatch) ||
                    (matchesFilter && negativeMatch))
                {
                    return false;
                }
            }
            return true;
        }

        private static bool MatchesFilter(object[,] dataTable, int dataRow, object[,] filterTable, int filterRow)
        {
            if (filterTable[filterRow, 0] is double)
            {
                var col = (int)(double)filterTable[filterRow, 0] - 1;
                if (filterTable[filterRow, 1] is ExcelEmpty)
                {
                    return dataTable[dataRow, col].Equals(filterTable[filterRow, 1]);
                }
                else
                {
                    return dataTable[dataRow, col].ToString().CompareTo(filterTable[filterRow, 1].ToString()) > 0 && dataTable[dataRow, col].ToString().CompareTo(filterTable[filterRow, 2].ToString()) < 0;
                }
            }
            return true;
        }

        [ExcelFunction]
        public static object[,] VJoin(object[,] table1, object[,] table2)
        {
            var joined = new object[table1.GetLength(0) + table2.GetLength(0), Math.Max(table1.GetLength(1), table2.GetLength(1))];

            var c = 0;

            foreach (var table in new[] { table1, table2 })
            {
                for (var i = 0; i < table.GetLength(0); i++)
                {
                    for (var j = 0; j < table.GetLength(1); j++)
                    {
                        joined[c, j] = table[i, j] is ExcelEmpty ? "" : table[i, j];
                    }
                    c++;
                }
            }

            return joined;
        }

        [ExcelFunction]
        public static object[,] Arr(double x) 
        {
            return new object[,] { { x, x, x }, { x, x, x }, { x, x, x } };

        }

        [ExcelFunction]
        public static double Sqr(double val)
        {
            return val * val;
        }

        private static object[,] To2DArray(object[][] multiArray)
        {
            var rows = multiArray.Length;
            var cols = multiArray.Select(a => a.Length).Max();

            var returnArray = new object[rows, cols];
            for (var i = 0; i < multiArray.Length; i++)
            {
                for (var j = 0; j < multiArray[i].Length; j++)
                {
                    returnArray[i, j] = multiArray[i][j];
                }
            }
            return returnArray;
        }

        private static object[][] ToRows(object[,] table)
        {
            var rows = new object[table.GetLength(0)][];
            for (var i = 0; i < table.GetLength(0); i++)
            {
                var row = new object[table.GetLength(1)];
                for (var j = 0; j < table.GetLength(1); j++)
                {
                    row[j] = table[i, j] is ExcelEmpty
                        ? ""
                        : table[i, j] is ExcelError ? null : table[i, j];
                }
                rows[i] = row;
            }
            return rows;
        }

        private static object[,] CreateArray(int rows, int cols) {
            var array = new object[rows, cols];
            for (var i = 0; i < rows; i++)
            {
                for (var j = 0; j < cols; j++)
                {
                    array[i, j] = ExcelError.ExcelErrorNA;
                }
            }
            return array;
        }

        private static string ToKey(object[,] data, int row)
        {
            object[] keyArray = new object[data.GetLength(1)];
            for (var j = 0; j < data.GetLength(1); j++)
            {
                keyArray[j] = data[row, j];
            }
            return string.Join(",", keyArray);
        }
    }
}
