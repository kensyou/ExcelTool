using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.Helper
{
    public static class LinqExtensions
    {
        public static IEnumerable<TSource> Except<TSource>(this IEnumerable<TSource> first, IEnumerable<TSource> second, Func<TSource, TSource, bool> comparer) => first.Where(x => second.Count(y => comparer(x, y)) == 0);
        public static IEnumerable<TSource> Intersect<TSource>(this IEnumerable<TSource> first, IEnumerable<TSource> second, Func<TSource, TSource, bool> comparer) => first.Where(x => second.Count(y => comparer(x, y)) == 1);

        public static T[,] CreateRectangularArray<T>(this IList<T>[] arrays)
        {
            // TODO: Validation and special-casing for arrays.Count == 0
            int minorLength = arrays[0].Count();
            T[,] ret = new T[arrays.Length, minorLength];
            for (int i = 0; i < arrays.Length; i++)
            {
                var array = arrays[i];
                if (array.Count != minorLength)
                {
                    throw new ArgumentException
                        ("All arrays must be the same length");
                }
                for (int j = 0; j < minorLength; j++)
                {
                    ret[i, j] = array[j];
                }
            }
            return ret;
        }
        public static IList<T[]> CreateList<T>(this T[,] source)
        {
            return Enumerable.Range(source.GetLowerBound(0), source.GetUpperBound(0))
                .Select(row => Enumerable.Range(source.GetLowerBound(1), source.GetUpperBound(1))
                .Select(col => source[row, col]).ToArray())
                .ToList();
        }
    }
}
