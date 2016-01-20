using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTFHelper.Extensions
{
    static class EnumerableExtension
    {
        public static IEnumerable<IList<T>> Buffer<T>(this IEnumerable<T> lst, int num)
        {
            var buffer = new List<T>(num);
            foreach (T x in lst)
            {
                buffer.Add(x);
                if (buffer.Count == num)
                {
                    yield return buffer;
                    buffer = new List<T>(num);
                }
            }
            if (buffer.Any())
            {
                yield return buffer;
            }
        }
    }
}
