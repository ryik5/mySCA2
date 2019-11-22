using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ASTA.Classes
{
    
    /*
    public static class TaskExtensions
    {
        public static Task ForEachAsync<T>(this IEnumerable<T> source, int dop, Func<T, Task> body)
        {
            return Task.WhenAll(
                from partition in Partitioner.Create(source).GetPartitions(dop)
                select Task.Run(async delegate
                {
                    using (partition)
                        while (partition.MoveNext())
                            await body(partition.Current);
                }));
        }



        public static Task ForEachAsyncWithExceptios<T>(
            this IEnumerable<T> source, int dop, Func<T, Task> body)
        {
            ConcurrentQueue<Exception> exceptions = null;

            return Task.WhenAll(
                from partition in Partitioner.Create(source).GetPartitions(dop)
                select Task.Run(async delegate
                {
                    using (partition)
                    {
                        while (partition.MoveNext())
                        {
                            try
                            { await body(partition.Current); }
                            catch(Exception e)
                            {
                                LazyInitializer.
                                EnsureInitialized(ref exceptions).Enqueue(e);
                            }
                        }
                    }
                }));
        }
    }
    */
}
