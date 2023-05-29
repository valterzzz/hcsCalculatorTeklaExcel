using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace hcsCalculatorTeklaExcel
{
    public static class TimeUtils
    {
        private static readonly Stopwatch stopwatch = new Stopwatch();

        public static void StartTimer()
        {
            stopwatch.Restart();
        }

        public static long StopTimer()
        {
            stopwatch.Stop();
            return stopwatch.ElapsedMilliseconds;
        }
    }
}
