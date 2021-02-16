using System;
using System.Collections.Generic;
using System.Text;

namespace test_all_features_2
{
    class HandlingScheduler
    {
        public static void IntervalInDays(int hour, int min, double interval, Action task)
        {
            interval = interval * 24;
            Scheduler.Instance.ScheduleTask(hour, min, interval, task);
        }

        public static void IntervalInMinutes(int hour, int min, double interval, Action task)
        {
            interval = interval / 60;
            Scheduler.Instance.ScheduleTask(hour, min, interval, task);
        }
    }
}
