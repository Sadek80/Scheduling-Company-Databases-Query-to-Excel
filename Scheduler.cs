using System;
using System.Collections.Generic;
using System.Threading;

namespace test_all_features_2
{
    class Scheduler
    {
        private static Scheduler _instance;
        private Scheduler() { }

        // Sengelton
        public static Scheduler Instance => _instance ?? (_instance = new Scheduler());

        public void ScheduleTask(TimeSpan timeToGo, Action task)
        {

            if (timeToGo <= TimeSpan.Zero)
            {
                timeToGo = TimeSpan.Zero;
            }
            var timer = new Timer(x =>
            {
                task.Invoke();
            }, null, timeToGo, new TimeSpan(0, 0, 0, 0, -1));
        }

    }
}
