
using System;
using System.Collections.Generic;
using System.Threading;

namespace test_all_features_2
{
    /// <summary>
    /// Singelton Scheduler Class to handle the timer and repeat Tasks
    /// </summary>
    class Scheduler
    {
        private static Scheduler _instance;
        private static List<Timer> timers = new List<Timer>();

        public Dictionary<string, Timer> timerDictionary = new Dictionary<string, Timer>();



        private Scheduler() { }

        /// <summary>
        /// if the Instance is Null, then make the singelton Instance
        /// </summary>
        public static Scheduler Instance => _instance ?? (_instance = new Scheduler());

        /// <summary>
        /// Timer Method starts with the given hour and minute in the day
        /// then re-execute its self depending on the period of time and the duration until repeat
        /// </summary>
        /// <param name="hour">Starting Hour</param>
        /// <param name="min">Starting Minute</param>
        /// <param name="interval">Period of Time</param>
        /// <param name="repeat">Duration Until Repeat</param>
        /// <param name="timerName">Timer Name</param>
        /// <param name="task">Schedule Task Name</param>
        public void ScheduleTask(int hour, int min, string interval, double repeat, string timerName, Action task)
        {
            // Changing the Duration Depinding on the period of time
            if (interval == "Minutes")
            {
                repeat = repeat / 60;
            }
            else if (interval == "Days")
            {
                repeat = repeat * 24;

            }

            // Get the Current Time and the First Run Time
            DateTime now = DateTime.Now;
            DateTime firstRun = new DateTime(now.Year, now.Month, now.Day, hour, min, 0, 0);

            // If Time is passed, then add a day to the Current Time
            if (now > firstRun)
            {
                firstRun = firstRun.AddDays(1);
            }

            // Time to Start
            TimeSpan timeToGo = firstRun - now;
            if (timeToGo <= TimeSpan.Zero)
            {
                timeToGo = TimeSpan.Zero;
            }

            // Set the Timer to Start the Task due time and repeat in a given duration
            var timer = new Timer(x =>
            {
                task.Invoke();
            }, null, timeToGo, TimeSpan.FromHours(repeat));

            // Save the Timer Data
            timers.Add(timer);
            timerDictionary.Add(timerName, timer);


        }

        /// <summary>
        /// Stopping the Given Timer i.e. Task After Editing or Deleting Related Date to the Schedule Task
        /// </summary>
        /// <param name="timerName">Timer Name</param>
        public void killTimer(string timerName)
        {
            try
            {
                // Get Timer Object
                Timer getTimer = timerDictionary[timerName];

                // Stop the Timer
                timers.Find(timer => timer == getTimer).Change(Timeout.Infinite, Timeout.Infinite);

                // Remove the Timer Saved Data
                timerDictionary.Remove(timerName);
                timers.Remove(getTimer);

            }
            catch (Exception e)
            {
                Console.WriteLine("in the scheduler: " + e.Message);
            }

        }


    }
}
