using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PersonViewerSCA2
{
    public static class TaskManager
    {
        //https://codereview.stackexchange.com/questions/28760/custom-task-scheduler
        private static List<Scheduler> _schedulers;

        /// <summary>
        /// Gets a list of currently running schedulers.
        /// </summary>
        public static IEnumerable<Scheduler> RunningSchedulers
        {
            get { return _schedulers.Where(scheduler => scheduler.IsRunning()); }
        }

        /// <summary>
        /// Gets all scheduleres.
        /// </summary>
        public static IEnumerable<Scheduler> AllSchedulers(string name)
        {
            return _schedulers;
        }

        /// <summary>
        /// Gets a scheduler by it's name.
        /// </summary>
        public static Scheduler GetScheduler(string name)
        {
            return _schedulers.Find(scheduler => scheduler.Name == name);
        }

        /// <summary>
        /// Attempts to remove a scheduler from the list.
        /// </summary>
        public static bool RemoveScheduler(Scheduler scheduler)
        {
            return _schedulers.Remove(scheduler);
        }

        /// <summary>
        /// Creates a new scheduler and adds it the the list.
        /// </summary>
        public static Scheduler CreateScheduler(int priority, string name = null)
        {
            if (_schedulers == null)
                _schedulers = new List<Scheduler>();

            var scheduler = new Scheduler(priority, name);
            _schedulers.Add(scheduler);
            return scheduler;
        }
    }

    public class Scheduler : Queue<Job>
    {
        public string Name { get; private set; }
        public int Priority { get; set; }
        private Timer _triggerTimer;
        public int Interval { get; private set; }
        public EventHandler<Job> Trigger;

        public Scheduler(int priority, string name = null)
        {
            Priority = priority;
            Name = name;
            Interval = 20;
        }

        private void TriggerTimerCallBack(object state)
        {
            foreach (Job job in this.Where(job => DateTime.Now >= job.Execution && !job.Triggered))
            {
                job.Triggered = true;
                if (job.Repeating)
                    job.StartRepeating();
                Trigger(this, job);
            }
        }

        /// <summary>
        /// Sets the trigger timer interval.
        /// </summary>
        public void SetRefreshInterval(int interval)
        {
            Interval = interval;
        }

        /// <summary>
        /// Gets all the triggered jobs.
        /// </summary>
        public IEnumerable<Job> GetTriggeredJobs()
        {
            return this.Where(job => job.Triggered);
        }

        /// <summary>
        /// Gets all the repeat enabled jobs.
        /// </summary>
        public IEnumerable<Job> GetRepeatingJobs()
        {
            return this.Where(job => job.Repeating);
        }

        /// <summary>
        /// Returns whether the schedular is running.
        /// </summary>
        public bool IsRunning()
        {
            return _triggerTimer != null;
        }

        /// <summary>
        /// Stops the scheduler.
        /// </summary>
        public void Stop()
        {
            if (null == _triggerTimer) return;
            _triggerTimer.Dispose();
            _triggerTimer = null;
        }

        /// <summary>
        /// Starts the scheduler.
        /// </summary>
        public void Start()
        {
            if (_triggerTimer != null) return;
            _triggerTimer = new Timer(TriggerTimerCallBack, null, 0, Interval);
        }
    }

    public class Job : EventArgs, IDisposable
    {
        public string Name { get; set; }
        public Func<Task> Task { get; set; }
        public DateTime Execution { get; private set; }
        public bool Triggered { get; set; }
        public bool Repeating { get; set; }
        private int _interval;
        private Timer _triggerTimer;

        public Job(Func<Task> task, DateTime executionTime, string name = null)
        {
            Name = name;
            Task = task;
            Execution = executionTime;
        }

        public Job TriggerEvery(int interval)
        {
            _interval = interval;
            Repeating = true;
            return this;
        }

        public Job Seconds()
        {
            _interval = (_interval) * 1000;
            return this;
        }

        public Job Minutes()
        {
            _interval = _interval * (1000 * 60);
            return this;
        }

        public Job Hours()
        {
            _interval = _interval * (60 * 60 * 1000);
            return this;
        }

        public Job Days()
        {
            _interval = _interval * (60 * 60 * 24 * 1000);
            return this;
        }

        public void TriggerTimerCallBack(object sender)
        {
            if (Triggered)
                Triggered = false;
        }

        public void StartRepeating()
        {
            if (null != _triggerTimer) return;
            _triggerTimer = new Timer(TriggerTimerCallBack, null, _interval, _interval);
        }

        public void StopRepeating()
        {
            if (null == _triggerTimer) return;
            _triggerTimer.Dispose();
            _triggerTimer = null;
        }

        public void Dispose()
        {
            StopRepeating();
        }
    }
}
