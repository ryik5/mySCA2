using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;
using System.Diagnostics;

namespace ASTA
{
    /*
    public static class PeriodicTask
    {
        //keep a list of scheduled actions
        static Dictionary<Action, CancellationTokenSource> _scheduledActions = new Dictionary<Action, CancellationTokenSource>();

        //schedule an action
        public static async Task Run(Action action, TimeSpan period)
        {
            if (_scheduledActions.ContainsKey(action)) return; //this action is already scheduled, get out

            var cancellationTokenSrc = new CancellationTokenSource(); //create cancellation toket so we can abort later when we want

            _scheduledActions.Add(action, cancellationTokenSrc); //save it all into our disctionary

            //main scheduling loop
            Task task = null;
            while (!cancellationTokenSrc.IsCancellationRequested)
            {
                if (task == null || task.IsCompleted) //skip if previous invocation is still running (our requirements allow that)
                    task = Task.Run(action);

                await Task.Delay(period, cancellationTokenSrc.Token);
            }
        }

        public static void Stop(Action action)
        {
            if (!_scheduledActions.ContainsKey(action)) return; //this method has not been scheduled, get out
            _scheduledActions[action].Cancel(); //stop the task
            _scheduledActions.Remove(action); //remove from the collection
        }
    }


    public static class PeriodicTaskFactory
    {
        /// <summary>
        /// Starts the periodic task.
        /// </summary>
        /// <param name="action">The action.</param>
        /// <param name="intervalInMilliseconds">The interval in milliseconds.</param>
        /// <param name="delayInMilliseconds">The delay in milliseconds, i.e. how long it waits to kick off the timer.</param>
        /// <param name="duration">The duration.
        /// <example>If the duration is set to 10 seconds, the maximum time this task is allowed to run is 10 seconds.</example></param>
        /// <param name="maxIterations">The max iterations.</param>
        /// <param name="synchronous">if set to <c>true</c> executes each period in a blocking fashion and each periodic execution of the task
        /// is included in the total duration of the Task.</param>
        /// <param name="cancelToken">The cancel token.</param>
        /// <param name="periodicTaskCreationOptions"><see cref="TaskCreationOptions"/> used to create the task for executing the <see cref="Action"/>.</param>
        /// <returns>A <see cref="Task"/></returns>
        /// <remarks>
        /// Exceptions that occur in the <paramref name="action"/> need to be handled in the action itself. These exceptions will not be 
        /// bubbled up to the periodic task.
        /// </remarks>
        public static Task Start(Action action,
                                 int intervalInMilliseconds = Timeout.Infinite,
                                 int delayInMilliseconds = 0,
                                 int duration = Timeout.Infinite,
                                 int maxIterations = -1,
                                 bool synchronous = false,
                                 System.Threading.CancellationToken cancelToken = new CancellationToken(),
                                 TaskCreationOptions periodicTaskCreationOptions = TaskCreationOptions.None)
        {
            Stopwatch stopWatch = new Stopwatch();
            Action wrapperAction = () =>
            {
                CheckIfCancelled(cancelToken);
                action();
            };

            Action mainAction = () =>
            {
                MainPeriodicTaskAction(intervalInMilliseconds, delayInMilliseconds, duration, maxIterations, cancelToken, stopWatch, synchronous, wrapperAction, periodicTaskCreationOptions);
            };

            return Task.Factory.StartNew(mainAction, cancelToken, TaskCreationOptions.LongRunning, TaskScheduler.Current);
        }

        /// <summary>
        /// Mains the periodic task action.
        /// </summary>
        /// <param name="intervalInMilliseconds">The interval in milliseconds.</param>
        /// <param name="delayInMilliseconds">The delay in milliseconds.</param>
        /// <param name="duration">The duration.</param>
        /// <param name="maxIterations">The max iterations.</param>
        /// <param name="cancelToken">The cancel token.</param>
        /// <param name="stopWatch">The stop watch.</param>
        /// <param name="synchronous">if set to <c>true</c> executes each period in a blocking fashion and each periodic execution of the task
        /// is included in the total duration of the Task.</param>
        /// <param name="wrapperAction">The wrapper action.</param>
        /// <param name="periodicTaskCreationOptions"><see cref="TaskCreationOptions"/> used to create a sub task for executing the <see cref="Action"/>.</param>
        private static void MainPeriodicTaskAction(int intervalInMilliseconds,
                                                   int delayInMilliseconds,
                                                   int duration,
                                                   int maxIterations,
                                                   CancellationToken cancelToken,
                                                   Stopwatch stopWatch,
                                                   bool synchronous,
                                                   Action wrapperAction,
                                                   TaskCreationOptions periodicTaskCreationOptions)
        {
            TaskCreationOptions subTaskCreationOptions = TaskCreationOptions.AttachedToParent | periodicTaskCreationOptions;

            CheckIfCancelled(cancelToken);

            if (delayInMilliseconds > 0)
            {
                Thread.Sleep(delayInMilliseconds);
            }

            if (maxIterations == 0) { return; }

            int iteration = 0;

            ////////////////////////////////////////////////////////////////////////////
            // using a ManualResetEventSlim as it is more efficient in small intervals.
            // In the case where longer intervals are used, it will automatically use 
            // a standard WaitHandle....
            // see http://msdn.microsoft.com/en-us/library/vstudio/5hbefs30(v=vs.100).aspx
            using (ManualResetEventSlim periodResetEvent = new ManualResetEventSlim(false))
            {
                ////////////////////////////////////////////////////////////
                // Main periodic logic. Basically loop through this block
                // executing the action
                while (true)
                {
                    CheckIfCancelled(cancelToken);

                    Task subTask = Task.Factory.StartNew(wrapperAction, cancelToken, subTaskCreationOptions, TaskScheduler.Current);

                    if (synchronous)
                    {
                        stopWatch.Start();
                        try
                        {
                            subTask.Wait(cancelToken);
                        }
                        catch { }
                        stopWatch.Stop();
                    }

                    // use the same Timeout setting as the System.Threading.Timer, infinite timeout will execute only one iteration.
                    if (intervalInMilliseconds == Timeout.Infinite) { break; }

                    iteration++;

                    if (maxIterations > 0 && iteration >= maxIterations) { break; }

                    try
                    {
                        stopWatch.Start();
                        periodResetEvent.Wait(intervalInMilliseconds, cancelToken);
                        stopWatch.Stop();
                    }
                    finally
                    {
                        periodResetEvent.Reset();
                    }

                    CheckIfCancelled(cancelToken);

                    if (duration > 0 && stopWatch.ElapsedMilliseconds >= duration) { break; }
                }
            }
        }

        /// <summary>
        /// Checks if cancelled.
        /// </summary>
        /// <param name="cancelToken">The cancel token.</param>
        private static void CheckIfCancelled(CancellationToken cancellationToken)
        {
            if (cancellationToken == null)
                throw new ArgumentNullException("cancellationToken");

            cancellationToken.ThrowIfCancellationRequested();
        }
    }
    */

    /*
    public static class TaskManager
    {
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

    public class Scheduler :Queue<Job>
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

    public class Job :EventArgs, IDisposable
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
    */

}
