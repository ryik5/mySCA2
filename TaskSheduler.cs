using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Serialization.Formatters.Binary;
using System.Security;

namespace TaskSheduler
{
    //@Author Dennis Austin, 17 Dec 2007 
    //https://www.codeproject.com/Articles/2407/A-New-Task-Scheduler-Class-Library-for-NET#Sample%20Code
   
        /*
     Samples of using

        Listing the Scheduled Tasks from a Computer

This example connects to a computer named "DALLAS" and prints a summary of its scheduled tasks on the console. If DALLAS could not be accessed or the user account running the code does not have administrator privileges on DALLAS, then the constructor will throw an exception. Such exceptions aren't handled in the sample code.
Hide   Copy Code

// Get a ScheduledTasks object for the computer named "DALLAS"
ScheduledTasks st = new ScheduledTasks(@"\\DALLAS");

// Get an array of all the task names
string[] taskNames = st.GetTaskNames();

// Open each task, write a descriptive string to the console
foreach (string name in taskNames) {
    Task t = st.OpenTask(name);
    Console.WriteLine("  " + t.ToString());
    t.Close();
}

// Dispose the ScheduledTasks object to release COM resources.
st.Dispose();

Scheduling a New Task to be Run

Create a new task named "D checker" that runs chkdsk on the D: drive.
Hide   Shrink   Copy Code

//Get a ScheduledTasks object for the local computer.
ScheduledTasks st = new ScheduledTasks();

// Create a task
Task t;
try {
    t = st.CreateTask("D checker");
} catch (ArgumentException) {
    Console.WriteLine("Task name already exists");
    return;
}

// Fill in the program info
t.ApplicationName = "chkdsk.exe";
t.Parameters = "d: /f";
t.Comment = "Checks and fixes errors on D: drive";

// Set the account under which the task should run.
t.SetAccountInformation(@"THEDOMAIN\TheUser", "HisPasswd");

// Declare that the system must have been idle for ten minutes before 
// the task will start
t.IdleWaitMinutes = 10;

// Allow the task to run for no more than 2 hours, 30 minutes.
t.MaxRunTime = new TimeSpan(2, 30, 0);

// Set priority to only run when system is idle.
t.Priority = System.Diagnostics.ProcessPriorityClass.Idle;

// Create a trigger to start the task every Sunday at 6:30 AM.
t.Triggers.Add(new WeeklyTrigger(6, 30, DaysOfTheWeek.Sunday));

// Save the changes that have been made.
t.Save();
// Close the task to release its COM resources.
t.Close();
// Dispose the ScheduledTasks to release its COM resources.
st.Dispose();

Change the Time that a Task will be Run

This code opens a particular task and then updates any trigger with a start time, changing the time to 4:15 am. This makes use of the StartableTrigger abstract class because only those triggers have a start time.
Hide   Copy Code

// Get a ScheduledTasks object for the local computer.
ScheduledTasks st = new ScheduledTasks();

// Open a task we're interested in
Task task = st.OpenTask("D checker");

// Be sure the task was found before proceeding
if (task != null) {
    // Enumerate each trigger in the TriggerList of this task
    foreach (Trigger tr in task.Triggers) {
        // If this trigger has a start time, change it to 4:15 AM.
        if (tr is StartableTrigger) {
            (tr as StartableTrigger).StartHour = 4;
            (tr as StartableTrigger).StartMinute = 15;
        }
    }
    task.Save();
    task.Close();
}
st.Dispose();
         */



    public class TriggerList : IList, IDisposable
    {
        // Internal COM interface to access task that this list is associated with.
        private ITask iTask;
        // Trigger objects store in an ArrayList
        private ArrayList oTriggers;

        /// <summary>
        /// Internal constructor creates TriggerList using an ITask interface to initialize.
        /// </summary>
        /// <param name="iTask">Instance of an ITask.</param>
        internal TriggerList(ITask iTask)
        {
            this.iTask = iTask;
            ushort cnt = 0;
            iTask.GetTriggerCount(out cnt);
            oTriggers = new ArrayList(cnt + 5); //Allow for five additional entries without growing base array
            for (int i = 0; i < cnt; i++)
            {
                ITaskTrigger iTaskTrigger;
                iTask.GetTrigger((ushort)i, out iTaskTrigger);
                oTriggers.Add(Trigger.CreateTrigger(iTaskTrigger));
            }
        }

        /// <summary>
        /// Enumerator for TriggerList; implements IEnumerator interface.
        /// </summary>
        private class Enumerator : IEnumerator
        {
            private TriggerList outer;
            private int currentIndex;

            /// <summary>
            /// Internal constructor - Only accessible through <see cref="IEnumerable.GetEnumerator()"/>.
            /// </summary>
            /// <param name="outer">Instance of a TriggerList.</param>
            internal Enumerator(TriggerList outer)
            {
                this.outer = outer;
                Reset();
            }

            /// <summary>
            /// Moves to the next trigger. See <see cref="IEnumerator.MoveNext()"/> for more information.
            /// </summary>
            /// <returns>False if there is no next trigger.</returns>
            public bool MoveNext()
            {
                return ++currentIndex < outer.oTriggers.Count;
            }

            /// <summary>
            /// Reset trigger enumeration. See <see cref="IEnumerator.Reset()"/> for more information.
            /// </summary>
            public void Reset()
            {
                currentIndex = -1;
            }

            /// <summary>
            /// Retrieves the current trigger.  See <see cref="IEnumerator.Current"/> for more information.
            /// </summary>
            public object Current
            {
                get { return outer.oTriggers[currentIndex]; }
            }
        }

        #region Implementation of IList
        /// <summary>
        /// Removes the trigger at a specified index.
        /// </summary>
        /// <param name="index">Index of trigger to remove.</param>
        /// <exception cref="ArgumentOutOfRangeException">Index out of range.</exception>
        public void RemoveAt(int index)
        {
            if (index >= Count)
                throw new ArgumentOutOfRangeException("index", index, "Failed to remove Trigger. Index out of range.");
            ((Trigger)oTriggers[index]).Unbind(); //releases resources in the trigger
            oTriggers.RemoveAt(index); //Remove the Trigger object from the array representing the list
            iTask.DeleteTrigger((ushort)index); //Remove the trigger from the Task Scheduler
        }

        /// <summary>
        /// Not implemented; throws NotImplementedException.
        /// If implemented, would insert a trigger at a specified index. 
        /// </summary>
        /// <param name="index">Index to insert trigger.</param>
        /// <param name="value">Value of trigger to insert.</param>
        void IList.Insert(int index, object value)
        {
            throw new NotImplementedException("TriggerList does not support Insert().");
        }

        /// <summary>
        /// Removes the trigger from the collection.  If the trigger is not in
        /// the collection, nothing happens.  (No exception.)
        /// </summary>
        /// <param name="trigger">Trigger to remove.</param>
        public void Remove(Trigger trigger)
        {
            int i = IndexOf(trigger);
            if (i != -1)
                RemoveAt(i);
        }

        /// <summary>
        /// IList.Remove implementation.
        /// </summary>
        void IList.Remove(object value)
        {
            Remove(value as Trigger);
        }

        /// <summary>
        /// Test to see if trigger is part of the collection.
        /// </summary>
        /// <param name="trigger">Trigger to find.</param>
        /// <returns>true if trigger found in collection.</returns>
        public bool Contains(Trigger trigger)
        {
            return (IndexOf(trigger) != -1);
        }

        /// <summary>
        /// IList.Contains implementation.
        /// </summary>
        bool IList.Contains(object value)
        {
            return Contains(value as Trigger);
        }

        /// <summary>
        /// Remove all triggers from collection.
        /// </summary>
        public void Clear()
        {
            for (int i = Count - 1; i >= 0; i--)
            {
                RemoveAt(i);
            }
        }

        /// <summary>
        /// Returns the index of the supplied Trigger.
        /// </summary>
        /// <param name="trigger">Trigger to find.</param>
        /// <returns>Zero based index in collection, -1 if not a member.</returns>
        public int IndexOf(Trigger trigger)
        {
            for (int i = 0; i < Count; i++)
            {
                if (this[i].Equals(trigger))
                    return i;
            }
            return -1;
        }

        /// <summary>
        /// IList.IndexOf implementation.
        /// </summary>
        int IList.IndexOf(object value)
        {
            return IndexOf(value as Trigger);
        }

        /// <summary>
        /// Add the supplied Trigger to the collection.  The Trigger to be added must be unbound,
        /// i.e. it must not be a current member of a TriggerList--this or any other.
        /// </summary>
        /// <param name="trigger">Trigger to add.</param>
        /// <returns>Index of added trigger.</returns>
        /// <exception cref="ArgumentException">Trigger being added is already bound.</exception>
        public int Add(Trigger trigger)
        {
            // if trigger is already bound a list throw an exception
            if (trigger.Bound)
                throw new ArgumentException("A Trigger cannot be added if it is already in a list.");
            // Add a trigger to the task for this TaskList
            ITaskTrigger iTrigger;
            ushort index;
            iTask.CreateTrigger(out index, out iTrigger);
            // Add the Trigger to the TaskList
            trigger.Bind(iTrigger);
            int index2 = oTriggers.Add(trigger);
            // Verify index is the same in task and in list
            if (index2 != (int)index)
                throw new ApplicationException("Assertion Failure");
            return (int)index;
        }

        /// <summary>
        /// IList.Add implementation.
        /// </summary>
        int IList.Add(object value)
        {
            return Add(value as Trigger);
        }

        /// <summary>
        /// Gets read-only state of collection. Always false for TriggerLists.
        /// </summary>
        public bool IsReadOnly
        {
            get { return false; }
        }

        /// <summary>
        /// Access the Trigger at a specified index.  Assigning to a TriggerList element requires
        /// the value to unbound.  The previous list element becomes unbound and lost,
        /// while the newly assigned Trigger becomes bound in its place.
        /// </summary>
        /// <exception cref="ArgumentOutOfRangeException">Collection index out of range.</exception>
        public Trigger this[int index]
        {
            get
            {
                if (index >= Count)
                    throw new ArgumentOutOfRangeException("index", index, "TriggerList collection");
                return (Trigger)oTriggers[index];
            }
            set
            {
                if (index >= Count)
                    throw new ArgumentOutOfRangeException("index", index, "TriggerList collection");
                Trigger previous = (Trigger)oTriggers[index];
                value.Bind(previous);
                oTriggers[index] = value;
            }
        }

        /// <summary>
        /// IList.this[int] implementation.
        /// </summary>
        object IList.this[int index]
        {
            get { return this[index]; }
            set { this[index] = (value as Trigger); }
        }

        /// <summary>
        /// Returns whether collection is a fixed size. Always returns false for TriggerLists.
        /// </summary>
        public bool IsFixedSize
        {
            get { return false; }
        }
        #endregion

        #region Implementation of ICollection
        /// <summary>
        /// Gets the number of Triggers in the collection.
        /// </summary>
        public int Count
        {
            get
            {
                return oTriggers.Count;
            }
        }

        /// <summary>
        /// Copies all the Triggers in the collection to an array, beginning at the given index. 
        /// The Triggers assigned to the array are cloned from the originals, implying they are
        /// unbound copies.  (Can't tell if cloning is the intended semantics for this ICollection method,
        /// but it seems a good choice for TriggerLists.) 
        /// </summary>
        /// <param name="array">Array to copy triggers into.</param>
        /// <param name="index">Index at which to start copying.</param>
        public void CopyTo(System.Array array, int index)
        {
            if (oTriggers.Count > array.Length - index)
            {
                throw new ArgumentException("Array has insufficient space to copy the collection.");
            }
            for (int i = 0; i < oTriggers.Count; i++)
            {
                array.SetValue(((Trigger)oTriggers[i]).Clone(), index + i);
            }
        }

        /// <summary>
        /// Returns synchronizable state. Always false since the Task Scheduler is not
        /// thread safe.
        /// </summary>
        public bool IsSynchronized
        {
            get { return false; }
        }

        /// <summary>
        /// Gets the root object for synchronization. Always null since TriggerLists aren't synchronized.
        /// </summary>
        public object SyncRoot
        {
            get { return null; }
        }
        #endregion

        #region Implementation of IEnumerable
        /// <summary>
        /// Gets a TriggerList enumerator.
        /// </summary>
        /// <returns>Enumerator for TriggerList.</returns>
        public System.Collections.IEnumerator GetEnumerator()
        {
            return new Enumerator(this);
        }
        #endregion

        #region Implementation of IDisposable
        /// <summary>
        /// Unbinds and Disposes all the Triggers in the collection, releasing the com interfaces they hold.
        /// Destroys the internal private pointer to the ITask com interface, but does not 
        /// specifically release the interface because it is also in the containing task.
        /// </summary>
        public void Dispose()
        {
            foreach (object o in oTriggers)
            {
                ((Trigger)o).Unbind();
            }
            oTriggers = null;
            iTask = null;
        }
        #endregion
    }

    internal enum TriggerType
    {
        /// <summary>
        /// Trigger is set to run the task a single time. 
        /// </summary>
        RunOnce = 0,
        /// <summary>
        /// Trigger is set to run the task on a daily interval. 
        /// </summary>
        RunDaily = 1,
        /// <summary>
        /// Trigger is set to run the work item on specific days of a specific week of a specific month. 
        /// </summary>
        RunWeekly = 2,
        /// <summary>
        /// Trigger is set to run the task on a specific day(s) of the month.
        /// </summary>
        RunMonthly = 3,
        /// <summary>
        /// Trigger is set to run the task on specific days, weeks, and months.
        /// </summary>
        RunMonthlyDOW = 4,
        /// <summary>
        /// Trigger is set to run the task if the system remains idle for the amount of time specified by the idle wait time of the task.
        /// </summary>
        OnIdle = 5,
        /// <summary>
        /// Trigger is set to run the task at system startup.
        /// </summary>
        OnSystemStart = 6,
        /// <summary>
        /// Trigger is set to run the task when a user logs on. 
        /// </summary>
        OnLogon = 7
    }

    /// <summary>
    /// Values for days of the week (Monday, Tuesday, etc.)  These carry the Flags
    /// attribute so DaysOfTheWeek and be combined with | (or).
    /// </summary>
    [Flags]
    public enum DaysOfTheWeek : short
    {
        /// <summary>
        /// Sunday
        /// </summary>
        Sunday = 0x1,
        /// <summary>
        /// Monday
        /// </summary>
        Monday = 0x2,
        /// <summary>
        /// Tuesday
        /// </summary>
        Tuesday = 0x4,
        /// <summary>
        /// Wednesday
        /// </summary>
        Wednesday = 0x8,
        /// <summary>
        /// Thursday
        /// </summary>
        Thursday = 0x10,
        /// <summary>
        /// Friday
        /// </summary>
        Friday = 0x20,
        /// <summary>
        /// Saturday
        /// </summary>
        Saturday = 0x40
    }

    /// <summary>
    /// Values for week of month (first, second, ..., last)
    /// </summary>
    public enum WhichWeek : short
    {
        /// <summary>
        /// First week of the month
        /// </summary>
        FirstWeek = 1,
        /// <summary>
        /// Second week of the month
        /// </summary>
        SecondWeek = 2,
        /// <summary>
        /// Third week of the month
        /// </summary>
        ThirdWeek = 3,
        /// <summary>
        /// Fourth week of the month
        /// </summary>
        FourthWeek = 4,
        /// <summary>
        /// Last week of the month
        /// </summary>
        LastWeek = 5
    }

    /// <summary>
    /// Values for months of the year (January, February, etc.)  These carry the Flags
    /// attribute so DaysOfTheWeek and be combined with | (or).
    /// </summary>
    [Flags]
    public enum MonthsOfTheYear : short
    {
        /// <summary>
        /// January
        /// </summary>
        January = 0x1,
        /// <summary>
        /// February
        /// </summary>
        February = 0x2,
        /// <summary>
        /// March
        /// </summary>
        March = 0x4,
        /// <summary>
        /// April
        /// </summary>
        April = 0x8,
        /// <summary>
        ///May 
        /// </summary>
        May = 0x10,
        /// <summary>
        /// June
        /// </summary>
        June = 0x20,
        /// <summary>
        /// July
        /// </summary>
        July = 0x40,
        /// <summary>
        /// August
        /// </summary>
        August = 0x80,
        /// <summary>
        /// September
        /// </summary>
        September = 0x100,
        /// <summary>
        /// October
        /// </summary>
        October = 0x200,
        /// <summary>
        /// November
        /// </summary>
        November = 0x400,
        /// <summary>
        /// December
        /// </summary>
        December = 0x800
    }

    /// <summary>
    /// Trigger is a generalization of all the concrete trigger classes, and any actual
    /// Trigger object is one of those types.  When included in the TriggerList of a
    /// Task, a Trigger determines when a scheduled task will be run.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Create a concrete trigger for a specific start condition and then call TriggerList.Add
    /// to include it in a task's TriggerList.</para>
    /// <para>
    /// A Trigger that is not yet in a Task's TriggerList is said to be unbound and it holds
    /// no resources (i.e. COM interfaces).  Once it is added to a TriggerList, it is bound and
    /// holds a COM interface that is only released when the Trigger is removed from the list or
    /// the corresponding Task is closed.</para>  
    /// <para>
    /// A Trigger that is already bound cannot be added to a TriggerList.  To copy a Trigger from
    /// one list to another, use <see cref="Clone()"/> to create an unbound copy and then add the
    /// copy to the new list.  To move a Trigger from one list to another, use <see cref="TriggerList.Remove"/>
    /// to extract the Trigger from the first list before adding it to the second.</para>
    /// </remarks>
    public abstract class Trigger : ICloneable
    {
        #region Enums
        /// <summary>
        /// Flags for triggers
        /// </summary>
        [Flags]
        private enum TaskTriggerFlags
        {
            HasEndDate = 0x1,
            KillAtDurationEnd = 0x2,
            Disabled = 0x4
        }
        #endregion

        #region Fields
        private ITaskTrigger iTaskTrigger; //null for an unbound Trigger
        internal TaskTrigger taskTrigger;
        #endregion

        #region Constructors and Initializers
        /// <summary>
        /// Internal base constructor for an unbound Trigger.
        /// </summary>
        internal Trigger()
        {
            iTaskTrigger = null;
            taskTrigger = new TaskTrigger();
            taskTrigger.TriggerSize = (ushort)Marshal.SizeOf(taskTrigger);
            taskTrigger.BeginYear = (ushort)DateTime.Today.Year;
            taskTrigger.BeginMonth = (ushort)DateTime.Today.Month;
            taskTrigger.BeginDay = (ushort)DateTime.Today.Day;
        }

        /// <summary>
        /// Internal constructor which initializes itself from
        /// from an ITaskTrigger interface.
        /// </summary>
        /// <param name="iTrigger">Instance of ITaskTrigger from system task scheduler.</param>
        internal Trigger(ITaskTrigger iTrigger)
        {
            if (iTrigger == null)
                throw new ArgumentNullException("iTrigger", "ITaskTrigger instance cannot be null");
            taskTrigger = new TaskTrigger();
            taskTrigger.TriggerSize = (ushort)Marshal.SizeOf(taskTrigger);
            iTrigger.GetTrigger(ref taskTrigger);
            iTaskTrigger = iTrigger;
        }

        #endregion

        #region Implement ICloneable
        /// <summary>
        /// Clone returns an unbound copy of the Trigger object.  It can be use
        /// on either bound or unbound original.
        /// </summary>
        /// <returns></returns>
        public object Clone()
        {
            Trigger newTrigger = (Trigger)this.MemberwiseClone();
            newTrigger.iTaskTrigger = null; // The clone is not bound
            return newTrigger;
        }
        #endregion

        #region Properties

        /// <summary>
        /// Get whether the Trigger is currently bound
        /// </summary>
        internal bool Bound
        {
            get
            {
                return iTaskTrigger != null;
            }
        }

        /// <summary>
        /// Gets/sets the beginning year, month, and day for the trigger.
        /// </summary>
        public DateTime BeginDate
        {
            get
            {
                return new DateTime(taskTrigger.BeginYear, taskTrigger.BeginMonth, taskTrigger.BeginDay);
            }
            set
            {
                taskTrigger.BeginYear = (ushort)value.Year;
                taskTrigger.BeginMonth = (ushort)value.Month;
                taskTrigger.BeginDay = (ushort)value.Day;
                SyncTrigger();
            }
        }

        /// <summary>
        /// Gets/sets indication that the task uses an EndDate.  Returns true if a value has been
        /// set for the EndDate property.  Set can only be used to turn indication off.  
        /// </summary>
        /// <exception cref="ArgumentException">Has EndDate becomes true only by setting the EndDate
        /// property.</exception>
        public bool HasEndDate
        {
            get
            {
                return ((taskTrigger.Flags & (uint)TaskTriggerFlags.HasEndDate) == (uint)TaskTriggerFlags.HasEndDate);
            }
            set
            {
                if (value)
                    throw new ArgumentException("HasEndDate can only be set false");
                taskTrigger.Flags &= ~(uint)TaskTriggerFlags.HasEndDate;
                SyncTrigger();
            }
        }

        /// <summary>
        /// Gets/sets the ending year, month, and day for the trigger.  After a value has been set
        /// with EndDate, HasEndDate becomes true.
        /// </summary>
        public DateTime EndDate
        {
            get
            {
                if (taskTrigger.EndYear == 0)
                    return DateTime.MinValue;
                return new DateTime(taskTrigger.EndYear, taskTrigger.EndMonth, taskTrigger.EndDay);
            }
            set
            {
                taskTrigger.Flags |= (uint)TaskTriggerFlags.HasEndDate;
                taskTrigger.EndYear = (ushort)value.Year;
                taskTrigger.EndMonth = (ushort)value.Month;
                taskTrigger.EndDay = (ushort)value.Day;
                SyncTrigger();
            }
        }

        /// <summary>
        /// Gets/sets the number of minutes after the trigger fires that it remains active.  Used
        /// in conjunction with <see cref="IntervalMinutes"/> to run a task repeatedly for a period of time.
        /// For example, if you want to start a task at 8:00 A.M. repeatedly restart it until 5:00 P.M.,
        /// there would be 540 minutes (9 hours) in the duration.
        /// Can also be used to terminate a task that is running when the DurationMinutes expire.  Use
        /// <see cref="KillAtDurationEnd"/> to specify that task should be terminated at that time.
        /// </summary>
        /// <exception cref="ArgumentOutOfRangeException">Setting must be greater than or equal
        /// to the IntervalMinutes setting.</exception>
        public int DurationMinutes
        {
            get
            {
                return (int)taskTrigger.MinutesDuration;
            }
            set
            {
                if (value < taskTrigger.MinutesInterval)
                    throw new ArgumentOutOfRangeException("DurationMinutes", value, "DurationMinutes must be greater than or equal the IntervalMinutes value");
                taskTrigger.MinutesDuration = (uint)value;
                SyncTrigger();
            }
        }

        /// <summary>
        /// Gets/sets the number of minutes between executions for a task that is to be run repeatedly.
        /// Repetition continues until the interval specified in <see cref="DurationMinutes"/> expires.
        /// IntervalMinutes are counted from the start of the previous execution.
        /// </summary>
        /// <exception cref="ArgumentOutOfRangeException">Setting must be less than
        /// to the DurationMinutes setting.</exception>
        public int IntervalMinutes
        {
            get
            {
                return (int)taskTrigger.MinutesInterval;
            }
            set
            {
                if (value > taskTrigger.MinutesDuration)
                    throw new ArgumentOutOfRangeException("IntervalMinutes", value, "IntervalMinutes must be less than or equal the DurationMinutes value");
                taskTrigger.MinutesInterval = (uint)value;
                SyncTrigger();
            }
        }

        /// <summary>
        /// Gets/sets whether task will be killed (terminated) when DurationMinutes expires. 
        /// See <see cref="Trigger.DurationMinutes"/>.
        /// </summary>
        public bool KillAtDurationEnd
        {
            get
            {
                return ((taskTrigger.Flags & (uint)TaskTriggerFlags.KillAtDurationEnd) == (uint)TaskTriggerFlags.KillAtDurationEnd);
            }
            set
            {
                if (value)
                    taskTrigger.Flags |= (uint)TaskTriggerFlags.KillAtDurationEnd;
                else
                    taskTrigger.Flags &= ~(uint)TaskTriggerFlags.KillAtDurationEnd;
                SyncTrigger();
            }
        }

        /// <summary>
        /// Gets/sets whether trigger is disabled.
        /// </summary>
        public bool Disabled
        {
            get
            {
                return ((taskTrigger.Flags & (uint)TaskTriggerFlags.Disabled) == (uint)TaskTriggerFlags.Disabled);
            }
            set
            {
                if (value)
                    taskTrigger.Flags |= (uint)TaskTriggerFlags.Disabled;
                else
                    taskTrigger.Flags &= ~(uint)TaskTriggerFlags.Disabled;
                SyncTrigger();
            }
        }
        #endregion

        #region Methods
        /// <summary>
        /// Creates a new, bound Trigger object from an ITaskTrigger interface.  The type of the
        /// concrete object created is determined by the type of ITaskTrigger.
        /// </summary>
        /// <param name="iTaskTrigger">Instance of ITaskTrigger.</param>
        /// <returns>One of the concrete classes derived from Trigger.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException">Unable to recognize trigger type.</exception>
        internal static Trigger CreateTrigger(ITaskTrigger iTaskTrigger)
        {
            if (iTaskTrigger == null)
                throw new ArgumentNullException("iTaskTrigger", "Instance of ITaskTrigger cannot be null");
            TaskTrigger sTaskTrigger = new TaskTrigger();
            sTaskTrigger.TriggerSize = (ushort)Marshal.SizeOf(sTaskTrigger);
            iTaskTrigger.GetTrigger(ref sTaskTrigger);
            switch ((TriggerType)sTaskTrigger.Type)
            {
                case TriggerType.RunOnce:
                    return new RunOnceTrigger(iTaskTrigger);
                case TriggerType.RunDaily:
                    return new DailyTrigger(iTaskTrigger);
                case TriggerType.RunWeekly:
                    return new WeeklyTrigger(iTaskTrigger);
                case TriggerType.RunMonthlyDOW:
                    return new MonthlyDOWTrigger(iTaskTrigger);
                case TriggerType.RunMonthly:
                    return new MonthlyTrigger(iTaskTrigger);
                case TriggerType.OnIdle:
                    return new OnIdleTrigger(iTaskTrigger);
                case TriggerType.OnSystemStart:
                    return new OnSystemStartTrigger(iTaskTrigger);
                case TriggerType.OnLogon:
                    return new OnLogonTrigger(iTaskTrigger);
                default:
                    throw new ArgumentException("Unable to recognize type of trigger referenced in iTaskTrigger",
                                                "iTaskTrigger");
            }
        }

        /// <summary>
        /// When a bound Trigger is changed, the corresponding trigger in the system
        /// Task Scheduler is updated to stay in sync with the local structure.
        /// </summary>
        protected void SyncTrigger()
        {
            if (iTaskTrigger != null) iTaskTrigger.SetTrigger(ref taskTrigger);
        }

        /// <summary>
        /// Bind a Trigger object to an ITaskTrigger interface.  This causes the Trigger to
        /// sync itself with the interface and remain in sync whenever it is modified in the future.
        /// If the Trigger is already bound, an ArgumentException is thrown.
        /// </summary>
        /// <param name="iTaskTrigger">An interface representing a trigger in Task Scheduler.</param>
        /// <exception cref="ArgumentException">Attempt to bind and already bound trigger.</exception>
        internal void Bind(ITaskTrigger iTaskTrigger)
        {
            if (this.iTaskTrigger != null)
                throw new ArgumentException("Attempt to bind an already bound trigger");
            this.iTaskTrigger = iTaskTrigger;
            iTaskTrigger.SetTrigger(ref taskTrigger);
        }
        /// <summary>
        /// Bind a Trigger to the same interface the argument trigger is bound to.  
        /// </summary>
        /// <param name="trigger">A bound Trigger. </param>
        internal void Bind(Trigger trigger)
        {
            Bind(trigger.iTaskTrigger);
        }

        /// <summary>
        /// Break the connection between this Trigger and the system Task Scheduler.  This
        /// releases COM resources used in bound Triggers.
        /// </summary>
        internal void Unbind()
        {
            if (iTaskTrigger != null)
            {
                Marshal.ReleaseComObject(iTaskTrigger);
                iTaskTrigger = null;
            }
        }

        /// <summary>
        /// Gets a string, supplied by the WindowsTask Scheduler, of a bound Trigger. 
        /// For an unbound trigger, returns "Unbound Trigger".
        /// </summary>
        /// <returns>String representation of the trigger.</returns>
        public override string ToString()
        {
            if (iTaskTrigger != null)
            {
                IntPtr lpwstr;
                iTaskTrigger.GetTriggerString(out lpwstr);
                return CoTaskMem.LPWStrToString(lpwstr);
            }
            else
            {
                return "Unbound " + this.GetType().ToString();
            }
        }

        /// <summary>
        /// Determines if two triggers are internally equal.  Does not consider whether
        /// the Triggers are bound or not.
        /// </summary>
        /// <param name="obj">Value of trigger to compare.</param>
        /// <returns>true if triggers are equivalent.</returns>
        public override bool Equals(object obj)
        {
            return taskTrigger.Equals(((Trigger)obj).taskTrigger);
        }

        /// <summary>
        /// Gets a hash code for the current trigger.  A Trigger has the same hash
        /// code whether it is bound or not.
        /// </summary>
        /// <returns>Hash code value.</returns>
        public override int GetHashCode()
        {
            return taskTrigger.GetHashCode();
        }
        #endregion

    }

    /// <summary>
    /// Generalization of all triggers that have a start time.
    /// </summary>
    /// <remarks>StartableTrigger serves as a base class for triggers with a
    /// start time, but it has little use to clients.</remarks>
    public abstract class StartableTrigger : Trigger
    {
        /// <summary>
        /// Internal constructor, same as base.
        /// </summary>
        internal StartableTrigger() : base()
        {
        }

        /// <summary>
        /// Internal constructor from ITaskTrigger interface.
        /// </summary>
        /// <param name="iTrigger">ITaskTrigger from system Task Scheduler.</param>
        internal StartableTrigger(ITaskTrigger iTrigger) : base(iTrigger)
        {
        }

        /// <summary>
        /// Sets the start time of the trigger.
        /// </summary>
        /// <param name="hour">Hour of the day that the trigger will fire.</param>
        /// <param name="minute">Minute of the hour.</param>
        /// <exception cref="ArgumentOutOfRangeException">The hour is not between 0 and 23 or the minute is not between 0 and 59.</exception>
        protected void SetStartTime(ushort hour, ushort minute)
        {
            //			if (hour < 0 || hour > 23)
            //				throw new ArgumentOutOfRangeException("hour", hour, "hour must be between 0 and 23");
            //			if (minute < 0 || minute > 59)
            //				throw new ArgumentOutOfRangeException("minute", minute, "minute must be between 0 and 59");
            //			taskTrigger.StartHour = hour;
            //			taskTrigger.StartMinute = minute;
            //			base.SyncTrigger();
            StartHour = (short)hour;
            StartMinute = (short)minute;
        }

        /// <summary>
        /// Gets/sets hour of the day that trigger will fire (24 hour clock).
        /// </summary>
        public short StartHour
        {
            get
            {
                return (short)taskTrigger.StartHour;
            }
            set
            {
                if (value < 0 || value > 23)
                    throw new ArgumentOutOfRangeException("hour", value, "hour must be between 0 and 23");
                taskTrigger.StartHour = (ushort)value;
                base.SyncTrigger();
            }
        }

        /// <summary>
        /// Gets/sets minute of the hour (specified in <see cref="StartHour"/>) that trigger will fire.
        /// </summary>
        public short StartMinute
        {
            get
            {
                return (short)taskTrigger.StartMinute;
            }
            set
            {
                if (value < 0 || value > 59)
                    throw new ArgumentOutOfRangeException("minute", value, "minute must be between 0 and 59");
                taskTrigger.StartMinute = (ushort)value;
                base.SyncTrigger();
            }
        }
    }

    /// <summary>
    /// Trigger that fires once only.
    /// </summary>
    public class RunOnceTrigger : StartableTrigger
    {
        /// <summary>
        /// Create a RunOnceTrigger that fires when specified.
        /// </summary>
        /// <param name="runDateTime">Date and time to fire.</param>
        public RunOnceTrigger(DateTime runDateTime) : base()
        {
            taskTrigger.BeginYear = (ushort)runDateTime.Year;
            taskTrigger.BeginMonth = (ushort)runDateTime.Month;
            taskTrigger.BeginDay = (ushort)runDateTime.Day;
            SetStartTime((ushort)runDateTime.Hour, (ushort)runDateTime.Minute);
            taskTrigger.Type = TaskTriggerType.TIME_TRIGGER_ONCE;
        }

        /// <summary>
        /// Internal constructor to create from existing ITaskTrigger interface.
        /// </summary>
        /// <param name="iTrigger">ITaskTrigger from system Task Scheduler.</param>
        internal RunOnceTrigger(ITaskTrigger iTrigger) : base(iTrigger)
        {
        }
    }

    /// <summary>
    /// Trigger that fires at a specified time, every so many days.
    /// </summary>
    public class DailyTrigger : StartableTrigger
    {
        /// <summary>
        /// Creates a DailyTrigger that fires only at an interval of so many days.
        /// </summary>
        /// <param name="hour">Hour of day trigger will fire.</param>
        /// <param name="minutes">Minutes of the hour trigger will fire.</param>
        /// <param name="daysInterval">Number of days between task runs.</param>
        public DailyTrigger(short hour, short minutes, short daysInterval) : base()
        {
            SetStartTime((ushort)hour, (ushort)minutes);
            taskTrigger.Type = TaskTriggerType.TIME_TRIGGER_DAILY;
            taskTrigger.Data.daily.DaysInterval = (ushort)daysInterval;
        }

        /// <summary>
        /// Creates DailyTrigger that fires every day.
        /// </summary>
        /// <param name="hour">Hour of day trigger will fire.</param>
        /// <param name="minutes">Minutes of hour (specified in "hour") trigger will fire.</param>
        public DailyTrigger(short hour, short minutes) : this(hour, minutes, 1)
        {
        }

        /// <summary>
        /// Internal constructor to create from existing ITaskTrigger interface.
        /// </summary>
        /// <param name="iTrigger">ITaskTrigger from system Task Scheduler.</param>
        internal DailyTrigger(ITaskTrigger iTrigger) : base(iTrigger)
        {
        }

        /// <summary>
        /// Gets/sets the number of days between successive firings.
        /// </summary>
        public short DaysInterval
        {
            get
            {
                return (short)taskTrigger.Data.daily.DaysInterval;
            }
            set
            {
                taskTrigger.Data.daily.DaysInterval = (ushort)value;
                base.SyncTrigger();
            }
        }
    }

    /// <summary>
    /// Trigger that fires at a specified time, on specified days of the week,
    /// every so many weeks.
    /// </summary>
    public class WeeklyTrigger : StartableTrigger
    {
        /// <summary>
        /// Creates a WeeklyTrigger that is eligible to fire only during certain weeks.
        /// </summary>
        /// <param name="hour">Hour of day trigger will fire.</param>
        /// <param name="minutes">Minutes of hour (specified in "hour") trigger will fire.</param>
        /// <param name="daysOfTheWeek">Days of the week task will run.</param>
        /// <param name="weeksInterval">Number of weeks between task runs.</param>
        public WeeklyTrigger(short hour, short minutes, DaysOfTheWeek daysOfTheWeek, short weeksInterval) : base()
        {
            SetStartTime((ushort)hour, (ushort)minutes);
            taskTrigger.Type = TaskTriggerType.TIME_TRIGGER_WEEKLY;
            taskTrigger.Data.weekly.WeeksInterval = (ushort)weeksInterval;
            taskTrigger.Data.weekly.DaysOfTheWeek = (ushort)daysOfTheWeek;
        }

        /// <summary>
        /// Creates a WeeklyTrigger that is eligible to fire during any week.
        /// </summary>
        /// <param name="hour">Hour of day trigger will fire.</param>
        /// <param name="minutes">Minutes of hour (specified in "hour") trigger will fire.</param>
        /// <param name="daysOfTheWeek">Days of the week task will run.</param>
        public WeeklyTrigger(short hour, short minutes, DaysOfTheWeek daysOfTheWeek) : this(hour, minutes, daysOfTheWeek, 1)
        {
        }

        /// <summary>
        /// Internal constructor to create from existing ITaskTrigger interface.
        /// </summary>
        /// <param name="iTrigger">ITaskTrigger interface from system Task Scheduler.</param>
        internal WeeklyTrigger(ITaskTrigger iTrigger) : base(iTrigger)
        {
        }

        /// <summary>
        /// Gets/sets number of weeks from one eligible week to the next.
        /// </summary>
        public short WeeksInterval
        {
            get
            {
                return (short)taskTrigger.Data.weekly.WeeksInterval;
            }
            set
            {
                taskTrigger.Data.weekly.WeeksInterval = (ushort)value;
                base.SyncTrigger();
            }
        }

        /// <summary>
        /// Gets/sets the days of the week on which the trigger fires.
        /// </summary>
        public DaysOfTheWeek WeekDays
        {
            get
            {
                return (DaysOfTheWeek)taskTrigger.Data.weekly.DaysOfTheWeek;
            }
            set
            {
                taskTrigger.Data.weekly.DaysOfTheWeek = (ushort)value;
                base.SyncTrigger();
            }
        }
    }

    /// <summary>
    /// Trigger that fires at a specified time, on specified days of the week, 
    /// in specified weeks of the month, during specified months of the year.
    /// </summary>
    public class MonthlyDOWTrigger : StartableTrigger
    {
        /// <summary>
        /// Creates a MonthlyDOWTrigger that fires during specified months only.
        /// </summary>
        /// <param name="hour">Hour of day trigger will fire.</param>
        /// <param name="minutes">Minute of the hour trigger will fire.</param>
        /// <param name="daysOfTheWeek">Days of the week trigger will fire.</param>
        /// <param name="whichWeeks">Weeks of the month trigger will fire.</param>
        /// <param name="months">Months of the year trigger will fire.</param>
        public MonthlyDOWTrigger(short hour, short minutes, DaysOfTheWeek daysOfTheWeek, WhichWeek whichWeeks, MonthsOfTheYear months) : base()
        {
            SetStartTime((ushort)hour, (ushort)minutes);
            taskTrigger.Type = TaskTriggerType.TIME_TRIGGER_MONTHLYDOW;
            taskTrigger.Data.monthlyDOW.WhichWeek = (ushort)whichWeeks;
            taskTrigger.Data.monthlyDOW.DaysOfTheWeek = (ushort)daysOfTheWeek;
            taskTrigger.Data.monthlyDOW.Months = (ushort)months;
        }

        /// <summary>
        /// Creates a MonthlyDOWTrigger that fires every month.
        /// </summary>
        /// <param name="hour">Hour of day trigger will fire.</param>
        /// <param name="minutes">Minute of the hour trigger will fire.</param>
        /// <param name="daysOfTheWeek">Days of the week trigger will fire.</param>
        /// <param name="whichWeeks">Weeks of the month trigger will fire.</param>
        public MonthlyDOWTrigger(short hour, short minutes, DaysOfTheWeek daysOfTheWeek, WhichWeek whichWeeks) :
            this(hour, minutes, daysOfTheWeek, whichWeeks,
            MonthsOfTheYear.January | MonthsOfTheYear.February | MonthsOfTheYear.March | MonthsOfTheYear.April | MonthsOfTheYear.May | MonthsOfTheYear.June | MonthsOfTheYear.July | MonthsOfTheYear.August | MonthsOfTheYear.September | MonthsOfTheYear.October | MonthsOfTheYear.November | MonthsOfTheYear.December)
        {
        }

        /// <summary>
        /// Internal constructor to create from existing ITaskTrigger interface.
        /// </summary>
        /// <param name="iTrigger">ITaskTrigger from the system Task Scheduler.</param>
        internal MonthlyDOWTrigger(ITaskTrigger iTrigger) : base(iTrigger)
        {
        }

        /// <summary>
        /// Gets/sets weeks of the month in which trigger will fire.
        /// </summary>
        public short WhichWeeks
        {
            get
            {
                return (short)taskTrigger.Data.monthlyDOW.WhichWeek;
            }
            set
            {
                taskTrigger.Data.monthlyDOW.WhichWeek = (ushort)value;
                base.SyncTrigger();
            }
        }

        /// <summary>
        /// Gets/sets days of the week on which trigger will fire.
        /// </summary>
        public DaysOfTheWeek WeekDays
        {
            get
            {
                return (DaysOfTheWeek)taskTrigger.Data.monthlyDOW.DaysOfTheWeek;
            }
            set
            {
                taskTrigger.Data.monthlyDOW.DaysOfTheWeek = (ushort)value;
                base.SyncTrigger();
            }
        }

        /// <summary>
        /// Gets/sets months of the year in which trigger will fire.
        /// </summary>
        public MonthsOfTheYear Months
        {
            get
            {
                return (MonthsOfTheYear)taskTrigger.Data.monthlyDOW.Months;
            }
            set
            {
                taskTrigger.Data.monthlyDOW.Months = (ushort)value;
                base.SyncTrigger();
            }
        }
    }

    /// <summary>
    /// Trigger that fires at a specified time, on specified days of themonth,
    /// on specified months of the year.
    /// </summary>
    public class MonthlyTrigger : StartableTrigger
    {
        /// <summary>
        /// Creates a MonthlyTrigger that fires only during specified months of the year.
        /// </summary>
        /// <param name="hour">Hour of day trigger will fire.</param>
        /// <param name="minutes">Minutes of hour (specified in "hour") trigger will fire.</param>
        /// <param name="daysOfMonth">Days of the month trigger will fire.  (See <see cref="Days"/> property.</param>
        /// <param name="months">Months of the year trigger will fire.</param>
        public MonthlyTrigger(short hour, short minutes, int[] daysOfMonth, MonthsOfTheYear months) : base()
        {
            SetStartTime((ushort)hour, (ushort)minutes);
            taskTrigger.Type = TaskTriggerType.TIME_TRIGGER_MONTHLYDATE;
            taskTrigger.Data.monthlyDate.Months = (ushort)months;
            taskTrigger.Data.monthlyDate.Days = (uint)IndicesToMask(daysOfMonth);
        }

        /// <summary>
        /// Creates a MonthlyTrigger that fires during any month.
        /// </summary>
        /// <param name="hour">Hour of day trigger will fire.</param>
        /// <param name="minutes">Minutes of hour (specified in "hour") trigger will fire.</param>
        /// <param name="daysOfMonth">Days of the month trigger will fire.  (See <see cref="Days"/> property.</param>
        public MonthlyTrigger(short hour, short minutes, int[] daysOfMonth) :
            this(hour, minutes, daysOfMonth,
            MonthsOfTheYear.January | MonthsOfTheYear.February | MonthsOfTheYear.March | MonthsOfTheYear.April | MonthsOfTheYear.May | MonthsOfTheYear.June | MonthsOfTheYear.July | MonthsOfTheYear.August | MonthsOfTheYear.September | MonthsOfTheYear.October | MonthsOfTheYear.November | MonthsOfTheYear.December)
        {
        }

        /// <summary>
        /// Internal constructor to create from existing ITaskTrigger interface.
        /// </summary>
        /// <param name="iTrigger">ITaskTrigger from system Task Scheduler.</param>
        internal MonthlyTrigger(ITaskTrigger iTrigger) : base(iTrigger)
        {
        }


        /// <summary>
        /// Gets/sets months of the year trigger will fire.
        /// </summary>
        public MonthsOfTheYear Months
        {
            get
            {
                return (MonthsOfTheYear)taskTrigger.Data.monthlyDate.Months;
            }
            set
            {
                taskTrigger.Data.monthlyDOW.Months = (ushort)value;
                base.SyncTrigger();
            }
        }

        /// <summary>
        /// Convert an integer representing a mask to an array where each element contains the index
        /// of a bit that is ON in the mask.  Bits are considered to number from 1 to 32.
        /// </summary>
        /// <param name="mask">An interger to be interpreted as a mask.</param>
        /// <returns>An array with an element for each bit of the mask which is ON.</returns>
        private static int[] MaskToIndices(int mask)
        {
            //count bits in mask
            int cnt = 0;
            for (int i = 0; (mask >> i) > 0; i++)
                cnt = cnt + (1 & (mask >> i));
            //allocate return array with one entry for each bit
            int[] indices = new int[cnt];
            //fill array with bit indices
            cnt = 0;
            for (int i = 0; (mask >> i) > 0; i++)
                if ((1 & (mask >> i)) == 1)
                    indices[cnt++] = i + 1;
            return indices;
        }
        /// <summary>
        /// Converts an array of bit indices into a mask with bits  turned ON at every index
        /// contained in the array.  Indices must be from 1 to 32 and bits are numbered the same.
        /// </summary>
        /// <param name="indices">An array with an element for each bit of the mask which is ON.</param>
        /// <returns>An interger to be interpreted as a mask.</returns>
        private static int IndicesToMask(int[] indices)
        {
            int mask = 0;
            foreach (int index in indices)
            {
                if (index < 1 || index > 31) throw new ArgumentException("Days must be in the range 1..31");
                mask = mask | 1 << (index - 1);
            }
            return mask;
        }

        /// <summary>
        /// Gets/sets days of the month trigger will fire.
        /// </summary>
        /// <value>An array with one element for each day that the trigger will fire.
        /// The value of the element is the number of the day, in the range 1..31.</value>
        public int[] Days
        {
            get
            {
                return MaskToIndices((int)taskTrigger.Data.monthlyDate.Days);
            }
            set
            {
                taskTrigger.Data.monthlyDate.Days = (uint)IndicesToMask(value);
                base.SyncTrigger();
            }
        }
    }

    /// <summary>
    /// Trigger that fires when the system is idle for a period of time.
    /// Length of period set by <see cref="Task.IdleWaitMinutes"/>.
    /// </summary>
    public class OnIdleTrigger : Trigger
    {
        /// <summary>
        /// Creates an OnIdleTrigger.  Idle period set separately.
        /// See <see cref="Task.IdleWaitMinutes"/> inherited property.
        /// </summary>
        public OnIdleTrigger() : base()
        {
            taskTrigger.Type = TaskTriggerType.EVENT_TRIGGER_ON_IDLE;
        }

        /// <summary>
        /// Internal constructor to create from existing ITaskTrigger interface.
        /// </summary>
        /// <param name="iTrigger">Current base Trigger.</param>
        internal OnIdleTrigger(ITaskTrigger iTrigger) : base(iTrigger)
        {
        }
    }

    /// <summary>
    /// Trigger that fires when the system starts.
    /// </summary>
    public class OnSystemStartTrigger : Trigger
    {
        /// <summary>
        /// Creates an OnSystemStartTrigger.
        /// </summary>
        public OnSystemStartTrigger() : base()
        {
            taskTrigger.Type = TaskTriggerType.EVENT_TRIGGER_AT_SYSTEMSTART;
        }

        /// <summary>
        /// Internal constructor to create from existing ITaskTrigger interface.
        /// </summary>
        /// <param name="iTrigger">ITaskTrigger interface from system Task Scheduler.</param>
        internal OnSystemStartTrigger(ITaskTrigger iTrigger) : base(iTrigger)
        {
        }
    }

    /// <summary>
    /// Trigger that fires when a user logs on.
    /// </summary>
    /// <remarks>Triggers of this type fire when any user logs on, not just the
    /// user identified in the account information.</remarks>
    public class OnLogonTrigger : Trigger
    {
        /// <summary>
        /// Creates an OnLogonTrigger.
        /// </summary>
        public OnLogonTrigger() : base()
        {
            taskTrigger.Type = TaskTriggerType.EVENT_TRIGGER_AT_LOGON;
        }
        /// <summary>
        /// Internal constructor to create from existing ITaskTrigger interface.
        /// </summary>
        /// <param name="iTrigger">ITaskTrigger from system Task Scheduler.</param>
        internal OnLogonTrigger(ITaskTrigger iTrigger) : base(iTrigger)
        {
        }
    }
    
    public class TaskList : IEnumerable, IDisposable
    {
        /// <summary>
        /// Scheduled Tasks folder supporting this TaskList.
        /// </summary>
        private ScheduledTasks st = null;

        /// <summary>
        /// Name of the target computer whose Scheduled Tasks are to be accessed.
        /// </summary>
        private string nameComputer;

        /// <summary>
        /// Constructors - marked internal so you have to create using Scheduler class.
        /// </summary>
        internal TaskList()
        {
            st = new ScheduledTasks();
        }

        internal TaskList(string computer)
        {
            st = new ScheduledTasks(computer);
        }

        /// <summary>
        /// Enumerator for <c>TaskList</c>
        /// </summary>
        private class Enumerator : IEnumerator
        {
            private ScheduledTasks outer;
            private string[] nameTask;
            private int curIndex;
            private Task curTask;

            /// <summary>
            /// Internal constructor - Only accessable through <see cref="IEnumerable.GetEnumerator()"/>
            /// </summary>
            /// <param name="st">ScheduledTasks object</param>
            internal Enumerator(ScheduledTasks st)
            {
                outer = st;
                nameTask = st.GetTaskNames();
                Reset();
            }

            /// <summary>
            /// Moves to the next task. See <see cref="IEnumerator.MoveNext()"/> for more information.
            /// </summary>
            /// <returns>true if next task found, false if no more tasks.</returns>
            public bool MoveNext()
            {
                bool ok = ++curIndex < nameTask.Length;
                if (ok) curTask = outer.OpenTask(nameTask[curIndex]);
                return ok;
            }

            /// <summary>
            /// Reset task enumeration. See <see cref="IEnumerator.Reset()"/> for more information.
            /// </summary>
            public void Reset()
            {
                curIndex = -1;
                curTask = null;
            }

            /// <summary>
            /// Retrieves the current task.  See <see cref="IEnumerator.Current"/> for more information.
            /// </summary>
            public object Current
            {
                get
                {
                    return curTask;
                }
            }
        }

        /// <summary>
        /// Name of target computer
        /// </summary>
        internal string TargetComputer
        {
            get
            {
                return nameComputer;
            }
            set
            {
                st.Dispose();
                st = new ScheduledTasks(value);
                nameComputer = value;
            }
        }

        /// <summary>
        /// Creates a new task on the system with the supplied <paramref name="name" />.
        /// </summary>
        /// <param name="name">Unique display name for the task. If not unique, an ArgumentException will be thrown.</param>
        /// <returns>Instance of new task</returns>
        /// <exception cref="ArgumentException">There is already a task of the same name as the one supplied for the new task.</exception>
        public Task NewTask(string name)
        {
            return st.CreateTask(name);
        }

        /// <summary>
        /// Deletes the task of the given <paramref name="name" />.
        /// </summary>
        /// <param name="name">Name of task to delete</param>
        public void Delete(string name)
        {
            st.DeleteTask(name);
        }

        /// <summary>
        /// Indexer which retrieves task of given <paramref name="name" />.
        /// </summary>
        /// <param name="name">Name of task to retrieve</param>
        public Task this[string name]
        {
            get
            {
                return st.OpenTask(name);
            }
        }

        #region Implementation of IEnumerable
        /// <summary>
        /// Gets a TaskList enumerator
        /// </summary>
        /// <returns>Enumerator for TaskList</returns>
        public System.Collections.IEnumerator GetEnumerator()
        {
            return new Enumerator(st);
        }
        #endregion

        #region Implementation of IDisposable
        /// <summary>
        /// Disposes TaskList
        /// </summary>
        public void Dispose()
        {
            st.Dispose();
        }
        #endregion

    }
    
    public class ScheduledTasks : IDisposable
    {
        /// <summary>
        /// Underlying COM interface.
        /// </summary>
        private ITaskScheduler its = null;

        // --- Contructors ---

        /// <summary>
        /// Constructor to use Scheduled tasks of a remote computer identified by a UNC
        /// name.  The calling process must have administrative privileges on the remote machine. 
        /// May throw exception if the computer's task scheduler cannot be reached, and may
        /// give strange results if the argument is not in UNC format.
        /// </summary>
        /// <param name="computer">The remote computer's UNC name, e.g. "\\DALLAS".</param>
        /// <exception cref="ArgumentException">The Task Scheduler could not be accessed.</exception>
        public ScheduledTasks(string computer) : this()
        {
            its.SetTargetComputer(computer);
        }

        /// <summary>
        /// Constructor to use Scheduled Tasks of the local computer.
        /// </summary>
        public ScheduledTasks()
        {
            CTaskScheduler cts = new CTaskScheduler();
            its = (ITaskScheduler)cts;
        }

        // --- Methods ---

        private string[] GrowStringArray(string[] s, uint n)
        {
            string[] sl = new string[s.Length + n];
            for (int i = 0; i < s.Length; i++) { sl[i] = s[i]; }
            return sl;
        }

        /// <summary>
        /// Return the names of all scheduled tasks.  The names returned include the file extension ".job";
        /// methods requiring a task name can take the name with or without the extension.
        /// </summary>
        /// <returns>The names in a string array.</returns>
        public string[] GetTaskNames()
        {
            const int TASKS_TO_FETCH = 10;
            string[] taskNames = { };
            int nTaskNames = 0;

            IEnumWorkItems ienum;
            its.Enum(out ienum);

            uint nFetchedTasks;
            IntPtr pNames;

            while (ienum.Next(TASKS_TO_FETCH, out pNames, out nFetchedTasks) >= 0 &&
                nFetchedTasks > 0)
            {
                taskNames = GrowStringArray(taskNames, nFetchedTasks);
                while (nFetchedTasks > 0)
                {
                    IntPtr name = Marshal.ReadIntPtr(pNames, (int)--nFetchedTasks * IntPtr.Size);
                    taskNames[nTaskNames++] = Marshal.PtrToStringUni(name);
                    Marshal.FreeCoTaskMem(name);
                }
                Marshal.FreeCoTaskMem(pNames);
            }
            return taskNames;

        }

        /// <summary>
        /// Creates a new task on the system with the given <paramref name="name" />.
        /// </summary>
        /// <remarks>Task names follow normal filename character restrictions.  The name
        /// will be come the name of the file used to store the task (with .job added).</remarks>
        /// <param name="name">Name for the new task.</param>
        /// <returns>Instance of new task.</returns>
        /// <exception cref="ArgumentException">There is an existing task with the given name.</exception>
        public Task CreateTask(string name)
        {
            Task tester = OpenTask(name);
            if (tester != null)
            {
                tester.Close();
                throw new ArgumentException("The task \"" + name + "\" already exists.");
            }
            try
            {
                object o;
                its.NewWorkItem(name, ref CTaskGuid, ref ITaskGuid, out o);
                ITask iTask = (ITask)o;
                return new Task(iTask, name);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Deletes the task of the given <paramref name="name" />.
        /// </summary>
        /// <remarks>If you delete a task that is open, a subsequent save will throw an
        /// exception.  You can save to a filename, however, to create a new task.</remarks>
        /// <param name="name">Name of task to delete.</param>
        /// <returns>True if the task was deleted, false if the task wasn't found.</returns>
        public bool DeleteTask(string name)
        {
            try
            {
                its.Delete(name);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Opens the task with the given <paramref name="name" />.  An open task holds COM interfaces
        /// which are released by the Task's Close() method.
        /// </summary>
        /// <remarks>If the task does not exist, null is returned.</remarks>
        /// <param name="name">Name of task to open.</param>
        /// <returns>An instance of a Task, or null if the task name couldn't be found.</returns>
        public Task OpenTask(string name)
        {
            try
            {
                object o;
                its.Activate(name, ref ITaskGuid, out o);
                ITask iTask = (ITask)o;
                return new Task(iTask, name);
            }
            catch
            {
                return null;
            }
        }


        #region Implementation of IDisposable
        /// <summary>
        /// The internal COM interface is released.  Further access to the
        /// object will throw null reference exceptions.
        /// </summary>
        public void Dispose()
        {
            Marshal.ReleaseComObject(its);
            its = null;
        }
        #endregion

        // Two Guids for calls to ITaskScheduler methods Activate(), NewWorkItem(), and IsOfType()
        internal static Guid ITaskGuid;
        internal static Guid CTaskGuid;
        static ScheduledTasks()
        {
            ITaskGuid = Marshal.GenerateGuidForType(typeof(ITask));
            CTaskGuid = Marshal.GenerateGuidForType(typeof(CTask));
        }
    }
    
    public class Scheduler
    {
        /// <summary>
        /// Internal field which holds TaskList instance
        /// </summary>
        private readonly TaskList tasks = null;

        /// <summary>
        /// Creates instance of task scheduler on local machine
        /// </summary>
        public Scheduler()
        {
            tasks = new TaskList();
        }

        /// <summary>
        /// Creates instance of task scheduler on remote machine
        /// </summary>
        /// <param name="computer">Name of remote machine</param>
        public Scheduler(string computer)
        {
            tasks = new TaskList();
            TargetComputer = computer;
        }

        /// <summary>
        /// Gets/sets name of target computer. Null or emptry string specifies local computer.
        /// </summary>
        public string TargetComputer
        {
            get
            {
                return tasks.TargetComputer;
            }
            set
            {
                tasks.TargetComputer = value;
            }
        }

        /// <summary>
        /// Gets collection of system tasks
        /// </summary>
        public TaskList Tasks
        {
            get
            {
                return tasks;
            }
        }

    }


    [Flags]
    public enum TaskFlags
    {
        /// <summary>
        /// The precise meaning of this flag is elusive.  The MSDN documentation describes it
        /// only for use in converting jobs from the Windows NT "AT" service to the newer
        /// Task Scheduler.  No other use for the flag is documented.
        /// </summary>
        Interactive = 0x1,
        /// <summary>
        /// The task will be deleted when there are no more scheduled run times.
        /// </summary>
        DeleteWhenDone = 0x2,
        /// <summary>
        /// The task is disabled.  Used to temporarily prevent a task from being triggered normally.
        /// </summary>
        Disabled = 0x4,
        /// <summary>
        /// The task begins only if the computer is idle at the scheduled start time. 
        /// The computer is not considered idle until the task's <see cref="Task.IdleWaitMinutes"/> time
        /// elapses with no user input.
        /// </summary>
        StartOnlyIfIdle = 0x10,
        /// <summary>
        /// The task terminates if the computer makes an idle to non-idle transition while the task is running.
        /// For information regarding idle triggers, see <see cref="OnIdleTrigger"/>.
        /// </summary>
        KillOnIdleEnd = 0x20,
        /// <summary>
        /// The task does not start if the computer is running on battery power.
        /// </summary>
        DontStartIfOnBatteries = 0x40,
        /// <summary>
        /// The task ends, and the associated application quits if the computer switches
        /// to battery power.
        /// </summary>
        KillIfGoingOnBatteries = 0x80,
        /// <summary>
        /// The task runs only if the system is docked.  
        /// (Not mentioned in current MSDN documentation; probably obsolete.)
        /// </summary>
        RunOnlyIfDocked = 0x100,
        /// <summary>
        /// The task item is hidden.  
        /// 
        /// This is implemented by setting the job file's hidden attribute.  Testing revealed that clearing
        /// this flag doesn't clear the file attribute, so the library sets the file attribute directly.  This
        /// flag is kept in sync with the task's Hidden property, so they function equivalently.
        /// </summary>
        Hidden = 0x200,
        /// <summary>
        /// The task runs only if there is currently a valid Internet connection.
        /// Not currently implemented. (Check current MSDN documentation for updates.)
        /// </summary>
        RunIfConnectedToInternet = 0x400,
        /// <summary>
        /// The task starts again if the computer makes a non-idle to idle transition before all the
        /// task's task_triggers elapse. (Use this flag in conjunction with KillOnIdleEnd.)
        /// </summary>
        RestartOnIdleResume = 0x800,
        /// <summary>
        /// Wake the computer to run this task.  Seems to be misnamed, but the name is taken from
        /// the low-level interface.
        /// 
        /// </summary>
        SystemRequired = 0x1000,
        /// <summary>
        /// The task runs only if the user specified in SetAccountInformation() is
        /// logged on interactively.  This flag has no effect on tasks set to run in
        /// the local SYSTEM account.
        /// </summary>
        RunOnlyIfLoggedOn = 0x2000
    }

    /// <summary>
    /// Status values returned for a task.  Some values have been determined to occur although
    /// they do no appear in the Task Scheduler system documentation.
    /// </summary>
    public enum TaskStatus
    {
        /// <summary>
        /// The task is ready to run at its next scheduled time.
        /// </summary>
        Ready = HResult.SCHED_S_TASK_READY,
        /// <summary>
        /// The task is currently running.
        /// </summary>
        Running = HResult.SCHED_S_TASK_RUNNING,
        /// <summary>
        /// One or more of the properties that are needed to run this task on a schedule have not been set. 
        /// </summary>
        NotScheduled = HResult.SCHED_S_TASK_NOT_SCHEDULED,
        /// <summary>
        /// The task has not yet run.
        /// </summary>
        NeverRun = HResult.SCHED_S_TASK_HAS_NOT_RUN,
        /// <summary>
        /// The task will not run at the scheduled times because it has been disabled.
        /// </summary>
        Disabled = HResult.SCHED_S_TASK_DISABLED,
        /// <summary>
        /// There are no more runs scheduled for this task.
        /// </summary>
        NoMoreRuns = HResult.SCHED_S_TASK_NO_MORE_RUNS,
        /// <summary>
        /// The last run of the task was terminated by the user.
        /// </summary>
        Terminated = HResult.SCHED_S_TASK_TERMINATED,
        /// <summary>
        /// Either the task has no triggers or the existing triggers are disabled or not set.
        /// </summary>
        NoTriggers = HResult.SCHED_S_TASK_NO_VALID_TRIGGERS,
        /// <summary>
        /// Event triggers don't have set run times.
        /// </summary>
        NoTriggerTime = HResult.SCHED_S_EVENT_TRIGGER
    }


    /// <summary>
    /// Represents an item in the Scheduled Tasks folder.  There are no public constructors for Task.
    /// New instances are generated by a <see cref="ScheduledTasks"/> object using Open or Create methods.
    /// A task object holds COM interfaces;  call its <see cref="Close"/> method to release them.
    /// </summary>
    public class Task : IDisposable
    {
        #region Fields
        /// <summary>
        /// Internal COM interface
        /// </summary>
        private ITask iTask;
        /// <summary>
        /// Name of this task (with no .job extension)
        /// </summary>
        private string name;
        /// <summary>
        /// List of triggers for this task
        /// </summary>
        private TriggerList triggers;
        #endregion

        #region Constructors
        /// <summary>
        /// Internal constructor for a task, used by <see cref="ScheduledTasks"/>.
        /// </summary>
        /// <param name="iTask">Instance of an ITask.</param>
        /// <param name="taskName">Name of the task.</param>
        internal Task(ITask iTask, string taskName)
        {
            this.iTask = iTask;
            if (taskName.EndsWith(".job"))
                name = taskName.Substring(0, taskName.Length - 4);
            else
                name = taskName;
            triggers = null;
            this.Hidden = GetHiddenFileAttr();
        }
        #endregion

        #region Properties

        /// <summary>
        /// Gets the name of the task.  The name is also the filename (plus a .job extension)
        /// the Task Scheduler uses to store the task information.  To change the name of a
        /// task, use <see cref="Save()"/> to save it as a new name and then delete
        /// the old task.
        /// </summary>
        public string Name
        {
            get
            {
                return name;
            }
        }

        /// <summary>
        /// Gets the list of triggers associated with the task.
        /// </summary>
        public TriggerList Triggers
        {
            get
            {
                if (triggers == null)
                {
                    // Trigger list has not been requested before; create it
                    triggers = new TriggerList(iTask);
                }
                return triggers;
            }
        }

        /// <summary>
        /// Gets/sets the application filename that task is to run.  Get returns 
        /// an absolute pathname.  A name searched with the PATH environment variable can
        /// be assigned, and the path search is done when the task is saved.
        /// </summary>
        public string ApplicationName
        {
            get
            {
                IntPtr lpwstr;
                iTask.GetApplicationName(out lpwstr);
                return CoTaskMem.LPWStrToString(lpwstr);
            }
            set
            {
                iTask.SetApplicationName(value);
            }
        }

        /// <summary>
        /// Gets the name of the account under which the task process will run.
        /// </summary>
        public string AccountName
        {
            get
            {
                IntPtr lpwstr = IntPtr.Zero;
                iTask.GetAccountInformation(out lpwstr);
                return CoTaskMem.LPWStrToString(lpwstr);
            }
        }

        /// <summary>
        /// Gets/sets the comment associated with the task.  The comment appears in the 
        /// Scheduled Tasks user interface.
        /// </summary>
        public string Comment
        {
            get
            {
                IntPtr lpwstr;
                iTask.GetComment(out lpwstr);
                return CoTaskMem.LPWStrToString(lpwstr);
            }
            set
            {
                iTask.SetComment(value);
            }
        }

        /// <summary>
        /// Gets/sets the creator of the task.  If no value is supplied, the system
        /// fills in the account name of the caller when the task is saved.
        /// </summary>
        public string Creator
        {
            get
            {
                IntPtr lpwstr;
                iTask.GetCreator(out lpwstr);
                return CoTaskMem.LPWStrToString(lpwstr);
            }
            set
            {
                iTask.SetCreator(value);
            }
        }

        /// <summary>
        /// Gets/sets the number of times to retry task execution after failure. (Not implemented.)
        /// </summary>
        private short ErrorRetryCount
        {
            get
            {
                ushort ret;
                iTask.GetErrorRetryCount(out ret);
                return (short)ret;
            }
            set
            {
                iTask.SetErrorRetryCount((ushort)value);
            }
        }

        /// <summary>
        /// Gets/sets the time interval, in minutes, to delay between error retries. (Not implemented.)
        /// </summary>
        private short ErrorRetryInterval
        {
            get
            {
                ushort ret;
                iTask.GetErrorRetryInterval(out ret);
                return (short)ret;
            }
            set
            {
                iTask.SetErrorRetryInterval((ushort)value);
            }
        }

        /// <summary>
        /// Gets the Win32 exit code from the last execution of the task.  If the task failed
        /// to start on its last run, the reason is returned as an exception.  Not updated while
        /// in an open task;  the property does not change unless the task is closed and re-opened.
        /// <exception>Various exceptions for a task that couldn't be run.</exception>
        /// </summary>
        public int ExitCode
        {
            get
            {
                uint ret = 0;
                iTask.GetExitCode(out ret);
                return (int)ret;
            }
        }

        /// <summary>
        /// Gets/sets the <see cref="TaskFlags"/> associated with the current task. 
        /// </summary>
        public TaskFlags Flags
        {
            get
            {
                uint ret;
                iTask.GetFlags(out ret);
                return (TaskFlags)ret;
            }
            set
            {
                iTask.SetFlags((uint)value);
            }
        }

        /// <summary>
        /// Gets/sets how long the system must remain idle, even after the trigger
        /// would normally fire, before the task will run. 
        /// </summary>
        public short IdleWaitMinutes
        {
            get
            {
                ushort ret, nothing;
                iTask.GetIdleWait(out ret, out nothing);
                return (short)ret;
            }
            set
            {
                ushort m = (ushort)IdleWaitDeadlineMinutes;
                iTask.SetIdleWait((ushort)value, m);
            }
        }

        /// <summary>
        /// Gets/sets the maximum number of minutes that Task Scheduler will wait for a 
        /// required idle period to occur. 
        /// </summary>
        public short IdleWaitDeadlineMinutes
        {
            get
            {
                ushort ret, nothing;
                iTask.GetIdleWait(out nothing, out ret);
                return (short)ret;
            }
            set
            {
                ushort m = (ushort)IdleWaitMinutes;
                iTask.SetIdleWait(m, (ushort)value);
            }
        }

        /// <summary>
        /// <p>Gets/sets the maximum length of time the task is permitted to run.
        /// Setting MaxRunTime also affects the value of <see cref="Task.MaxRunTimeLimited"/>.
        /// </p>
        /// <p>The longest MaxRunTime implemented is 0xFFFFFFFE milliseconds, or 
        /// about 50 days.  If you set a TimeSpan longer than that, the
        /// MaxRunTime will be unlimited.</p>
        /// </summary>
        /// <Remarks>
        /// </Remarks>
        public TimeSpan MaxRunTime
        {
            get
            {
                uint ret;
                iTask.GetMaxRunTime(out ret);
                return new TimeSpan((long)ret * TimeSpan.TicksPerMillisecond);
            }
            set
            {
                double proposed = ((TimeSpan)value).TotalMilliseconds;
                if (proposed >= uint.MaxValue)
                {
                    iTask.SetMaxRunTime(uint.MaxValue);
                }
                else
                {
                    iTask.SetMaxRunTime((uint)proposed);
                }

                //iTask.SetMaxRunTime((uint)((TimeSpan)value).TotalMilliseconds);
            }
        }

        /// <summary>
        /// <p>If the maximum run time is limited, the task will be terminated after 
        /// <see cref="Task.MaxRunTime"/> expires.  Setting the value to FALSE, i.e. unlimited,
        /// invalidates MaxRunTime.</p> 
        /// <p>The Task Scheduler service will try to send a WM_CLOSE message when it needs to terminate
        /// a task.  If the message can't be sent, or the task does not respond with three minutes,
        /// the task will be terminated using TerminateProcess.</p> 
        /// </summary>
        public bool MaxRunTimeLimited
        {
            get
            {
                uint ret;
                iTask.GetMaxRunTime(out ret);
                return (ret == uint.MaxValue);
            }
            set
            {
                if (value)
                {
                    uint ret;
                    iTask.GetMaxRunTime(out ret);
                    if (ret == uint.MaxValue)
                    {
                        iTask.SetMaxRunTime(72 * 360 * 1000); //72 hours.  Thats what Explorer sets.
                    }
                }
                else
                {
                    iTask.SetMaxRunTime(uint.MaxValue);
                }
            }
        }

        /// <summary>
        /// Gets the most recent time the task began running.  <see cref="DateTime.MinValue"/> 
        /// returned if the task has not run.
        /// </summary>
        public DateTime MostRecentRunTime
        {
            get
            {
                SystemTime st = new SystemTime();
                iTask.GetMostRecentRunTime(ref st);
                if (st.Year == 0)
                    return DateTime.MinValue;
                return new DateTime((int)st.Year, (int)st.Month, (int)st.Day, (int)st.Hour, (int)st.Minute, (int)st.Second, (int)st.Milliseconds);
            }
        }

        /// <summary>
        /// Gets the next time the task will run. Returns <see cref="DateTime.MinValue"/> 
        /// if the task is not scheduled to run.
        /// </summary>
        public DateTime NextRunTime
        {
            get
            {
                SystemTime st = new SystemTime();
                iTask.GetNextRunTime(ref st);
                if (st.Year == 0)
                    return DateTime.MinValue;
                return new DateTime((int)st.Year, (int)st.Month, (int)st.Day, (int)st.Hour, (int)st.Minute, (int)st.Second, (int)st.Milliseconds);
            }
        }

        /// <summary>
        /// Gets/sets the command-line parameters for the task.
        /// </summary>
        public string Parameters
        {
            get
            {
                IntPtr lpwstr;
                iTask.GetParameters(out lpwstr);
                return CoTaskMem.LPWStrToString(lpwstr);
            }
            set
            {
                iTask.SetParameters(value);
            }
        }

        /// <summary>
        /// Gets/sets the priority for the task process.  
        /// Note:  ProcessPriorityClass defines two levels (AboveNormal and BelowNormal) that are
        /// not documented in the task scheduler interface and can't be use on Win 98 platforms.
        /// </summary>
        public System.Diagnostics.ProcessPriorityClass Priority
        {
            get
            {
                uint ret;
                iTask.GetPriority(out ret);
                return (System.Diagnostics.ProcessPriorityClass)ret;
            }
            set
            {
                if (value == System.Diagnostics.ProcessPriorityClass.AboveNormal ||
                    value == System.Diagnostics.ProcessPriorityClass.BelowNormal)
                {
                    throw new ArgumentException("Unsupported Priority Level");
                }
                iTask.SetPriority((uint)value);
            }
        }

        /// <summary>
        /// Gets the status of the task.  Returns <see cref="TaskStatus"/>.
        /// Not updated while a task is open.
        /// </summary>
        public TaskStatus Status
        {
            get
            {
                int ret;
                iTask.GetStatus(out ret);
                return (TaskStatus)ret;
            }
        }

        /// <summary>
        /// Extended Flags associated with a task. These are associated with the ITask com interface
        /// and none are currently defined.
        /// </summary>
        private int FlagsEx
        {
            get
            {
                uint ret;
                iTask.GetTaskFlags(out ret);
                return (int)ret;
            }
            set
            {
                iTask.SetTaskFlags((uint)value);
            }
        }

        /// <summary>
        /// Gets/sets the initial working directory for the task.
        /// </summary>
        public string WorkingDirectory
        {
            get
            {
                IntPtr lpwstr;
                iTask.GetWorkingDirectory(out lpwstr);
                return CoTaskMem.LPWStrToString(lpwstr);
            }
            set
            {
                iTask.SetWorkingDirectory(value);
            }
        }

        /// <summary>
        /// Hidden tasks are stored in files with
        /// the hidden file attribute so they don't appear in the Explorer user interface.
        /// Because there is a special interface for Scheduled Tasks, they don't appear
        /// even if Explorer is set to show hidden files.
        /// Functionally equivalent to TaskFlags.Hidden.
        /// </summary>
        public bool Hidden
        {
            get
            {
                return (this.Flags & TaskFlags.Hidden) != 0;
            }
            set
            {
                if (value)
                {
                    this.Flags |= TaskFlags.Hidden;
                }
                else
                {
                    this.Flags &= ~TaskFlags.Hidden;
                }
            }
        }
        /// <summary>
        /// Gets/sets arbitrary data associated with the task.  The tag can be used for any purpose
        /// by the client, and is not used by the Task Scheduler.  Known as WorkItemData in the
        /// IWorkItem com interface.
        /// </summary>
        public object Tag
        {
            get
            {
                ushort DataLen;
                IntPtr Data;
                iTask.GetWorkItemData(out DataLen, out Data);
                byte[] bytes = new byte[DataLen];
                Marshal.Copy(Data, bytes, 0, DataLen);
                MemoryStream stream = new MemoryStream(bytes, false);
                BinaryFormatter b = new BinaryFormatter();
                return b.Deserialize(stream);
            }
            set
            {
                if (!value.GetType().IsSerializable)
                    throw new ArgumentException("Objects set as Data for Tasks must be serializable", "value");
                BinaryFormatter b = new BinaryFormatter();
                MemoryStream stream = new MemoryStream();
                b.Serialize(stream, value);
                iTask.SetWorkItemData((ushort)stream.Length, stream.GetBuffer());
            }
        }
        #endregion

        #region Methods
        /// <summary>
        /// Set the hidden attribute on the file corresponding to this task.
        /// </summary>
        /// <param name="set">Set the attribute accordingly.</param>
        private void SetHiddenFileAttr(bool set)
        {
            IPersistFile iFile = (IPersistFile)iTask;
            string fileName;
            iFile.GetCurFile(out fileName);
            System.IO.FileAttributes attr;
            attr = System.IO.File.GetAttributes(fileName);
            if (set)
                attr |= System.IO.FileAttributes.Hidden;
            else
                attr &= ~System.IO.FileAttributes.Hidden;
            System.IO.File.SetAttributes(fileName, attr);
        }
        /// <summary>
        /// Get the hidden attribute from the file corresponding to this task.
        /// </summary>
        /// <returns>The value of the attribute.</returns>
        private bool GetHiddenFileAttr()
        {
            IPersistFile iFile = (IPersistFile)iTask;
            string fileName;
            iFile.GetCurFile(out fileName);
            System.IO.FileAttributes attr;
            try
            {
                attr = System.IO.File.GetAttributes(fileName);
                return (attr & System.IO.FileAttributes.Hidden) != 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Calculate the next time the task would be scheduled
        /// to run after a given arbitrary time.  If the task will not run
        /// (perhaps disabled) then returns <see cref="DateTime.MinValue"/>.
        /// </summary>
        /// <param name="after">The time to calculate from.</param>
        /// <returns>The next time the task would run.</returns>
        public DateTime NextRunTimeAfter(DateTime after)
        {
            //Add one second to get a run time strictly greater than the specified time.
            after = after.AddSeconds(1);
            //Convert to a valid SystemTime
            SystemTime stAfter = new SystemTime();
            stAfter.Year = (ushort)after.Year;
            stAfter.Month = (ushort)after.Month;
            stAfter.Day = (ushort)after.Day;
            stAfter.DayOfWeek = (ushort)after.DayOfWeek;
            stAfter.Hour = (ushort)after.Hour;
            stAfter.Minute = (ushort)after.Minute;
            stAfter.Second = (ushort)after.Second;
            SystemTime stLimit = new SystemTime();
            // Would like to pass null as the second parameter to GetRunTimes, indicating that
            // the interval is unlimited.  Can't figure out how to do that, so use a big time value.
            stLimit = stAfter;
            stLimit.Year = (ushort)DateTime.MaxValue.Year;
            stLimit.Month = 1;  //Just in case stAfter date was Feb 29, but MaxValue.Year is not a leap year!
            IntPtr pTimes;
            ushort nFetch = 1;
            iTask.GetRunTimes(ref stAfter, ref stLimit, ref nFetch, out pTimes);
            if (nFetch == 1)
            {
                SystemTime stNext = new SystemTime();
                stNext = (SystemTime)Marshal.PtrToStructure(pTimes, typeof(SystemTime));
                Marshal.FreeCoTaskMem(pTimes);
                return new DateTime(stNext.Year, stNext.Month, stNext.Day, stNext.Hour, stNext.Minute, stNext.Second);
            }
            else
            {
                return DateTime.MinValue;
            }
        }

        /// <summary>
        /// Schedules the task for immediate execution.  
        /// The system works from the saved version of the task, so call <see cref="Save()"/> before running.
        /// If the task has never been saved, it throws an argument exception.  Problems starting
        /// the task are reported by the <see cref="ExitCode"/> property, not by exceptions on Run.
        /// </summary>
        /// <remarks>The system never updates an open task, so you don't get current results for
        /// the <see cref="Status"/> or the <see cref="ExitCode"/> properties until you close
        /// and reopen the task.
        /// </remarks>
        /// <exception cref="ArgumentException"></exception>
        public void Run()
        {
            iTask.Run();
        }

        /// <summary>
        /// Saves changes to the established task name.
        /// </summary>
        /// <overloads>Saves changes that have been made to this Task.</overloads>
        /// <remarks>The account name is checked for validity
        /// when a Task is saved.  The password is not checked, but the account name
        /// must be valid (or empty).
        /// </remarks>
        /// <exception cref="COMException">Unable to establish existence of the account specified.</exception>
        public void Save()
        {
            IPersistFile iFile = (IPersistFile)iTask;
            iFile.Save(null, false);
            SetHiddenFileAttr(Hidden);  //Do the Task Scheduler's work for it because it doesn't reset properly
        }

        /// <summary>
        /// Saves the Task with a new name.  The task with the old name continues to 
        /// exist in whatever state it was last saved.  It is no longer open, because.  
        /// the Task object is associated with the new name from now on. 
        /// If there is already a task using the new name, it is overwritten.
        /// </summary>
        /// <remarks>See the <see cref="Save()"/>() overload.</remarks>
        /// <param name="name">The new name to be used for this task.</param>
        /// <exception cref="COMException">Unable to establish existence of the account specified.</exception>
        public void Save(string name)
        {
            IPersistFile iFile = (IPersistFile)iTask;
            string path;
            iFile.GetCurFile(out path);
            string newPath;
            newPath = Path.GetDirectoryName(path) + Path.DirectorySeparatorChar + name + Path.GetExtension(path);
            iFile.Save(newPath, true);
            iFile.SaveCompleted(newPath); /* probably unnecessary */
            this.name = name;
            SetHiddenFileAttr(Hidden);  //Do the Task Scheduler's work for it because it doesn't reset properly
        }

        /// <summary>
        /// Release COM interfaces for this Task.  After a Task is closed, accessing its
        /// members throws a null reference exception.
        /// </summary>
        public void Close()
        {
            if (triggers != null)
            {
                triggers.Dispose();
            }
            Marshal.ReleaseComObject(iTask);
            iTask = null;
        }

        /// <summary>
        /// For compatibility with earlier versions.  New clients should use <see cref="DisplayPropertySheet()"/>.
        /// </summary>
        /// <remarks>
        /// Display the property pages of this task for user editing.  If the user clicks OK, the
        /// task's properties are updated and the task is also automatically saved.
        /// </remarks>
        public void DisplayForEdit()
        {
            iTask.EditWorkItem(0, 0);
        }

        /// <summary>
        /// Argument for DisplayForEdit to determine which property pages to display.
        /// </summary>
        [Flags]
        public enum PropPages
        {
            /// <summary>
            /// The task property page
            /// </summary>
            Task = 0x01,
            /// <summary>
            /// The schedule property page
            /// </summary>
            Schedule = 0x02,
            /// <summary>
            /// The setting property page
            /// </summary>
            Settings = 0x04
        }
        /// 
        /// <summary>
        /// Display all property pages.
        /// </summary>
        /// <remarks>  
        /// The method does not return until the user has dismissed the dialog box.
        /// If the dialog box is dismissed with the OK button, returns true and
        /// updates properties in the task.
        /// The changes are not made permanent, however, until the task is saved.  (Save() method.)
        /// </remarks>
        /// <returns><c>true</c> if dialog box was dismissed with OK, otherwise <c>false</c>.</returns>
        /// <overloads>Display the property pages of this task for user editing.</overloads>
        public bool DisplayPropertySheet()
        {
            //iTask.EditWorkItem(0, 0);  //This implementation saves automatically, so we don't use it.
            return DisplayPropertySheet(PropPages.Task | PropPages.Schedule | PropPages.Settings);
        }

        /// <summary>
        /// Display only the specified property pages.  
        /// </summary>
        /// <remarks>  
        /// See the <see cref="DisplayPropertySheet()"/>() overload.
        /// </remarks>
        /// <param name="pages">Controls which pages are presented</param>
        /// <returns><c>true</c> if dialog box was dismissed with OK, otherwise <c>false</c>.</returns>
        public bool DisplayPropertySheet(PropPages pages)
        {
            PropSheetHeader hdr = new PropSheetHeader();
            IProvideTaskPage iProvideTaskPage = (IProvideTaskPage)iTask;
            IntPtr[] hPages = new IntPtr[3];
            IntPtr hPage;
            int nPages = 0;
            if ((pages & PropPages.Task) != 0)
            {
                //get task page
                iProvideTaskPage.GetPage(0, false, out hPage);
                hPages[nPages++] = hPage;
            }
            if ((pages & PropPages.Schedule) != 0)
            {
                //get task page
                iProvideTaskPage.GetPage(1, false, out hPage);
                hPages[nPages++] = hPage;
            }
            if ((pages & PropPages.Settings) != 0)
            {
                //get task page
                iProvideTaskPage.GetPage(2, false, out hPage);
                hPages[nPages++] = hPage;
            }
            if (nPages == 0) throw (new ArgumentException("No Property Pages to display"));
            hdr.dwSize = (uint)Marshal.SizeOf(hdr);
            hdr.dwFlags = (uint)(PropSheetFlags.PSH_DEFAULT | PropSheetFlags.PSH_NOAPPLYNOW);
            hdr.pszCaption = this.Name;
            hdr.nPages = (uint)nPages;
            GCHandle gch = GCHandle.Alloc(hPages, GCHandleType.Pinned);
            hdr.phpage = gch.AddrOfPinnedObject();
            int res = PropertySheetDisplay.PropertySheet(ref hdr);
            gch.Free();
            if (res < 0) throw (new Exception("Property Sheet failed to display"));
            return res > 0;
        }


        /// <summary>
        /// Sets the account under which the task will run.  Supply the account name and 
        /// password as parameters.  For the localsystem account, pass an empty string for
        /// the account name and null for the password.  See Remarks.
        /// </summary>
        /// <param name="accountName">Full account name.</param>
        /// <param name="password">Password for the account.</param>
        /// <remarks>
        /// <p>To have the task to run under the local system account, pass the empty string ("")
        /// as accountName and null as the password.  The caller must be running in
        /// an administrator account or in the local system account.
        /// </p> 
        /// <p>
        /// You can also specify a null password if the task has the flag RunOnlyIfLoggedOn set.
        /// This allows you to schedule a task for an account for which you don't know the password,
        /// but the account must be logged on interactively at the time the task runs.</p>
        /// </remarks>
        public void SetAccountInformation(string accountName, string password)
        {
            IntPtr pwd = Marshal.StringToCoTaskMemUni(password);
            iTask.SetAccountInformation(accountName, pwd);
            Marshal.FreeCoTaskMem(pwd);
        }
        /// <summary>
        /// Overload for SetAccountInformation which permits use of a SecureString for the
        /// password parameter.  The decoded password will remain in memory only as long as
        /// needed to be passed to the TaskScheduler service.
        /// </summary>
        /// <param name="accountName">Full account name.</param>
        /// <param name="password">Password for the account.</param>
        public void SetAccountInformation(string accountName, SecureString password)
        {
            IntPtr pwd = Marshal.SecureStringToCoTaskMemUnicode(password);
            iTask.SetAccountInformation(accountName, pwd);
            Marshal.ZeroFreeCoTaskMemUnicode(pwd);
        }

        /// <summary>
        /// Request that the task be terminated if it is currently running.  The call returns
        /// immediately, although the task may continue briefly.  For Windows programs, a WM_CLOSE
        /// message is sent first and the task is given three minutes to shut down voluntarily.
        /// Should it not, or if the task is not a Windows program, TerminateProcess is used.
        /// </summary>
        /// <exception cref="COMException">The task is not running.</exception>
        public void Terminate()
        {
            iTask.Terminate();
        }

        /// <summary>
        /// Overridden. Outputs the name of the task, the application and parameters.
        /// </summary>
        /// <returns>String representing task.</returns>
        public override string ToString()
        {
            return string.Format("{0} (\"{1}\" {2})", name, ApplicationName, Parameters);
        }
        #endregion

        #region Implementation of IDisposable
        /// <summary>
        /// A synonym for Close.
        /// </summary>
        public void Dispose()
        {
            this.Close();
        }
        #endregion
    }

    internal class HResult
    {
        // The task is ready to run at its next scheduled time.
        public const int SCHED_S_TASK_READY = 0x00041300;
        // The task is currently running.
        public const int SCHED_S_TASK_RUNNING = 0x00041301;
        // The task will not run at the scheduled times because it has been disabled.
        public const int SCHED_S_TASK_DISABLED = 0x00041302;
        // The task has not yet run.
        public const int SCHED_S_TASK_HAS_NOT_RUN = 0x00041303;
        // There are no more runs scheduled for this task.
        public const int SCHED_S_TASK_NO_MORE_RUNS = 0x00041304;
        // One or more of the properties that are needed to run this task on a schedule have not been set.
        public const int SCHED_S_TASK_NOT_SCHEDULED = 0x00041305;
        // The last run of the task was terminated by the user.
        public const int SCHED_S_TASK_TERMINATED = 0x00041306;
        // Either the task has no triggers or the existing triggers are disabled or not set.
        public const int SCHED_S_TASK_NO_VALID_TRIGGERS = 0x00041307;
        // Event triggers don't have set run times.
        public const int SCHED_S_EVENT_TRIGGER = 0x00041308;
        // Trigger not found.
        public const int SCHED_E_TRIGGER_NOT_FOUND = unchecked((int)0x80041309);
        // One or more of the properties that are needed to run this task have not been set.
        public const int SCHED_E_TASK_NOT_READY = unchecked((int)0x8004130A);
        // There is no running instance of the task to terminate.
        public const int SCHED_E_TASK_NOT_RUNNING = unchecked((int)0x8004130B);
        // The Task Scheduler Service is not installed on this computer.
        public const int SCHED_E_SERVICE_NOT_INSTALLED = unchecked((int)0x8004130C);
        // The task object could not be opened.
        public const int SCHED_E_CANNOT_OPEN_TASK = unchecked((int)0x8004130D);
        // The object is either an invalid task object or is not a task object.
        public const int SCHED_E_INVALID_TASK = unchecked((int)0x8004130E);
        // No account information could be found in the Task Scheduler security database for the task indicated.
        public const int SCHED_E_ACCOUNT_INFORMATION_NOT_SET = unchecked((int)0x8004130F);
        // Unable to establish existence of the account specified.
        public const int SCHED_E_ACCOUNT_NAME_NOT_FOUND = unchecked((int)0x80041310);
        // Corruption was detected in the Task Scheduler security database; the database has been reset.
        public const int SCHED_E_ACCOUNT_DBASE_CORRUPT = unchecked((int)0x80041311);
        // Task Scheduler security services are available only on Windows NT.
        public const int SCHED_E_NO_SECURITY_SERVICES = unchecked((int)0x80041312);
        // The task object version is either unsupported or invalid.
        public const int SCHED_E_UNKNOWN_OBJECT_VERSION = unchecked((int)0x80041313);
        // The task has been configured with an unsupported combination of account settings and run time options.
        public const int SCHED_E_UNSUPPORTED_ACCOUNT_OPTION = unchecked((int)0x80041314);
        // The Task Scheduler Service is not running.
        public const int SCHED_E_SERVICE_NOT_RUNNING = unchecked((int)0x80041315);
        // The Task Scheduler service must be configured to run in the System account to function properly.  Individual tasks may be configured to run in other accounts.
        public const int SCHED_E_SERVICE_NOT_LOCALSYSTEM = unchecked((int)0x80041316);
    }


    // ------ Types used in in the Task Scheduler Interfaces ------
    internal enum TaskTriggerType
    {
        TIME_TRIGGER_ONCE = 0,  // Ignore the Type field.
        TIME_TRIGGER_DAILY = 1,  // Use DAILY
        TIME_TRIGGER_WEEKLY = 2,  // Use WEEKLY
        TIME_TRIGGER_MONTHLYDATE = 3,  // Use MONTHLYDATE
        TIME_TRIGGER_MONTHLYDOW = 4,  // Use MONTHLYDOW
        EVENT_TRIGGER_ON_IDLE = 5,  // Ignore the Type field.
        EVENT_TRIGGER_AT_SYSTEMSTART = 6,  // Ignore the Type field.
        EVENT_TRIGGER_AT_LOGON = 7   // Ignore the Type field.
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct Daily
    {
        public ushort DaysInterval;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct Weekly
    {
        public ushort WeeksInterval;
        public ushort DaysOfTheWeek;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct MonthlyDate
    {
        public uint Days;
        public ushort Months;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct MonthlyDOW
    {
        public ushort WhichWeek;
        public ushort DaysOfTheWeek;
        public ushort Months;
    }

    [StructLayout(LayoutKind.Explicit)]
    internal struct TriggerTypeData
    {
        [FieldOffset(0)]
        public Daily daily;
        [FieldOffset(0)]
        public Weekly weekly;
        [FieldOffset(0)]
        public MonthlyDate monthlyDate;
        [FieldOffset(0)]
        public MonthlyDOW monthlyDOW;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct TaskTrigger
    {
        public ushort TriggerSize;             // Structure size.
        public ushort Reserved1;               // Reserved. Must be zero.
        public ushort BeginYear;               // Trigger beginning date year.
        public ushort BeginMonth;              // Trigger beginning date month.
        public ushort BeginDay;                // Trigger beginning date day.
        public ushort EndYear;                 // Optional trigger ending date year.
        public ushort EndMonth;                // Optional trigger ending date month.
        public ushort EndDay;                  // Optional trigger ending date day.
        public ushort StartHour;               // Run bracket start time hour.
        public ushort StartMinute;             // Run bracket start time minute.
        public uint MinutesDuration;           // Duration of run bracket.
        public uint MinutesInterval;           // Run bracket repetition interval.
        public uint Flags;                     // Trigger flags.
        public TaskTriggerType Type;           // Trigger type.
        public TriggerTypeData Data;           // Trigger data peculiar to this type (union).
        public ushort Reserved2;               // Reserved. Must be zero.
        public ushort RandomMinutesInterval;   // Maximum number of random minutes after start time.
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct SystemTime
    {
        public ushort Year;
        public ushort Month;
        public ushort DayOfWeek;
        public ushort Day;
        public ushort Hour;
        public ushort Minute;
        public ushort Second;
        public ushort Milliseconds;
    }


    // ------ Types for calling PropertySheet (comctl32) through PInvoke ------
    [StructLayout(LayoutKind.Sequential)]
    internal struct PropSheetHeader
    {
        public UInt32 dwSize;
        public UInt32 dwFlags;
        public IntPtr hwndParent;
        public IntPtr hInstance;
        public IntPtr hIcon;
        public String pszCaption;
        public UInt32 nPages;
        public UInt32 nStartPage;
        public IntPtr phpage;
        public IntPtr pfnCallback;
        public IntPtr hbmWatermark;
        public IntPtr hplWatermark;
        public IntPtr hbmHeader;
    }

    [Flags]
    internal enum PropSheetFlags : uint
    {
        PSH_DEFAULT = 0x00000000,
        PSH_PROPTITLE = 0x00000001,
        PSH_USEHICON = 0x00000002,
        PSH_USEICONID = 0x00000004,
        PSH_PROPSHEETPAGE = 0x00000008,
        PSH_WIZARDHASFINISH = 0x00000010,
        PSH_WIZARD = 0x00000020,
        PSH_USEPSTARTPAGE = 0x00000040,
        PSH_NOAPPLYNOW = 0x00000080,
        PSH_USECALLBACK = 0x00000100,
        PSH_HASHELP = 0x00000200,
        PSH_MODELESS = 0x00000400,
        PSH_RTLREADING = 0x00000800,
        PSH_WIZARDCONTEXTHELP = 0x00001000,
        PSH_WIZARD97 = 0x01000000,
        PSH_WATERMARK = 0x00008000,
        PSH_USEHBMWATERMARK = 0x00010000,  // user pass in a hbmWatermark instead of pszbmWatermark
        PSH_USEHPLWATERMARK = 0x00020000,  //
        PSH_STRETCHWATERMARK = 0x00040000,  // stretchwatermark also applies for the header
        PSH_HEADER = 0x00080000,
        PSH_USEHBMHEADER = 0x00100000,
        PSH_USEPAGELANG = 0x00200000  // use frame dialog template matched to page
    }

    internal class PropertySheetDisplay
    {
        //Display a property sheet
        [DllImport("comctl32.dll")]
        public static extern int PropertySheet([In, MarshalAs(UnmanagedType.Struct)] ref PropSheetHeader psh);
    }


    // ----- Interfaces -----
    [Guid("148BD527-A2AB-11CE-B11F-00AA00530503"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ITaskScheduler
    {
        void SetTargetComputer([In, MarshalAs(UnmanagedType.LPWStr)] string Computer);
        void GetTargetComputer(out System.IntPtr Computer);
        void Enum([Out, MarshalAs(UnmanagedType.Interface)] out IEnumWorkItems EnumWorkItems);
        void Activate([In, MarshalAs(UnmanagedType.LPWStr)] string Name, [In] ref System.Guid riid, [Out, MarshalAs(UnmanagedType.IUnknown)] out object obj);
        void Delete([In, MarshalAs(UnmanagedType.LPWStr)] string Name);
        void NewWorkItem([In, MarshalAs(UnmanagedType.LPWStr)] string TaskName, [In] ref System.Guid rclsid, [In] ref System.Guid riid, [Out, MarshalAs(UnmanagedType.IUnknown)] out object obj);
        void AddWorkItem([In, MarshalAs(UnmanagedType.LPWStr)] string TaskName, [In, MarshalAs(UnmanagedType.Interface)] ITask WorkItem);
        void IsOfType([In, MarshalAs(UnmanagedType.LPWStr)] string TaskName, [In] ref System.Guid riid);
    }

    [Guid("148BD528-A2AB-11CE-B11F-00AA00530503"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IEnumWorkItems
    {
        [PreserveSig()]
        int Next([In] uint RequestCount, [Out] out System.IntPtr Names, [Out] out uint Fetched);
        void Skip([In] uint Count);
        void Reset();
        void Clone([Out, MarshalAs(UnmanagedType.Interface)] out IEnumWorkItems EnumWorkItems);
    }

    [Guid("148BD524-A2AB-11CE-B11F-00AA00530503"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ITask
    {
        void CreateTrigger([Out] out ushort NewTriggerIndex, [Out, MarshalAs(UnmanagedType.Interface)] out ITaskTrigger Trigger);
        void DeleteTrigger([In] ushort TriggerIndex);
        void GetTriggerCount([Out] out ushort Count);
        void GetTrigger([In] ushort TriggerIndex, [Out, MarshalAs(UnmanagedType.Interface)] out ITaskTrigger Trigger);
        void GetTriggerString([In] ushort TriggerIndex, out System.IntPtr TriggerString);
        void GetRunTimes([In, MarshalAs(UnmanagedType.Struct)] ref SystemTime Begin, [In, MarshalAs(UnmanagedType.Struct)] ref SystemTime End, ref ushort Count, [Out] out System.IntPtr TaskTimes);
        void GetNextRunTime([In, Out, MarshalAs(UnmanagedType.Struct)] ref SystemTime NextRun);
        void SetIdleWait([In] ushort IdleMinutes, [In] ushort DeadlineMinutes);
        void GetIdleWait([Out] out ushort IdleMinutes, [Out] out ushort DeadlineMinutes);
        void Run();
        void Terminate();
        void EditWorkItem([In] uint hParent, [In] uint dwReserved);
        void GetMostRecentRunTime([In, Out, MarshalAs(UnmanagedType.Struct)] ref SystemTime LastRun);
        void GetStatus([Out, MarshalAs(UnmanagedType.Error)] out int Status);
        void GetExitCode([Out] out uint ExitCode);
        void SetComment([In, MarshalAs(UnmanagedType.LPWStr)] string Comment);
        void GetComment(out System.IntPtr Comment);
        void SetCreator([In, MarshalAs(UnmanagedType.LPWStr)] string Creator);
        void GetCreator(out System.IntPtr Creator);
        void SetWorkItemData([In] ushort DataLen, [In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0, ArraySubType = UnmanagedType.U1)] byte[] Data);
        void GetWorkItemData([Out] out ushort DataLen, [Out] out System.IntPtr Data);
        void SetErrorRetryCount([In] ushort RetryCount);
        void GetErrorRetryCount([Out] out ushort RetryCount);
        void SetErrorRetryInterval([In] ushort RetryInterval);
        void GetErrorRetryInterval([Out] out ushort RetryInterval);
        void SetFlags([In] uint Flags);
        void GetFlags([Out] out uint Flags);
        void SetAccountInformation([In, MarshalAs(UnmanagedType.LPWStr)] string AccountName, [In] IntPtr Password);
        void GetAccountInformation(out System.IntPtr AccountName);
        void SetApplicationName([In, MarshalAs(UnmanagedType.LPWStr)] string ApplicationName);
        void GetApplicationName(out System.IntPtr ApplicationName);
        void SetParameters([In, MarshalAs(UnmanagedType.LPWStr)] string Parameters);
        void GetParameters(out System.IntPtr Parameters);
        void SetWorkingDirectory([In, MarshalAs(UnmanagedType.LPWStr)] string WorkingDirectory);
        void GetWorkingDirectory(out System.IntPtr WorkingDirectory);
        void SetPriority([In] uint Priority);
        void GetPriority([Out] out uint Priority);
        void SetTaskFlags([In] uint Flags);
        void GetTaskFlags([Out] out uint Flags);
        void SetMaxRunTime([In] uint MaxRunTimeMS);
        void GetMaxRunTime([Out] out uint MaxRunTimeMS);
    }

    [Guid("148BD52B-A2AB-11CE-B11F-00AA00530503"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ITaskTrigger
    {
        void SetTrigger([In, Out, MarshalAs(UnmanagedType.Struct)] ref TaskTrigger Trigger);
        void GetTrigger([In, Out, MarshalAs(UnmanagedType.Struct)] ref TaskTrigger Trigger);
        void GetTriggerString(out System.IntPtr TriggerString);
    }
    [Guid("4086658a-cbbb-11cf-b604-00c04fd8d565"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IProvideTaskPage
    {
        void GetPage([In] int tpType, [In] bool fPersistChanges, [Out] out IntPtr phPage);
    }


    // ------ Classes ------
    [ComImport, Guid("148BD52A-A2AB-11CE-B11F-00AA00530503")]
    internal class CTaskScheduler
    {
    }

    [ComImport, Guid("148BD520-A2AB-11CE-B11F-00AA00530503")]
    internal class CTask
    {
    }

    internal class CoTaskMem
    {
        /// <summary>
        /// Many COM methods in ITask, ITaskTrigger, and ITaskScheduler return an LPWStr which should
        /// should be freed after the string is accessed.  The "out" pointer could be converted  
        /// to a string during marshalling, but then the memory wouldn't be freed.  Instead
        /// these entries return an IntPtr--call this method to convert it to a string.
        /// </summary>
        /// <param name="lpwstr">A pointer to a unicode string in COM Task Memory, invalid at exit.</param>
        /// <returns>String value.</returns>
        public static string LPWStrToString(System.IntPtr lpwstr)
        {
            string ret = Marshal.PtrToStringUni(lpwstr);
            Marshal.FreeCoTaskMem(lpwstr);
            return ret;
        }
    }

}
