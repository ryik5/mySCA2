using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ASTA.PersonDefinitions;

namespace ASTA
{
    class LoadingRegistrationOfIdCards:INotifyCollectionChanged
    {
        public Visitors visitors = new Visitors();

        //transfer to Main PO
        private void ContentCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {

            //todo
            //reload collection into datagrid
            //todo it only for today
            // dataGridView1.Visible = false;
        //    _dataGridViewShowData(visitors.collection);
            //  dataGridView1.Visible = true;
        }
        // Объявляем делегат
        public delegate void StatusStateHandler(string message);
        // Событие, возникающее при добавлении события
        public event StatusStateHandler StatusMain;
        public event StatusStateHandler StatusTrace;


        private static object lockerLoadingTimeInpsOutps;
        private static object lockerLoadingInsOuts = new object();
        private static string startLoadingTimeInpsOutps = "00:00:00";
        private static string startLoadingDayInpsOutps = "dddddday";
        private static bool checkInputsOutputs = true;

        public event NotifyCollectionChangedEventHandler CollectionChanged;

        private void LoadInputsOutputsOfVisitors(string wholeDay)
        {

            visitors?.collection?.Clear();
            visitors = new Visitors();
            visitors.collection.CollectionChanged += ContentCollectionChanged; //subscribe on changing data in visitors
            List<Visitor> visitorsTillNow;

            lockerLoadingTimeInpsOutps = new object();
            startLoadingDayInpsOutps = wholeDay;
            startLoadingTimeInpsOutps = "00:00:00";
            checkInputsOutputs = true;
            bool firstStage = true;

            lockerLoadingInsOuts = new object();
            StartStopTimer startStopTimer = new StartStopTimer(15);
            int timesChecking = 10;
            do
            {
                lock (lockerLoadingInsOuts)
                {
                    visitorsTillNow = GetInputsOutputs();
                    if (firstStage)
                    {
                        visitors.collection.AddRange(visitorsTillNow);
                        firstStage = false;
                    }
                    else
                    {
                        if (visitorsTillNow?.Count > 0)
                        {
                            visitorsTillNow.Reverse();
                            foreach (var visitor in visitorsTillNow)
                                visitors.Add(visitor, 0);
                        }
                    }
                }
                timesChecking--;
                if (timesChecking < 0)
                { checkInputsOutputs = false; }

                StatusMain("Данные собраны: " + (10 - timesChecking));

                startStopTimer.WaitTime();

            } while (checkInputsOutputs);

            StatusMain( "Данные собраны");
        }
        DateTimeConvertor timeConvertor = new DateTimeConvertor();

        private List<Visitor> GetInputsOutputs()
        {
            List<Visitor> visitors;
            bool startTimeNotSet = true;

            SideOfPassagePoint sideOfPassagePoint;

            string startDay = startLoadingDayInpsOutps + " " + startLoadingTimeInpsOutps; //string startDay = day + " 00:00:00";
            string endDay = startLoadingDayInpsOutps + " 23:59:59";
            string time, date, fullPointName, fio, action, action_descr, fac, card;
            int idCard = 0; string idCardDescr;

            string query = "SELECT p.param0 as param0, p.param1 as param1, p.action as action, p.objid as objid, p.objtype, " +
                " pe.tabnum as nav, pe.facility_code as fac, pe.card as card, " +
                " CONVERT(varchar, p.date, 120) AS date, CONVERT(varchar, p.time, 114) AS time, p.time dateTimeRegistration" +
                " FROM protocol p " +
                " LEFT JOIN OBJ_PERSON pe ON  p.param1=pe.id " +
                " where p.objtype like 'ABC_ARC_READER' AND p.param0 like '%%' " +
                " AND date > '" + startDay + "' AND date <= '" + endDay + "' " +
                " ORDER BY p.time DESC";

           StatusTrace("stringConnection: " + sqlServerConnectionString);
            StatusTrace("query: " + query);

            using (SqlDbReader sqlDbTableReader = new SqlDbReader(sqlServerConnectionString))
            {
                visitors = new List<Visitor>();

                System.Data.SqlClient.SqlDataReader sqlData = sqlDbTableReader.GetData(query);
                foreach (DbDataRecord record in sqlData)
                {
                    //look for PassagePoint
                    fullPointName = record["objid"]?.ToString()?.Trim();
                    sideOfPassagePoint = collectionOfPassagePoints.GetPoint(fullPointName);

                    //look for FIO
                    fio = record["param0"]?.ToString()?.Trim()?.Length > 0 ? record["param0"]?.ToString()?.Trim() : sServer1;

                    //look for date and time
                    date = record["date"]?.ToString()?.Trim()?.Split(' ')[0];
                    time = timeConvertor.GetStringHHMMSS( record["time"]?.ToString()?.Trim());

                    //look for  idCard
                    idCard = 0;
                    Int32.TryParse(record["param1"]?.ToString()?.Trim(), out idCard);
                    fac = record["fac"]?.ToString()?.Trim();
                    card = record["card"]?.ToString()?.Trim();
                    idCardDescr = idCard != 0 ? "№" + idCard + " (" + fac + "," + card + ")" : (fio == sServer1 ? "" : "Пропуск не зарегистрирован");

                    //look for action with idCard
                    action = record["action"]?.ToString()?.Trim();
                    action_descr = null;
                    if (CARD_REGISTERED_ACTION.TryGetValue(action, out action_descr))
                    { action = "Сервисное сообщение"; }
                    else if (fio == sServer1)
                    { action_descr = action; }

                    //write gathered data in the collection
                    StatusTrace(fio + " " + action_descr + " " + idCard + " " + idCardDescr + " " + record["action"]?.ToString()?.Trim() + " " + date + " " + time + " " + sideOfPassagePoint._namePoint + " " + sideOfPassagePoint._direction);

                    visitors.Add(new Visitor(fio, idCardDescr, date, time, action_descr, sideOfPassagePoint));

                    if (startTimeNotSet) //set starttime into last time at once
                    {
                        startLoadingTimeInpsOutps = time;
                        startTimeNotSet = false;
                    }
                }
                StatusTrace("visitors.Count: " + visitors.Count());
            }

            return visitors;
        }
    }
}
