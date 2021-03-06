﻿using System.Collections.Generic;

namespace ASTA.Classes
{
    public static class Names
    {
        public readonly static string[] allParametersOfConfig = new string[] {
                @"SKDServer" , @"SKDUser", @"SKDUserPassword",
                @"MySQLServer", @"MySQLUser", @"MySQLUserPassword",
                @"MailServer", @"MailUser", @"MailUserPassword",
                @"DEFAULT_DAY_OF_SENDING_REPORT", @"MailServerSMTPport",
                @"DomainOfUser", @"DomainController", @"UserName", @"UserPassword",
                @"clrRealRegistration", @"ShiftDaysBackOfSendingFromLastWorkDay",
                @"JobReportsReceiver", @"RemoteFolderUpdateURL"
            };

        public const string CARD_STATE = @"Событие точки доступа";

        public readonly static Dictionary<string, string> CARD_REGISTERED_ACTION = new Dictionary<string, string>()
        {
            {"ACCESS_IN", "Успешный проход" },
            {"NOACCESS_CARD",  "Неизвестная карта" },
            {"NOACCESS_LEVEL", "Доступ не разрешен" },
            {"RQ_NOACC_CARD" , "Запрос"}
        };

        //Names of collumns
        public const string NPP = @"№ п/п";

        public const string FIO = @"Фамилия Имя Отчество";
        public const string CODE = @"NAV-код";
        public const string GROUP = @"Группа";
        public const string GROUP_DECRIPTION = @"Описание группы";
        public const string PLACE_EMPLOYEE = @"Местонахождение сотрудника";
        public const string DEPARTMENT = @"Отдел";
        public const string DEPARTMENT_ID = @"Отдел (id)";
        public const string CHIEF_ID = @"Руководитель (код)";
        public const string EMPLOYEE_POSITION = @"Должность";

        public const string DESIRED_TIME_IN = @"Учетное время прихода ЧЧ:ММ";
        public const string DESIRED_TIME_OUT = @"Учетное время ухода ЧЧ:ММ";

        public const string N_ID = @"№ пропуска";
        public const string N_ID_STRING = @"Пропуск";
        public const string DATE_REGISTRATION = @"Дата регистрации";
        public const string DAY_OF_WEEK = @"День недели";
        public const string TIME_REGISTRATION = @"Время регистрации";
        public const string TIME_REGISTRATION_STRING = @"Время  регистрации";

        public const string SERVER_SKD = @"Сервер СКД";
        public const string CHECKPOINT_NAME = @"Имя точки прохода";
        public const string CHECKPOINT_ACTION = @"Результат";
        public const string CHECKPOINT_DIRECTION = @"Направление прохода";

        public const string REAL_TIME_IN = @"Фактич. время прихода ЧЧ:ММ:СС";
        public const string REAL_TIME_OUT = @"Фактич. время ухода ЧЧ:ММ:СС";

        public const string EMPLOYEE_SHIFT = @"График";
        public const string EMPLOYEE_SHIFT_COMMENT = @"Комментарии (командировка, на выезде, согласованное отсутствие…….)";
        public const string EMPLOYEE_HOOKY = @"Прогул (отпуск за свой счет)";
        public const string EMPLOYEE_ABSENCE = @"Отсутствовал на работе";
        public const string EMPLOYEE_SICK_LEAVE = @"Больничный";
        public const string EMPLOYEE_TRIP = @"Командировка";
        public const string EMPLOYEE_VACATION = @"Отпуск";
        public const string EMPLOYEE_BEING_LATE = @"Опоздание ЧЧ:ММ";
        public const string EMPLOYEE_EARLY_DEPARTURE = @"Ранний уход ЧЧ:ММ";
        public const string EMPLOYEE_PLAN_TIME_WORKED = @"Отработанное время ЧЧ:ММ";
        public const string EMPLOYEE_TIME_SPENT = @"Реальное отработанное время";

        //Mailing
        public const string RECEPIENTS_OF_REPORTS = @"Получатель рассылки";

        //Page of Report
        //Collumns of the table 'Day off'
        public const string DAYOFF_DATE = @"День";

        public const string DAYOFF_TYPE = @"Выходной или рабочий день";
        public const string DAYOFF_USED_BY = @"Персонально(NAV) или для всех(0)";
        public const string DAYOFF_ADDED = @"Дата добавления";

        //string constants
        public const string DAY_OFF_OR_WORK = @"Выходные(рабочие) дни";

        public const string DAY_OFF_OR_WORK_EDIT = @"Режим добавления/удаления праздничных дней";

        public const string WORK_WITH_A_PERSON = @"Работать с одной персоной";
        public const string WORK_WITH_A_GROUP = @"Работать с группой";

        public readonly static string[] arrayAllColumnsDataTablePeople =
            {
                 NPP,//0
                 FIO,//1
                 CODE,//2
                 GROUP,//3
                 N_ID, //6
                 N_ID_STRING,
                 DEPARTMENT,//7
                 PLACE_EMPLOYEE,//8
                 DATE_REGISTRATION,//9
                 TIME_REGISTRATION, //10
                 TIME_REGISTRATION_STRING,
                 REAL_TIME_IN,                      //17
                 REAL_TIME_OUT,                     //18
                 SERVER_SKD,                        //11
                 CHECKPOINT_NAME,                   //12
                 CHECKPOINT_DIRECTION,                     //13
                 CHECKPOINT_ACTION,                     //13
                 CARD_STATE,                //14
                 DESIRED_TIME_IN,                   //15
                 DESIRED_TIME_OUT,                  //16
                 EMPLOYEE_TIME_SPENT,               //19
                 EMPLOYEE_PLAN_TIME_WORKED,         //20
                 EMPLOYEE_BEING_LATE,               //21
                 EMPLOYEE_EARLY_DEPARTURE,          //22
                 EMPLOYEE_VACATION,                 //23
                 EMPLOYEE_TRIP,                     //24
                 DAY_OF_WEEK,                       //25
                 EMPLOYEE_SICK_LEAVE,               //26
                 EMPLOYEE_ABSENCE,                  //27
                 GROUP_DECRIPTION,                  //28
                 EMPLOYEE_SHIFT_COMMENT,            //29
                 EMPLOYEE_POSITION,                 //30
                 EMPLOYEE_SHIFT,                    //31
                 EMPLOYEE_HOOKY,                    //32
                 DEPARTMENT_ID,                     //33
                 CHIEF_ID                           //34
        };

        public readonly static string[] orderColumnsFinacialReport =
             {
                 FIO,//1
                 CODE,//2
                 DEPARTMENT,//11
                 PLACE_EMPLOYEE,//12
                 DATE_REGISTRATION,//12
                 DAY_OF_WEEK,                    //32
                 DESIRED_TIME_IN,//22
                 DESIRED_TIME_OUT,//23
                 REAL_TIME_IN,//24
                 REAL_TIME_OUT, //25
                 EMPLOYEE_PLAN_TIME_WORKED, //27
                 EMPLOYEE_BEING_LATE,                    //28
                 EMPLOYEE_EARLY_DEPARTURE,                 //29
                 EMPLOYEE_VACATION,                 //30
                 EMPLOYEE_HOOKY,    //41
                 EMPLOYEE_SICK_LEAVE,                    //33
                 EMPLOYEE_ABSENCE,      //34
                 EMPLOYEE_SHIFT_COMMENT,                    //38
                 EMPLOYEE_POSITION,                    //39
                 EMPLOYEE_SHIFT                    //40
        };

        public readonly static string[] orderColumnsLastRegistrations =
              {
                 FIO,//1
                 N_ID_STRING,
                 DATE_REGISTRATION,//12
                 TIME_REGISTRATION_STRING,
                 CHECKPOINT_DIRECTION,
                 CHECKPOINT_ACTION
        };

        public readonly static string[] orderColumnsRegistrations =
              {
                 FIO,//1
                 CODE,//2
                 GROUP,//3
                 N_ID, //6
                 DEPARTMENT,//7
                 PLACE_EMPLOYEE,//8
                 DATE_REGISTRATION,//9
                 REAL_TIME_IN,                      //17
                 REAL_TIME_OUT,                     //18
                 DESIRED_TIME_IN,                   //15
                 DESIRED_TIME_OUT,                  //16
                 EMPLOYEE_SHIFT_COMMENT,            //29
                 EMPLOYEE_POSITION,                 //30
                 EMPLOYEE_SHIFT,                    //31
                 DEPARTMENT_ID,                     //33
                 CHIEF_ID                           //34
        };

        public readonly static string[] orderColumnsFIO =
            {
                 FIO,//1
                 CODE,//2
                 GROUP,//3
                 DEPARTMENT,//7
                 PLACE_EMPLOYEE,//8
                 DESIRED_TIME_IN,                   //15
                 DESIRED_TIME_OUT,                  //16
                 EMPLOYEE_POSITION,                 //30
                 EMPLOYEE_SHIFT,                    //31
                 DEPARTMENT_ID,                     //33
                 CHIEF_ID                           //34
        };

        public const string END_EDIT = @"Завершить редактирование";
    }
}