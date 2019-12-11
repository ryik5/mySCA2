namespace ASTA.Classes.People
{
    public class EmployeeFull : Employee
    {
        public int idCard { get; set; }//= 0
        public string PositionInDepartment { get; set; }//= ""
        public string GroupPerson { get; set; }//= ""
        public string City { get; set; }//= ""
        public string Department { get; set; }//= ""
        public string DepartmentId { get; set; }//= ""
        public string DepartmentBossCode { get; set; }//= ""
        public int ControlInSeconds { get; set; }//= 32460
        public int ControlOutSeconds { get; set; }// =64800
        public string ControlInHHMM { get; set; }//= "09:00"
        public string ControlOutHHMM { get; set; }//= "18:00"
        public string Shift { get; set; }//= ""   /* персональный график*/
        public string Comment { get; set; }//= ""

        public new EmployeeFull Get() { return this; }
        public override string ToString()
        {
            return $"{idCard}\t{fio}\t{Code}\t{Department}\t{DepartmentId}\t{DepartmentBossCode}\t" +
                $"{PositionInDepartment}\t{GroupPerson}\t{City}\t" +
                $"{ControlInSeconds}\t{ControlOutSeconds}\t{ControlInHHMM}\t{ControlOutHHMM}\t{Shift}\t{Comment}";
        }

        public override bool Equals(object obj)
        {
            if (obj == null || !(obj is EmployeeFull df))
                return false;

            return ToString() == df.ToString();
        }

        public override int GetHashCode()
        { return ToString().GetHashCode(); }
    }
}