namespace ASTA.Classes.People
{
    public class EmployeeFull : Employee
    {
        public int idCard;//= 0
        public string PositionInDepartment;//= ""
        public string GroupPerson;//= ""
        public string City;//= ""
        public string Department;//= ""
        public string DepartmentId;//= ""
        public string DepartmentBossCode;//= ""
        public int ControlInSeconds;//= 32460
        public int ControlOutSeconds;// =64800
        public string ControlInHHMM;//= "09:00"
        public string ControlOutHHMM;//= "18:00"
        public string Shift;//= ""   /* персональный график*/
        public string Comment;//= ""

        public override string ToString()
        {
            return $"{idCard}\t{fio}\t{Code}\t{Department}\t{DepartmentId}\t{DepartmentBossCode}\t" +
                $"{PositionInDepartment}\t{GroupPerson}\t{City}\t" +
                $"{ControlInSeconds}\t{ControlOutSeconds}\t{ControlInHHMM}\t{ControlOutHHMM}\t{Shift}\t{Comment}";
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            if (!(obj is EmployeeFull df))
                return false;

            return ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }
}