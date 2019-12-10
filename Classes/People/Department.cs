namespace ASTA.Classes.People
{
    internal interface IDepartment
    {
        string DepartmentId { get; set; }
        string DepartmentDescr { get; set; }
        string DepartmentBossCode { get; set; }
    }

    public class Department : IDepartment
    {
        public string DepartmentId { get; set; }
        public string DepartmentDescr { get; set; }
        public string DepartmentBossCode { get; set; }

        public override string ToString()
        {
            return $"{DepartmentId}\t{DepartmentDescr}\t{DepartmentBossCode}";
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            var df = (Department)obj;
            if (df == null)
                return false;

            return ToString().Equals(df.ToString());
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    public class DepartmentFull : IDepartment
    {
        public string DepartmentId { get; set; }
        public string DepartmentDescr { get; set; }
        public string DepartmentBossCode { get; set; }
        public string DepartmentBossEmail { get; set; }

        public override string ToString()
        {
            return $"{DepartmentId}\t{DepartmentDescr}\t{DepartmentBossCode}\t{DepartmentBossEmail}";
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            DepartmentFull df = obj as DepartmentFull;
            if (df == null)
                return false;

            return ToString().Equals(df.ToString());
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }
}