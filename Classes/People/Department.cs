namespace ASTA.Classes.People
{
    public interface IDepartment
    {
        string DepartmentId { get; set; }
        string DepartmentDescr { get; set; }
        string DepartmentBossCode { get; set; }
    }

    public class Department : IDepartment
    {
        public Department() { }

        public Department(IDepartment department)
        {
            DepartmentId = department.DepartmentId;
            DepartmentDescr = department.DepartmentDescr;
            DepartmentBossCode = department.DepartmentBossCode;
        }

        public IDepartment Get() { return this; }
        public string DepartmentId { get; set; }
        public string DepartmentDescr { get; set; }
        public string DepartmentBossCode { get; set; }

        public override string ToString()
        {            return $"{DepartmentId}\t{DepartmentDescr}\t{DepartmentBossCode}";        }

        public override bool Equals(object obj)
        {
            var df = (Department)obj;
            if (obj == null || df == null)
                return false;

            return ToString().Equals(df.ToString());
        }

        public override int GetHashCode()
        {            return ToString().GetHashCode();        }
    }

    public class DepartmentFull : Department
    {
        public DepartmentFull() : base() { }
        
        public DepartmentFull(IDepartment department) : base() {
            DepartmentId = department.DepartmentId;
            DepartmentDescr = department.DepartmentDescr;
            DepartmentBossCode = department.DepartmentBossCode;
        }

        public string DepartmentBossEmail { get; set; }

        public new DepartmentFull Get() { return this; }

        public override string ToString()
        {            return $"{DepartmentId}\t{DepartmentDescr}\t{DepartmentBossCode}\t{DepartmentBossEmail}";        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            if (obj as DepartmentFull == null)
                return false;

            return ToString().Equals((obj as DepartmentFull).ToString());
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }
}