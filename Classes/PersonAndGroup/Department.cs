using System;

namespace ASTA
{    
    interface IDepartment
    {
        string _departmentId { get; set; }
        string _departmentDescription { get; set; }
        string _departmentBossCode { get; set; }
    }

    class Department : IDepartment
    {
        public string _departmentId { get; set; }
        public string _departmentDescription { get; set; }
        public string _departmentBossCode { get; set; }

        public override string ToString()
        {
            return _departmentId + "\t" + _departmentDescription + "\t" + _departmentBossCode;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            Department df = obj as Department;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }
        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    class DepartmentFull : IDepartment
    {
        public string _departmentId { get; set; }
        public string _departmentDescription { get; set; }
        public string _departmentBossCode { get; set; }
        public string _departmentBossEmail { get; set; }

        public override string ToString()
        {
            return _departmentId + "\t" + _departmentDescription + "\t" + _departmentBossCode + "\t" + _departmentBossEmail;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            DepartmentFull df = obj as DepartmentFull;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }
        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }
    
}
