using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using Core;
using Microsoft.ApplicationBlocks.Data;

namespace DataAccess
{
    public class Employee
    {
        #region Fields
        private int _EmployeeID;

        public int EmployeeID
        {
            get { return _EmployeeID; }
            set { _EmployeeID = value; }
        }
        private string _LastName;

        public string LastName
        {
            get { return _LastName; }
            set { _LastName = value; }
        }
        private string _FirstName;

        public string FirstName
        {
            get { return _FirstName; }
            set { _FirstName = value; }
        }
        private DateTime _BirthDate;

        public DateTime BirthDate
        {
            get { return _BirthDate; }
            set { _BirthDate = value; }
        }
        private DateTime _HireDate;

        public DateTime HireDate
        {
            get { return _HireDate; }
            set { _HireDate = value; }
        }
        private string _Address;

        public string Address
        {
            get { return _Address; }
            set { _Address = value; }
        }
        private string _PostalCode;

        public string PostalCode
        {
            get { return _PostalCode; }
            set { _PostalCode = value; }
        }
        private string _Country;

        public string Country
        {
            get { return _Country; }
            set { _Country = value; }
        }
        private string _HomePhone;

        public string HomePhone
        {
            get { return _HomePhone; }
            set { _HomePhone = value; }
        }

        private string _EmployeeName;

        public string EmployeeName
        {
            get { return _EmployeeName; }
            set { _EmployeeName = value; }
        }

        
        #endregion

        #region Basic Methods
        /// <summary>
        /// Employee tìm theo khóa chính
        /// </summary>
        /// <param name="employeeID">(int) EmployeeID</param>
        /// <returns>Employee</returns>
        public static Employee Single(int employeeID)
        {
            return CBO.FillObject<Employee>(DataProvider.Instance.ExecuteReader("Employees_Single", employeeID));
        }

        /// <summary>
        /// The count of employees
        /// </summary>
        /// <returns>count</returns>
        public static int Count()
        {
            object rs = DataProvider.Instance.ExecuteScalar("Employees_Count");
            return MyConvert.ToInt32(rs);
        }

        /// <summary>
        /// Lấy danh sách nhân viên từ csdl
        /// </summary>
        /// <returns>[List<Employee>]</returns>
        public static List<Employee> All()
        {
            return CBO.FillCollection<Employee>(DataProvider.Instance.ExecuteReader("Employees_All"));
        }
        /// <summary>
        /// Find By Birth Date
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static List<Employee> FindByBirthDate(DateTime date)
        {
            return CBO.FillCollection<Employee>(DataProvider.Instance.ExecuteReader("Employees_FindByBirthDate", date));
        }
        /// <summary>
        /// Find By Month of Birth Date
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static List<Employee> FindByMonthBirthDate(DateTime date)
        {
            return CBO.FillCollection<Employee>(DataProvider.Instance.ExecuteReader("Employees_FindByMonthBirthDate", date));
        }
        /// <summary>
        /// Find By Name
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static List<Employee> FindByName(string name)
        {
            return CBO.FillCollection<Employee>(DataProvider.Instance.ExecuteReader("Employees_FindByName", name));
        }
        /// <summary>
        /// Thêm một nhân viên vào csdl
        /// </summary>
        /// <param name="empl">New employee</param>
        /// <returns>(int)EmployeeID</returns>
        public static int Add(Employee empl)
        {
            object rs = DataProvider.Instance.ExecuteNonQueryWithOutput("@EmployeeID", "Employees_Add", 0, empl.LastName, empl.FirstName, empl.BirthDate, empl.HireDate, empl.Address, empl.PostalCode, empl.Country, empl.HomePhone);
            return MyConvert.ToInt32(rs);
        }
        /// <summary>
        /// Cập nhật thông tin một nhân viên
        /// </summary>
        /// <param name="empl"></param>
        /// <returns></returns>
        public static int Update(Employee empl)
        {
            int rs = DataProvider.Instance.ExecuteNonQuery("Employees_Update", empl.EmployeeID, empl.LastName, empl.FirstName, empl.BirthDate, empl.HireDate, empl.Address, empl.PostalCode, empl.Country, empl.HomePhone);
            return rs > 0 ? rs : 0;
        }

        /// <summary>
        /// Xoá một nhân viên khỏi csdl theo khóa chính
        /// </summary>
        /// <param name="empl">Employee</param>
        /// <returns>EmployeeID</returns>
        public static bool Delete(Int32 id)
        {
            int rs = DataProvider.Instance.ExecuteNonQuery("Employees_Delete", id);
            return rs > 0 ? true : false;
        }

       
        #endregion

        #region Methods Advance
        public static string[] ColumnNames()
        {
            DataSet ds = DataProvider.Instance.ExecuteDataset("Employees_AllStandar");
            DataTable dt = ds.Tables[0];
            string[] s = new string[dt.Columns.Count];
            int i = 0;
            foreach (DataColumn col in dt.Columns)
            {
                s[i++] = col.ColumnName;
            }
            return s;
        }

        public static List<Employee> Find(string expression)
        {
            expression = "SELECT * FROM Employees WHERE " + expression;
            return CBO.FillCollection<Employee>(DataProvider.Instance.ExecuteReader(expression));
        }

        public static List<Employee> Sort(string expression)
        {
            expression = "SELECT * FROM Employees ORDER BY " + expression;
            return CBO.FillCollection<Employee>(DataProvider.Instance.ExecuteReader(expression));
        }



        #endregion
    }
}
