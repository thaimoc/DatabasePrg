using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using Microsoft.ApplicationBlocks.Data;
using Core;


namespace DataAccess
{
    public class Customer
    {
        #region Fields
        private string _CustomerID;

        public string CustomerID
        {
            get { return _CustomerID; }
            set { _CustomerID = value; }
        }
        private string _CompanyName;

        public string CompanyName
        {
            get { return _CompanyName; }
            set { _CompanyName = value; }
        }
        private string _ContactName;

        public string ContactName
        {
            get { return _ContactName; }
            set { _ContactName = value; }
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
        private string _Phone;

        public string Phone
        {
            get { return _Phone; }
            set { _Phone = value; }
        }
        private string _Fax;

        public string Fax
        {
            get { return _Fax; }
            set { _Fax = value; }
        }

        #endregion Fields

        #region Basic Methods 

        /// <summary>
        /// Tìm theo khóa chính
        /// </summary>
        /// <param name="customerId">CustomerID</param>
        /// <returns>Customer</returns>
        public static Customer Single(string customerId)
        {
            return CBO.FillObject<Customer>(DataProvider.Instance.ExecuteReader("Customers_Single", customerId));
        }
        /// <summary>
        /// Lấy tất cả danh sách khách hàng
        /// </summary>
        /// <returns>List<Customer></returns>
        public static List<Customer> All(int sort)
        {
            return CBO.FillCollection<Customer>(DataProvider.Instance.ExecuteReader("Customers_All", sort));
        }

        /// <summary>
        /// Tìm tất cả các countries
        /// </summary>
        /// <returns>lis of Customers</returns>
        public static List<Customer> Countries()
        {
            return CBO.FillCollection<Customer>(DataProvider.Instance.ExecuteReader("Customers_Countries"));
        }
        /// <summary>
        /// Tìm nhanh theo Company Name
        /// </summary>
        /// <param name="companyName">Company Name</param>
        /// <returns>list of Customers</returns>
        public static List<Customer> FindByCompanyName(string companyName)
        {
            return CBO.FillCollection<Customer>(DataProvider.Instance.ExecuteReader("Customers_FindByCompanyName", companyName));
        }
        /// <summary>
        /// Tìm theo country
        /// </summary>
        /// <param name="country"></param>
        /// <returns></returns>
        public static List<Customer> FindByCountry(string country)
        {
            return CBO.FillCollection<Customer>(DataProvider.Instance.ExecuteReader("Customers_FindByCountry", country));
        }
        /// <summary>
        /// Count of Customers
        /// </summary>
        /// <returns>Count</returns>
        public static int Count()
        {
            object rs = DataProvider.Instance.ExecuteScalar("Customers_Count");
            return MyConvert.ToInt32(rs);
        }

        /// <summary>
        /// Add thêm một customer vào database
        /// </summary>
        /// <param name="cus">The customer to add</param>
        /// <returns>CustomerID hoặc Empty String</returns>
        public static string Add(Customer cus)
        {
            object rs = DataProvider.Instance.ExecuteNonQuery("Customers_Add", cus._CustomerID, cus._CompanyName, cus._ContactName, cus._Address, cus._PostalCode, cus._Country, cus._Phone, cus._Fax);
            int affected = MyConvert.ToInt32(rs);
            if (affected > 0)
                return cus._CustomerID;
            return "";
        }
        /// <summary>
        /// Cập nhật thông tin khách hàng theo CustomerID
        /// </summary>
        /// <param name="cus">New Customer</param>
        /// <returns>CustomerID</returns>
        public static bool Update(Customer cus)
        {
            int rs = DataProvider.Instance.ExecuteNonQuery("Customers_Update", cus._CustomerID, 
                cus._CompanyName, cus._ContactName, cus._Address, cus._PostalCode, cus._Country, cus._Phone, cus._Fax);

            return rs > 0;
        }

        /// <summary>
        /// Xóa một khách hàng theo dứ liệu trong csdl
        /// </summary>
        /// <param name="cus"></param>
        /// <returns></returns>
        public static bool Delete(string id)
        {
            try
            { 
                //Tiến hành xóa, và trả về số dòng bị ảnh hưởng
                int rs = DataProvider.Instance.ExecuteNonQuery("Customers_Delete", id);
                //nếu rs > 0 = true, xóa thành công; ngược lại, không thành công
                return rs > 0;
            }
            catch
            {
                //Bát lỗi cơ sở dữ liệu
                return false;
            }
        }
        #endregion

        #region Methods Advance

        public static string[] ColumnNames()
        {
            DataSet ds = DataProvider.Instance.ExecuteDataset("Customers_All");
            DataTable dt = ds.Tables[0];
            string[] s = new string[dt.Columns.Count];
            int i = 0;
            foreach (DataColumn col in dt.Columns)
            {
                s[i++] = col.ColumnName;
            }
            return s;
        }

        public static List<Customer> Find(string expression)
        {
            expression = "SELECT * FROM Customers WHERE " + expression;
            return CBO.FillCollection<Customer>(DataProvider.Instance.ExecuteReader(expression));
        }

        public static List<Customer> Sort(string expression)
        {
            expression = "SELECT * FROM Customers ORDER BY " + expression;
            return CBO.FillCollection<Customer>(DataProvider.Instance.ExecuteReader(expression));
        }

        #endregion
    }
}
