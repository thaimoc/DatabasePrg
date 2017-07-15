using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using Core;

namespace DataAccess
{
    public class Supplier
    {
        #region Fields
        private int _SupplierID;

        public int SupplierID
        {
            get { return _SupplierID; }
            set { _SupplierID = value; }
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
        private string _ContactTitle;

        public string ContactTitle
        {
            get { return _ContactTitle; }
            set { _ContactTitle = value; }
        }
        private string _Address;

        public string Address
        {
            get { return _Address; }
            set { _Address = value; }
        }
        private string _City;

        public string City
        {
            get { return _City; }
            set { _City = value; }
        }
        private string _Region;

        public string Region
        {
            get { return _Region; }
            set { _Region = value; }
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
        private string _HomePage;

        public string HomePage
        {
            get { return _HomePage; }
            set { _HomePage = value; }
        }
        #endregion

        #region Basic Methods
        /// <summary>
        /// Tìm tất cả nhà cung cấp
        /// </summary>
        /// <returns>List</returns>
        public static List<Supplier> All()
        {
            return CBO.FillCollection<Supplier>(DataProvider.Instance.ExecuteReader("Suppliers_All"));
        }
        /// <summary>
        /// Tìm tất cả các country cung cấp
        /// </summary>
        /// <returns></returns>
        public static List<Supplier> Countries()
        {
            return CBO.FillCollection<Supplier>(DataProvider.Instance.ExecuteReader("Suppliers_Countries"));
        }
        /// <summary>
        /// Tìm theo khóa chính
        /// </summary>
        /// <param name="supplierID">SupplierID</param>
        /// <returns>Supplier</returns>
        public static Supplier Single(int supplierID)
        {
            return CBO.FillObject<Supplier>(DataProvider.Instance.ExecuteReader("Suppliers_Single", supplierID));
        }
        public static int Count()
        {
            object rs = DataProvider.Instance.ExecuteScalar("Suppliers_Count");
            return MyConvert.ToInt32(rs);
        }
        /// <summary>
        /// Like company's Name
        /// </summary>
        /// <param name="companyName">company's Name</param>
        /// <returns>list</returns>
        public static List<Supplier> FindByLikeCompany(string companyName)
        {
            return CBO.FillCollection<Supplier>(DataProvider.Instance.ExecuteReader("Suppliers_FindByLikeCompany", companyName));
        }
        /// <summary>
        /// Find By Company
        /// </summary>
        /// <param name="companyName"></param>
        /// <returns></returns>
        public static List<Supplier> FindByCompany(string companyName)
        {
            return CBO.FillCollection<Supplier>(DataProvider.Instance.ExecuteReader("Suppliers_FindByCompany", companyName));
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="companyName"></param>
        /// <returns></returns>
        public static List<Supplier> FindByCountry(string contry)
        {
            return CBO.FillCollection<Supplier>(DataProvider.Instance.ExecuteReader("Suppliers_FindByContry", contry));
        }
        /// <summary>
        /// Thêm mới một nhà cung cấp
        /// </summary>
        /// <param name="supplier">New Supplier</param>
        /// <returns>SupplierID</returns>
        public static int Add(Supplier supplier)
        {
            object rs = DataProvider.Instance.ExecuteNonQueryWithOutput("@SupplierID", "Suppliers_Add", 0, supplier._CompanyName,
                supplier._ContactName, supplier._ContactTitle, supplier._Address, supplier._City, supplier._Region, supplier._PostalCode,
                supplier._Country, supplier._Phone, supplier._Fax, supplier._HomePage);
            return MyConvert.ToInt32(rs);//Trả về SupplierID tự tăng mới nhất
        }
        /// <summary>
        /// Cập nhật thông tin một nhà cung cấp
        /// </summary>
        /// <param name="supplier">New Supplier</param>
        /// <returns>Số dòng được update > 0 return true, ngược lại return false</returns>
        public static bool Update(Supplier supplier)
        {
            object rs = DataProvider.Instance.ExecuteNonQuery("Suppliers_Update", supplier._SupplierID, supplier._CompanyName,
                supplier._ContactName, supplier._ContactTitle, supplier._Address, supplier._City, supplier._Region, supplier._PostalCode,
                supplier._Country, supplier._Phone, supplier._Fax, supplier._HomePage);
            return MyConvert.ToInt32(rs) > 0;
        }
        /// <summary>
        /// Xóa thông tin của một nhà cung cấp theo khóa chính
        /// </summary>
        /// <param name="supplier">Supplier</param>
        /// <returns>Số dòng bị xóa</returns>
        public static bool Delete(int supplierID)
        {
            object rs = DataProvider.Instance.ExecuteNonQuery("Suppliers_Delete", supplierID);
            return MyConvert.ToInt32(rs) > 0;
        }


        

        #endregion

        public static string[] ColumnNames()
        {
            DataSet ds = DataProvider.Instance.ExecuteDataset("Suppliers_All");
            DataTable dt = ds.Tables[0];
            string[] s = new string[dt.Columns.Count];
            int i = 0;
            foreach (DataColumn col in dt.Columns)
            {
                s[i++] = col.ColumnName;
            }
            return s;
        }

        public static List<Supplier> Find(string expression)
        {
            expression = "SELECT * FROM Suppliers WHERE " + expression;
            return CBO.FillCollection<Supplier>(DataProvider.Instance.ExecuteReader(expression));
        }

        public static List<Supplier> Sort(string expression)
        {
            expression = "SELECT * FROM Suppliers ORDER BY " + expression;
            return CBO.FillCollection<Supplier>(DataProvider.Instance.ExecuteReader(expression));
        }

    }
}
