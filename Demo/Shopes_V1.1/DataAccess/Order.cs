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
    public class Order
    {
        #region Fields
        private int _OrderID;

        public int OrderID
        {
            get { return _OrderID; }
            set { _OrderID = value; }
        }
        private string _CustomerID;

        public string CustomerID
        {
            get { return _CustomerID; }
            set { _CustomerID = value; }
        }
        private int _EmployeeID;

        public int EmployeeID
        {
            get { return _EmployeeID; }
            set { _EmployeeID = value; }
        }
        private DateTime _OrderDate;

        public DateTime OrderDate
        {
            get { return _OrderDate; }
            set { _OrderDate = value; }
        }
        private DateTime _RequiredDate;

        public DateTime RequiredDate
        {
            get { return _RequiredDate; }
            set { _RequiredDate = value; }
        }
        private DateTime _ShippedDate;

        public DateTime ShippedDate
        {
            get { return _ShippedDate; }
            set { _ShippedDate = value; }
        }
        private double _Freight;

        public double Freight
        {
            get { return _Freight; }
            set { _Freight = value; }
        }
        private string _ShipAddress;

        public string ShipAddress
        {
            get { return _ShipAddress; }
            set { _ShipAddress = value; }
        }
        private string _ShipPostalCode;

        public string ShipPostalCode
        {
            get { return _ShipPostalCode; }
            set { _ShipPostalCode = value; }
        }
        private string _Status;

        public string Status
        {
            get { return _Status; }
            set { _Status = value; }
        }

        private string _Customer;

        public string Customer
        {
            get { return _Customer; }
            set { _Customer = value; }
        }

        private string _Employee;

        public string Employee
        {
            get { return _Employee; }
            set { _Employee = value; }
        }

        #endregion

        

        #region Basic Methods
        /// <summary>
        /// Tìm order theo khóa chính
        /// </summary>
        /// <param name="orderID">OrderID</param>
        /// <returns>Order</returns>
        public static Order Single(int orderID)
        {
           return CBO.FillObject<Order>(DataProvider.Instance.ExecuteReader("Orders_Single", orderID));
        }
        
        /// <summary>
        /// Lấy tất cả các hóa đơn trong csdl
        /// </summary>
        /// <returns>List</returns>
        public static List<Order> All()
        {
            return CBO.FillCollection<Order>(DataProvider.Instance.ExecuteReader("Orders_All"));
        }
        /// <summary>
        /// Find By EmployeeID
        /// </summary>
        /// <param name="employeeID"></param>
        /// <returns></returns>
        public static List<Order> FindByEmployeeID(int employeeID)
        {
            return CBO.FillCollection<Order>(DataProvider.Instance.ExecuteReader("Orders_FindByEmployeeID", employeeID));
        }
        /// <summary>
        /// Find By Customer
        /// </summary>
        /// <param name="customer"></param>
        /// <returns></returns>
        public static List<Order> FindByCustomer(string customer)
        {
            return CBO.FillCollection<Order>(DataProvider.Instance.ExecuteReader("Orders_FindByCustomer", customer));
        }
        /// <summary>
        /// All of status
        /// </summary>
        /// <returns>Status column</returns>
        public static List<Order> Statuses()
        {
            return CBO.FillCollection<Order>(DataProvider.Instance.ExecuteReader("Orders_Status"));
        }
        /// <summary>
        /// Find By Status
        /// </summary>
        /// <param name="status"></param>
        /// <returns></returns>
        public static List<Order> FindByStatus(string status)
        {
            return CBO.FillCollection<Order>(DataProvider.Instance.ExecuteReader("Orders_FindByStatus", status));
        }
        /// <summary>
        /// Count of orders
        /// </summary>
        /// <returns>Count</returns>
        public static int Count()
        {
            object rs = DataProvider.Instance.ExecuteScalar("Orders_Count");
            return MyConvert.ToInt32(rs);
        }
       
        /// <summary>
        /// Thêm mới một hóa đơn
        /// </summary>
        /// <param name="order">New Order</param>
        /// <returns>OrderID</returns>
        public static int Add(Order order)
        {
            object numberAffected = DataProvider.Instance.ExecuteNonQueryWithOutput("@OrderID", "Orders_Add", 0, order._CustomerID, order._EmployeeID, order._OrderDate, order._RequiredDate,
               order.ShippedDate, order._Freight, order._ShipAddress, order._ShipPostalCode, order._Status);
            int affected = MyConvert.ToInt32(numberAffected);
            if (affected > 0)
                return affected;
            else
                return 0;
        }
        /// <summary>
        /// Cập nhật thông tin hóa đơn theo khóa chính OrderID
        /// </summary>
        /// <param name="order"></param>
        /// <returns></returns>
        public static bool Update(Order order)
        {
            object numberAffected = DataProvider.Instance.ExecuteNonQueryWithOutput("@OrderID", "Orders_Update", 0, order._CustomerID, order._EmployeeID, order._OrderDate, order._RequiredDate,
               order.ShippedDate, order._Freight, order._ShipAddress, order._ShipPostalCode, order._Status);
            int affected = MyConvert.ToInt32(numberAffected);
            if (affected > 0)
                return true;
            else
                return false;
        }

        public static bool Delete(int orderID)
        {
            int rs = DataProvider.Instance.ExecuteNonQuery("Orders_Delete", orderID);
            return rs > 0;
        }

        #endregion

        public static string[] ColumnNames()
        {
            DataSet ds = DataProvider.Instance.ExecuteDataset("Orders_All");
            DataTable dt = ds.Tables[0];
            string[] s = new string[dt.Columns.Count];
            int i = 0;
            foreach (DataColumn col in dt.Columns)
            {
                s[i++] = col.ColumnName;
            }
            return s;
        }

        public static List<Order> Find(string expression)
        {
            expression = @"SELECT * FROM View_Orders WHERE " + expression;
            return CBO.FillCollection<Order>(DataProvider.Instance.ExecuteReader(expression));
        }

        public static List<Order> Sort(string expression)
        {
            expression = "SELECT * FROM View_Orders ORDER BY " + expression;
            return CBO.FillCollection<Order>(DataProvider.Instance.ExecuteReader(expression));
        }
        public static List<Order> FindBy(object value, object cbbvalue)
        {
            string query = "Select * from Orders where " + cbbvalue + " like '" + value + "%'";
            return CBO.FillCollection<Order>(DataProvider.Instance.ExecuteReader(query));
        }

    }
}
