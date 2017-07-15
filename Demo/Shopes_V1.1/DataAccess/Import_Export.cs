using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Data;

namespace DataAccess
{
    public static class Import_Export
    {

        private static string GetOleDBConnect(string FileName, bool hasHeaders)
        {
            string HDR = hasHeaders ? "Yes" : "No";
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                              FileName + ";Extended Properties=\"Excel 8.0;HDR=" +
                              HDR + ";IMEX=1\"";
            return strConn;
        }
        public static Employee GetDataEmployeeToExcel(OleDbDataReader reader)
        {
            Employee emp = new Employee();
            emp.EmployeeID = MyConvert.ToInt32(reader["EmployeeID"].ToString());
            emp.EmployeeName = MyConvert.ToString(reader["Name"].ToString());
            emp.HireDate = MyConvert.ToDateTime(reader["HireDate"].ToString());
            emp.BirthDate = MyConvert.ToDateTime(reader["BirthDate"].ToString());
            emp.Address = MyConvert.ToString(reader["Address"].ToString());
            emp.PostalCode = MyConvert.ToString(reader["PostalCode"].ToString());
            emp.Country = MyConvert.ToString(reader["Country"].ToString());
            emp.HomePhone = MyConvert.ToString(reader["HomePhone"].ToString());
            return emp;
        }
        public static Customer GetDataCustomerToExcel(OleDbDataReader reader)
        {
            Customer cus = new Customer();
            cus.CustomerID = MyConvert.ToString(reader["ID"]);
            cus.CompanyName = MyConvert.ToString(reader["CompanyName"]);
            cus.ContactName = MyConvert.ToString(reader["ContactName"]);
            cus.Address = MyConvert.ToString(reader["Address"]);
            cus.PostalCode = MyConvert.ToString(reader["PostalCode"]);
            cus.Country = MyConvert.ToString(reader["Country"]);
            cus.Phone = MyConvert.ToString(reader["Phone"]);
            cus.Fax = MyConvert.ToString(reader["Fax"]);
            return cus;
        }
        public static Order GetDataOrderToExcel(OleDbDataReader reader)
        {
            Order od = new Order();
            od.OrderID = MyConvert.ToInt32(reader["OrderID"]);
            od.CustomerID = MyConvert.ToString(reader["CustomerID"]);
            od.EmployeeID = MyConvert.ToInt32(reader["EmployeeID"]);
            od.OrderDate = MyConvert.ToDateTime(reader["OrderDate"]);
            od.RequiredDate = MyConvert.ToDateTime(reader["RequiredDate"]);
            od.ShippedDate = MyConvert.ToDateTime(reader["ShippedDate"]);
            od.Freight = MyConvert.ToDouble(reader["Freight"]);
            od.ShipAddress = MyConvert.ToString(reader["ShipAddress"]);
            od.ShipPostalCode = MyConvert.ToString(reader["ShipPostalCode"]);
            od.Status = MyConvert.ToString(reader["Status"]);
            return od;
        }
        public static Product GetDataProductToExcel(OleDbDataReader reader)
        {
            Product pro = new Product();
            pro.ProductID = MyConvert.ToInt32(reader["ProductID"]);
            pro.ProductName = MyConvert.ToString(reader["ProductName"]);
            pro.SupplierID = MyConvert.ToInt32(reader["SupplierID"]);
            pro.QuantityPerUnit = MyConvert.ToString(reader["ProductName"]);
            pro.UnitPrice = MyConvert.ToDouble(reader["UnitPrice"]);
            pro.UnitsInStock = MyConvert.ToInt16(reader["UnitsInStock"]);
            pro.UnitsOnOrder = MyConvert.ToInt16(reader["UnitsOnOrder"]);
            pro.Discontinued = MyConvert.ToBool(reader["Discontinued"]);
            return pro;
        }
        public static Supplier GetDataSupplierToExcel(OleDbDataReader reader)
        {
            Supplier sup = new Supplier();
            sup.SupplierID = MyConvert.ToInt32(reader["SupplierID"]);
            sup.CompanyName = MyConvert.ToString(reader["CompanyName"]);
            sup.ContactName = MyConvert.ToString(reader["ContactName"]);
            sup.ContactTitle = MyConvert.ToString(reader["ContactTitle"]);
            sup.Address = MyConvert.ToString(reader["ContactTitle"]);
            sup.City = MyConvert.ToString(reader["City"]);
            sup.Country = MyConvert.ToString(reader["Country"]);
            sup.Region = MyConvert.ToString(reader["Region"]);
            sup.PostalCode = MyConvert.ToString(reader["Region"]);
            sup.Phone = MyConvert.ToString(reader["Region"]);
            sup.Fax = MyConvert.ToString(reader["Region"]);
            sup.HomePage = MyConvert.ToString(reader["Region"]);
            return sup;
        }
        public static OrderDetail GetDataOrderDetailToExcel(OleDbDataReader reader)
        {
            OrderDetail odd = new OrderDetail();
            odd.OrderID = MyConvert.ToInt32(reader["OrderID"]);
            odd.ProductID = MyConvert.ToInt32(reader["ProductID"]);
            odd.UnitPrice = MyConvert.ToDouble(reader["UnitPrice"]);
            odd.Quantity = MyConvert.ToInt16(reader["Quantity"]);
            odd.Discount = MyConvert.ToInt32(reader["Discount"]);
            return odd;
        }

        public static Worksheet ExportExcel()
        {
            Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = xls.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet ws = (Worksheet)xls.ActiveSheet;
            xls.Visible = true;
            return ws;
        }

        public static List<Employee> ImportExcel_Employee(string FileName, bool hasHeaders)
        {
            string strConn = GetOleDBConnect(FileName, hasHeaders);
            using (OleDbConnection con = new OleDbConnection(strConn))
            {
                string query = "SELECT * FROM [Sheet1$]";
                OleDbCommand cmd = new OleDbCommand(query, con);
                con.Open();
                using (OleDbDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                {
                    List<Employee> list = new List<Employee>();
                    Employee emp = null;
                    while (reader.Read())
                    {
                        emp = GetDataEmployeeToExcel(reader);
                        list.Add(emp);
                    }
                    return list;
                }
            }
        }
        public static List<Customer> ImportExcel_Customer(string FileName, bool hasHeaders)
        {
            string strConn = GetOleDBConnect(FileName, hasHeaders);
            using (OleDbConnection con = new OleDbConnection(strConn))
            {
                string query = "SELECT * FROM [Sheet1$]";
                OleDbCommand cmd = new OleDbCommand(query, con);
                con.Open();
                using (OleDbDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                {
                    List<Customer> list = new List<Customer>();
                    Customer cus = null;
                    while (reader.Read())
                    {
                        cus = GetDataCustomerToExcel(reader);
                        list.Add(cus);
                    }
                    return list;
                }
            }
        }
        public static List<Order> ImportExcel_Order(string FileName, bool hasHeaders)
        {
            string strConn = GetOleDBConnect(FileName, hasHeaders);
            using (OleDbConnection con = new OleDbConnection(strConn))
            {
                string query = "SELECT * FROM [Sheet1$]";
                OleDbCommand cmd = new OleDbCommand(query, con);
                con.Open();
                using (OleDbDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                {
                    List<Order> list = new List<Order>();
                    Order od = null;
                    while (reader.Read())
                    {
                        od = GetDataOrderToExcel(reader);
                        list.Add(od);
                    }
                    return list;
                }
            }
        }
        public static List<Product> ImportExcel_Product(string FileName, bool hasHeaders)
        {
            string strConn = GetOleDBConnect(FileName, hasHeaders);
            using (OleDbConnection con = new OleDbConnection(strConn))
            {
                string query = "SELECT * FROM [Sheet1$]";
                OleDbCommand cmd = new OleDbCommand(query, con);
                con.Open();
                using (OleDbDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                {
                    List<Product> list = new List<Product>();
                    Product pro = null;
                    while (reader.Read())
                    {
                        pro = GetDataProductToExcel(reader);
                        list.Add(pro);
                    }
                    return list;
                }
            }
        }
        public static List<Supplier> ImportExcel_Supplier(string FileName, bool hasHeaders)
        {
            string strConn = GetOleDBConnect(FileName, hasHeaders);
            using (OleDbConnection con = new OleDbConnection(strConn))
            {
                string query = "SELECT * FROM [Sheet1$]";
                OleDbCommand cmd = new OleDbCommand(query, con);
                con.Open();
                using (OleDbDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                {
                    List<Supplier> list = new List<Supplier>();
                    Supplier sup = null;
                    while (reader.Read())
                    {
                        sup = GetDataSupplierToExcel(reader);
                        list.Add(sup);
                    }
                    return list;
                }
            }
        }
        public static List<OrderDetail> ImportExcel_OrderDetail(string FileName, bool hasHeaders)
        {
            string strConn = GetOleDBConnect(FileName, hasHeaders);
            using (OleDbConnection con = new OleDbConnection(strConn))
            {
                string query = "SELECT * FROM [Sheet1$]";
                OleDbCommand cmd = new OleDbCommand(query, con);
                con.Open();
                using (OleDbDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                {
                    List<OrderDetail> list = new List<OrderDetail>();
                    OrderDetail odd = null;
                    while (reader.Read())
                    {
                        odd = GetDataOrderDetailToExcel(reader);
                        list.Add(odd);
                    }
                    return list;
                }
            }

        }
    }
}
