using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using Core;


namespace DataAccess
{
    public class Product
    {
        #region Fields
        private int _ProductID;

        public int ProductID
        {
            get { return _ProductID; }
            set { _ProductID = value; }
        }
        private string _ProductName;

        public string ProductName
        {
            get { return _ProductName; }
            set { _ProductName = value; }
        }
        private int _SupplierID;

        public int SupplierID
        {
            get { return _SupplierID; }
            set { _SupplierID = value; }
        }
        
        private string _QuantityPerUnit;

        public string QuantityPerUnit
        {
            get { return _QuantityPerUnit; }
            set { _QuantityPerUnit = value; }
        }
        private double _UnitPrice;

        public double UnitPrice
        {
            get { return _UnitPrice; }
            set { _UnitPrice = value; }
        }
        private short _UnitsInStock;

        public short UnitsInStock
        {
            get { return _UnitsInStock; }
            set { _UnitsInStock = value; }
        }
        private short _UnitsOnOrder;

        public short UnitsOnOrder
        {
            get { return _UnitsOnOrder; }
            set { _UnitsOnOrder = value; }
        }

        private bool _Discontinued;

        public bool Discontinued
        {
            get { return _Discontinued; }
            set { _Discontinued = value; }
        }

        private string _CompanyName;

        public string CompanyName
        {
            get { return _CompanyName; }
            set { _CompanyName = value; }
        }


        #endregion

        #region Basic Methods
        /// <summary>
        /// Tìm tất cả các Produts
        /// </summary>
        /// <returns>List</returns>
        public static List<Product> All()
        {
            return CBO.FillCollection<Product>(DataProvider.Instance.ExecuteReader("Products_All"));
        }
        
        /// <summary>
        /// Count of Products
        /// </summary>
        /// <returns>Count</returns>
        public static int Count()
        {
            object rs = DataProvider.Instance.ExecuteScalar("Products_Count");
            return MyConvert.ToInt32(rs);
        }

        /// <summary>
        /// Tìm theo khóa chính
        /// </summary>
        /// <param name="productID">ProductID</param>
        /// <returns>Product</returns>
        public static Product Single(int productID)
        {
            return CBO.FillObject<Product>(DataProvider.Instance.ExecuteReader("Products_Single", productID));
        }
        /// <summary>
        /// Tìm theo khóa ngoại SupplierID
        /// </summary>
        /// <param name="supplierID">SupplierID</param>
        /// <returns>List</returns>
        public static List<Product> FindBySupplierID(int supplierID)
        {
            return CBO.FillCollection<Product>(DataProvider.Instance.ExecuteReader("Products_FindBySupplierID", supplierID));
        }
        /// <summary>
        /// Find by product's name and supplier id
        /// </summary>
        /// <param name="productName">Product's name</param>
        /// <param name="supplierID">Supplier's id</param>
        /// <returns>list<Product></returns>
        public static List<Product> FindByProductNameSupplierID(string productName, int supplierID)
        {
            return CBO.FillCollection<Product>(DataProvider.Instance.ExecuteReader("Products_FindByProductNameSupplierID", productName, supplierID));
        }
        /// <summary>
        /// Find By Like Product Name
        /// </summary>
        /// <param name="productName"></param>
        /// <returns></returns>
        public static List<Product> FindByLikeProductName(string productName)
        {
            return CBO.FillCollection<Product>(DataProvider.Instance.ExecuteReader("Products_FindByLikeProductName", productName));
        }
        /// <summary>
        /// Find By Discontinued
        /// </summary>
        /// <param name="Discontinued">Yes/No</param>
        /// <returns></returns>
        public static List<Product> FindByDiscontinued(string discontinued)
        {
            return CBO.FillCollection<Product>(DataProvider.Instance.ExecuteReader("Products_FindByDiscontinued", discontinued));
        }
        
        /// <summary>
        /// Thêm một sản phẩm mới vào csdl
        /// </summary>
        /// <param name="product">New Product</param>
        /// <returns>ProductID</returns>
        public static int Add(Product product)
        {
            object rs = DataProvider.Instance.ExecuteNonQueryWithOutput("@ProductID", "Products_Add", 0, product._ProductName, product._SupplierID,
                product._QuantityPerUnit, product._UnitPrice, product._UnitsInStock, product._UnitsOnOrder, product._Discontinued);
            return MyConvert.ToInt32(rs);
        }
        /// <summary>
        /// Updete theo khóa chính
        /// </summary>
        /// <param name="product">New Product</param>
        /// <returns>Number or row with was updated</returns>
        public static bool Update(Product product)
        {
            object rs = DataProvider.Instance.ExecuteNonQuery("Products_Update", product._ProductID, product._ProductName, product._SupplierID,
                product._QuantityPerUnit, product._UnitPrice, product._UnitsInStock, product._UnitsOnOrder, product._Discontinued);
            return MyConvert.ToInt32(rs) > 0;
        }
        /// <summary>
        /// Xóa theo khóa chính
        /// </summary>
        /// <param name="product">The product you want to remove</param>
        /// <returns>productID Which was removed</returns>
        public static bool Delete(int productID)
        {
            object rs = DataProvider.Instance.ExecuteNonQuery("Products_Delete", productID);
            return MyConvert.ToInt32(rs) > 0;
        }
        #endregion

        #region Advance Methods

        public static string[] ColumnNames()
        {
            DataSet ds = DataProvider.Instance.ExecuteDataset("Products_All");
            DataTable dt = ds.Tables[0];
            string[] s = new string[dt.Columns.Count];
            int i = 0;
            foreach (DataColumn col in dt.Columns)
            {
                s[i++] = col.ColumnName;
            }
            return s;
        }

        public static List<Product> Find(string expression)
        {
            expression = "SELECT * FROM View_Products WHERE " + expression;
            return CBO.FillCollection<Product>(DataProvider.Instance.ExecuteReader(expression));
        }

        public static List<Product> Sort(string expression)
        {
            expression = "SELECT * FROM View_Products ORDER BY " + expression;
            return CBO.FillCollection<Product>(DataProvider.Instance.ExecuteReader(expression));
        }

        #endregion


        public static List<Product> FindByCustomerID(string customerID)
        {
            return CBO.FillCollection<Product>(DataProvider.Instance.ExecuteReader("Products_FindByCustomerID", customerID));
        }


    }
}
