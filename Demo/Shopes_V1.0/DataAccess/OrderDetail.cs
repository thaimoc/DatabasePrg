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
    public class OrderDetail
    {
        #region Fields
        private int _OrderID;

        public int OrderID
        {
            get { return _OrderID; }
            set { _OrderID = value; }
        }
        private int _ProductID;

        public int ProductID
        {
            get { return _ProductID; }
            set { _ProductID = value; }
        }
        private double _UnitPrice;

        public double UnitPrice
        {
            get { return _UnitPrice; }
            set { _UnitPrice = value; }
        }
        private short _Quantity;

        public short Quantity
        {
            get { return _Quantity; }
            set { _Quantity = value; }
        }
        private double _Discount;

        public double Discount
        {
            get { return _Discount; }
            set { _Discount = value; }
        }

        private string _Product;

        public string Product
        {
            get { return _Product; }
            set { _Product = value; }
        }

        #endregion

        #region Basic Methods
        /// <summary>
        /// Tìm theo khóa chính
        /// </summary>
        /// <param name="orderID"></param>
        /// <param name="productID"></param>
        /// <returns></returns>
        public static OrderDetail Single(int orderID, int productID)
        {
            return CBO.FillObject<OrderDetail>(DataProvider.Instance.ExecuteReader("[Order Details_Single]", orderID, productID));
        }
        /// <summary>
        /// Tìm theo khóa ngoại OrderID
        /// </summary>
        /// <param name="orderID">OrderID</param>
        /// <returns>list</returns>
        public static List<OrderDetail> FindByOrderID(int orderID)
        {
            return CBO.FillCollection<OrderDetail>(DataProvider.Instance.ExecuteReader("[Order Details_FindByOrderID]", orderID));
        }
        /// <summary>
        /// Tìm theo khóa ngoại ProductID
        /// </summary>
        /// <param name="productID"><ProductID/param>
        /// <returns>list</returns>
        public static List<OrderDetail> FindByProductID(int productID)
        {
            return CBO.FillCollection<OrderDetail>(DataProvider.Instance.ExecuteReader("[Order Details_FindByProductID]", productID));
        }
        /// <summary>
        /// Tìm tất cả các OrderDetail
        /// </summary>
        /// <returns>list</returns>
        public static List<OrderDetail> All()
        {
            return CBO.FillCollection<OrderDetail>(DataProvider.Instance.ExecuteReader("[Order Details_All]"));
        }
        /// <summary>
        /// Thêm mới một orderDetails vào csdl
        /// </summary>
        /// <param name="orderDetail">New OrderDetail</param>
        public static int Add(OrderDetail orderDetail)
        {
            object rs = DataProvider.Instance.ExecuteNonQuery("[Order Details_Add]", orderDetail._OrderID, orderDetail._ProductID, orderDetail._UnitPrice, orderDetail._Quantity, orderDetail.Discount);
            int affected = MyConvert.ToInt32(rs);
            return affected;
        }
        /// <summary>
        /// Update OrderDetail theo khóa chính
        /// </summary>
        /// <param name="orderDetail">New OrderDetail</param>
        /// <returns>Number rows was changed</returns>
        public static bool Update(OrderDetail orderDetail)
        {
            int rs = DataProvider.Instance.ExecuteNonQuery("[Order Details_Update]", orderDetail._OrderID, orderDetail._ProductID, orderDetail._UnitPrice, orderDetail._Quantity, orderDetail.Discount);
            return rs > 0;
        }
        /// <summary>
        /// Xóa theo khóa chính la OrderID và ProductID
        /// </summary>
        /// <param name="orderDetail">OrderDetail</param>
        /// <returns>employeeID đã bị delete</returns>
        public static bool Delete(int orderID, int productID)
        {
            try
            {
                int rs = DataProvider.Instance.ExecuteNonQuery("[Order Details_Delete]", orderID, productID);
                return rs > 0;
            }
            catch
            {
                return false;
            }
        }

        public static bool DeleteByOrderID(int orderID)
        {
            try
            {
                int rs = DataProvider.Instance.ExecuteNonQuery("[Order Details_DeleteByOrderID]", orderID);
                return rs > 0;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Advance Methods

        public static DataTable BuildDataTable()
        {
            DataTable productsTable = new DataTable("Products");
            using (SqlConnection connect = Helper.GetSqlConnection())
            {
                SqlCommand command = connect.CreateCommand();
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = "[Order Details_BuildData]";
                connect.Open();
                using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                {
                    adapter.FillSchema(productsTable, SchemaType.Mapped);
                }
            }
            return productsTable;           
        }
        

        public static DataTable BuildDataTable(int orderID)
        {
            DataTable productsTable = new DataTable("Products");
            using (SqlConnection connect = Helper.GetSqlConnection())
            {
                SqlCommand command = connect.CreateCommand();
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = "[Order Details_BuildData_OrderId]";
                command.Parameters.Add("@OrderID", SqlDbType.Int).Value = orderID;
                connect.Open();
                using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                {
                    adapter.FillSchema(productsTable, SchemaType.Mapped);
                    adapter.Fill(productsTable);
                }
            }
            return productsTable;
        }
        

        #endregion

    }
}
