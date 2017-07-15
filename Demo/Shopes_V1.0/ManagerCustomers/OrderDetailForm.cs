using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DataAccess;

namespace ManagerCustomers
{
    public partial class OrderDetailForm : Form
    {
        private int _OrderID = -1;

        public int OrderID
        {
            get { return _OrderID; }
            set { _OrderID = value; }
        }

        public OrderDetailForm()
        {
            InitializeComponent();
        }

        private void OrderDetailForm_Load(object sender, EventArgs e)
        {
            gvOrderDetails.AutoGenerateColumns = false;
            cbbCustomerLoad(Customer.All(0));
            cbbEmployeeLoad(Employee.All());
            BuildGridView();
            if (_OrderID == -1)
            {
                txtOrderID.Text = "-- New --";
                btnAdd.Visible = true;
                btnUpdate.Visible = false;                
            }
            else
            {
                btnAdd.Visible = false;
                txtOrderID.Text = _OrderID.ToString();
                ControlsLoad();
            }
        }

        private Order OrderLoad()
        {
            Order result = new Order();
            result.OrderID = _OrderID;
            result.CustomerID = (cbbCustomer.SelectedItem as Customer).CustomerID;
            result.EmployeeID = (cbbEmployee.SelectedItem as Employee).EmployeeID;
            result.OrderDate = dtpOrderDate.Value;
            result.RequiredDate = dtpRequiredDate.Value;
            result.ShippedDate = dtpShippedDate.Value;
            result.Freight = MyConvert.ToDouble(txtFreight.Text);
            result.ShipAddress = txtShipAddress.Text;
            result.ShipPostalCode = txtShipPostalCode.Text;
            result.Status = txtStatus.Text;
            return result;
        }

        private void ControlsLoad()
        {
            Order order = Order.Single(_OrderID);
            if (order != null)
            {
                txtOrderID.Text = order.OrderID.ToString();
                for(int i = 0; i < cbbCustomer.Items.Count; i++)
                {
                    if ((cbbCustomer.Items[i] as Customer).CustomerID == order.CustomerID)
                    {
                        cbbCustomer.SelectedIndex = i;
                        break;
                    }
                }
                for (int i = 0; i < cbbEmployee.Items.Count; i++)
                {
                    if ((cbbEmployee.Items[i] as Employee).EmployeeID == order.EmployeeID)
                    {
                        cbbEmployee.SelectedIndex = i;
                        break;
                    }
                }
                dtpOrderDate.Value = order.OrderDate;
                dtpRequiredDate.Value = order.RequiredDate;
                dtpShippedDate.Value = order.ShippedDate;
                txtFreight.Text = order.Freight.ToString();
                txtShipAddress.Text = order.ShipAddress;
                txtShipPostalCode.Text = order.ShipPostalCode;
                txtStatus.Text = order.Status;
                gvOrderDetails.DataSource = OrderDetail.FindByOrderID(_OrderID);                
            }
            else
            {
                MessageBox.Show("Loading data is fails", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }

        private void cbbCustomerLoad(List<Customer> list)
        {
            cbbCustomer.Items.Clear();
            foreach (var item in list)
            {
                cbbCustomer.Items.Add(item);
            }
            cbbCustomer.DisplayMember = "CompanyName";
            cbbCustomer.ValueMember = "CustomerID";
            cbbCustomer.SelectedIndex = -1;
        }

        private void cbbEmployeeLoad(List<Employee> list)
        {
            cbbEmployee.Items.Clear();
            foreach (var item in list)
            {
                cbbEmployee.Items.Add(item);
            }
            cbbEmployee.DisplayMember = "EmployeeName";
            cbbEmployee.ValueMember = "EmployeeID";
            cbbEmployee.SelectedIndex = -1;
        }

        private void BuildGridView()
        {
            DataGridViewComboBoxColumn productColumn = new DataGridViewComboBoxColumn();
            productColumn.DisplayIndex = 0;
            productColumn.Width = 200;
            productColumn.DataPropertyName = "ProductID";
            productColumn.DataSource = Product.All();
            productColumn.DisplayMember = "ProductName";
            productColumn.HeaderText = "Product Name";
            productColumn.Name = "ProductNameColumn";
            productColumn.ValueMember = "ProductID";
            gvOrderDetails.Columns.Insert(0, productColumn);
            gvOrderDetails.CellValueChanged += new DataGridViewCellEventHandler(gvOrderDetails_CellValueChanged);
            gvOrderDetails.DataSource = OrderDetail.BuildDataTable();
        }

        private void gvOrderDetails_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow currentRow = gvOrderDetails.Rows[e.RowIndex];

            object nid = currentRow.Cells[0].Value;
            for (int index = 0; index < e.RowIndex; index++)
            {
                if (gvOrderDetails.Rows[index].Cells[0].Value.ToString() == nid.ToString())
                {
                    MessageBox.Show("This product was had been selected!", "Invalid Value", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    gvOrderDetails.Rows.Remove(currentRow);
                    break;
                }
            }
        }

        private void btnNewCustomer_Click(object sender, EventArgs e)
        {
            NewCustomerForm frm = new NewCustomerForm();
            frm.ShowDialog();
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                cbbCustomerLoad(Customer.All(0));
        }

        private void btnNewEmployee_Click(object sender, EventArgs e)
        {
            NewEmployeeForm frm = new NewEmployeeForm();
            frm.ShowDialog();
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                cbbEmployeeLoad(Employee.All());
        }

        private void txtFreight_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (txtFreight.Text.Contains("."))
            {
                if (Char.IsDigit(e.KeyChar) || Char.IsControl(e.KeyChar))
                    e.Handled = false;
                else
                    e.Handled = true;
            }
            else
            {
                if (Char.IsDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || e.KeyChar == '.')
                    e.Handled = false;
                else
                    e.Handled = true;
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                Order temp = OrderLoad();
                if (gvOrderDetails.Rows.Count > 1)
                {
                    DialogResult dlg = MessageBox.Show("Do you want to add this order!", "Message Box", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (dlg == System.Windows.Forms.DialogResult.OK)
                    {
                        int oder = Order.Add(temp);

                        for (int i = 0; i < gvOrderDetails.Rows.Count - 1; i++)
                        {
                            DataGridViewRow currRow = gvOrderDetails.Rows[i];
                            OrderDetail tempo = new OrderDetail();
                            tempo.OrderID = oder;
                            tempo.ProductID = Convert.ToInt32(currRow.Cells[0].Value.ToString());
                            tempo.UnitPrice = Convert.ToDouble(Convert.ToDecimal(currRow.Cells[1].Value.ToString()));
                            tempo.Quantity = Convert.ToInt16(currRow.Cells[2].Value.ToString());
                            tempo.Discount = Convert.ToDouble(currRow.Cells[3].Value.ToString());
                            OrderDetail.Add(tempo);
                        }
                        this.DialogResult = DialogResult.OK;
                    }
                }
                else
                    MessageBox.Show("You can import order details!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void gvOrderDetails_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("Data error! You does not import corect data type!", "Invalid value");
            e.Cancel = true;
        }

        private void btnCancell_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            DialogResult dlg = MessageBox.Show("Do you want to update this order and the order details?", "?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (dlg == System.Windows.Forms.DialogResult.OK)
            {
                Order.Update(OrderLoad());
                UpdateLoadDowOrderDetails();
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
            }
        }

        private void UpdateLoadDowOrderDetails()
        {
            OrderDetail.DeleteByOrderID(OrderID);
            for (int i = 0; i < gvOrderDetails.Rows.Count - 1; i++)
            {
                DataGridViewRow currRow = gvOrderDetails.Rows[i];
                OrderDetail tempo = new OrderDetail();

                tempo.OrderID = _OrderID;
                tempo.ProductID = Convert.ToInt32(currRow.Cells[0].Value.ToString());

                tempo.UnitPrice = Convert.ToDouble(Convert.ToDecimal(currRow.Cells[1].Value.ToString()));
                tempo.Quantity = Convert.ToInt16(currRow.Cells[2].Value.ToString());
                tempo.Discount = Convert.ToDouble(currRow.Cells[3].Value.ToString());

                OrderDetail.Add(tempo);
            }
        }

        private void dtpOrderDate_ValueChanged(object sender, EventArgs e)
        {
            if (dtpOrderDate.Value > dtpRequiredDate.Value)
            {
                MessageBox.Show("Required date is not less than Order date!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DateTime dt = dtpRequiredDate.Value;
                dtpRequiredDate.Value = dtpRequiredDate.MaxDate;
                dtpOrderDate.Value = DateTime.Now;
                if (dtpOrderDate.Value > dt)
                    dtpRequiredDate.Value = DateTime.Now;
                else
                    dtpRequiredDate.Value = dt;
            }
        }

        private void dtpRequiredDate_ValueChanged(object sender, EventArgs e)
        {
            if (dtpOrderDate.Value > dtpRequiredDate.Value)
            {
                MessageBox.Show("Required date is not less than Order date!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DateTime dt = dtpOrderDate.Value;
                dtpOrderDate.Value = dtpOrderDate.MinDate;
                dtpRequiredDate.Value = DateTime.Now;
                if (dtpRequiredDate.Value < dt)
                    dtpOrderDate.Value = DateTime.Now;
                else
                    dtpOrderDate.Value = dt;                
            }
        }

        private void dtpShippedDate_ValueChanged(object sender, EventArgs e)
        {
            if (dtpOrderDate.Value > dtpShippedDate.Value)
            {
                MessageBox.Show("Shipped date is not less than Order date!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DateTime dt = dtpOrderDate.Value;
                dtpOrderDate.Value = dtpOrderDate.MinDate;
                dtpShippedDate.Value = DateTime.Now;
                if (dtpShippedDate.Value < dt)
                    dtpOrderDate.Value = DateTime.Now;
                else
                    dtpOrderDate.Value = dt;
            }
        }


    }
}
