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
    public partial class NewCustomerForm : Form
    {
        private string _CustomerID = "";

        public string CustomerID
        {
            get { return _CustomerID; }
            set { _CustomerID = value; }
        }
        public NewCustomerForm()
        {
            InitializeComponent();
        }

        private void NewCustomerForm_Load(object sender, EventArgs e)
        {
            if (_CustomerID == "")
            {
                btnUpdate.Visible = false;
                btnAdd.Enabled = false;
            }
            else
            {
                btnAdd.Visible = false;
                txtCustomerID.ReadOnly = true;
                controlsLoad();
                btnUpdate.Enabled = false;
            }
        }


        /// <summary>
        /// Hàm kiểm tra nhập companyName vào txtCompanyName chứa
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txt_TextChanged(object sender, EventArgs e)
        {
            if (btnAdd.Visible)
            {
                if (txtCustomerID.Text.Length > 5)
                    txtCustomerID.Text = txtCustomerID.Text.Substring(0, 5);
                if (txtCustomerID.Text.Length > 0 && txtCompanyName.Text.Length > 0)
                    btnAdd.Enabled = true;
                else
                    btnAdd.Enabled = false;
            }
            if(btnUpdate.Visible)
            {
                if (btnUpdate.Enabled == false)
                    btnUpdate.Enabled = true;
            }
        }

        private void controlsLoad()
        {
            Customer customer = Customer.Single(CustomerID);
            txtCustomerID.Text = customer.CustomerID;
            txtCompanyName.Text = customer.CompanyName;
            txtContactName.Text = customer.ContactName;
            txtAddress.Text = customer.Address;
            txtCountry.Text = customer.Country;
            txtPostalCode.Text = customer.PostalCode;
            txtPhone.Text = customer.Phone;
            txtFax.Text = customer.Fax;
        }

        private Customer customerLoad()
        {
            Customer newCus = new Customer();
            newCus.CustomerID = txtCustomerID.Text;
            newCus.CompanyName = txtCompanyName.Text;
            newCus.ContactName = txtContactName.Text;
            newCus.Address = txtAddress.Text;
            newCus.PostalCode = txtPostalCode.Text;
            newCus.Country = txtCountry.Text;
            if (txtPhone.Text != "(   )    -")
                newCus.Phone = txtPhone.Text;
            else
                newCus.Phone = "";
            newCus.Fax = txtFax.Text;
            return newCus;
        }

        private void btnCancell_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            Customer newCus = customerLoad();
            bool id = Customer.Update(newCus);
            if (id)
                MessageBox.Show("Updating is successful!", "Message Box!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
            {
                MessageBox.Show("Updating is faild! And data error", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            // Kiểm tra mã CustomerID xem đã tồn tại hay chưa
            Customer newCus = Customer.Single(txtCustomerID.Text);
            if (newCus != null)
            {
                MessageBox.Show("This customer ID is exted!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnAdd.Enabled = false;
                return;
            }

            newCus = customerLoad();
            Customer.Add(newCus);
            MessageBox.Show("Adding is successful!", "Message Box!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
        }
        /// <summary>
        /// Kiểm tra hợp lệ của số điện thoại nhập vào txtPhone, và txtFax
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtPhone_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || e.KeyChar.ToString() == "(" || e.KeyChar.ToString() == ")" || e.KeyChar.ToString() == "." || e.KeyChar.ToString() == "-")
            {
                e.Handled = false;
            }
            else
                e.Handled = true;
        }
    }
}
