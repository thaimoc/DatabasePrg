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
    public partial class NewSupplierForm : Form
    {
        private int _SupplierID = -1;

        public int SupplierID
        {
            get { return _SupplierID; }
            set { _SupplierID = value; }
        }

        public NewSupplierForm()
        {
            InitializeComponent();
        }

        private void NewSupplierForm_Load(object sender, EventArgs e)
        {
            if (SupplierID == -1)
            {
                btnUpdate.Visible = false;
                btnAdd.Enabled = false;
                txtSupplierID.Text = "-- New --";
            }
            else
            {
                btnAdd.Visible = false;
                btnUpdate.Visible = true;
                btnUpdate.Enabled = false;
                controlsLoad();
            }
        }

        private Supplier supplierLoad()
        {
            Supplier newSup = new Supplier(); 
            newSup.SupplierID = SupplierID;
            newSup.CompanyName = txtCompanyName.Text;
            newSup.ContactName = txtContactName.Text;
            newSup.ContactTitle = txtContactTitle.Text;
            newSup.Address = txtAddress.Text;
            newSup.City = txtCity.Text;
            newSup.Region = txtRegion.Text;
            newSup.PostalCode = txtPostalCode.Text;
            newSup.Country = txtCountry.Text;
            newSup.Phone = txtPhone.Text;
            newSup.Fax = txtFax.Text;
            newSup.HomePage = txtHomePage.Text;
            return newSup;
        }

        private void controlsLoad()
        {
            Supplier newSup = Supplier.Single(SupplierID);
            txtSupplierID.Text = newSup.SupplierID.ToString();
            txtCompanyName.Text = newSup.CompanyName;
            txtContactName.Text = newSup.ContactName;
            txtContactTitle.Text = newSup.ContactTitle;
            txtAddress.Text = newSup.Address;
            txtCity.Text = newSup.City;
            txtRegion.Text = newSup.Region;
            txtCountry.Text = newSup.Country;
            txtPostalCode.Text = newSup.PostalCode;
            txtPhone.Text = newSup.Phone;
            txtFax.Text = newSup.Fax;
            txtHomePage.Text = newSup.HomePage;
        }

        private void txtPhone_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || e.KeyChar.ToString() == "(" || e.KeyChar.ToString() == ")" || e.KeyChar.ToString() == "." || e.KeyChar.ToString() == "-")
            {
                e.Handled = false;
            }
            else
                e.Handled = true;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (Supplier.FindByCompany(txtCompanyName.Text).Count > 0)
            {
                MessageBox.Show("This customer ID is exted!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnAdd.Enabled = false;
                return;
            }
            Supplier newSupp = supplierLoad();
            Supplier.Add(newSupp);
            MessageBox.Show("Adding is successful!", "Message Box!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
        }

        private void txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (txtCompanyName.Text.Length == 0)
            {
                btnUpdate.Enabled = false;
                btnAdd.Enabled = false;
            }
            else
            {
                btnUpdate.Enabled = true;
                btnAdd.Enabled = true;
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            Supplier newSup = supplierLoad();
            bool affected = Supplier.Update(newSup);
            if (affected)
            {
                MessageBox.Show("Udating is successfull! New supplier id = " + txtSupplierID.Text, "Message Box", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
                btnUpdate.Enabled = false;
                return;
            }
            else
            {
                MessageBox.Show("Udating is fails", "Message Box", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }            
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
