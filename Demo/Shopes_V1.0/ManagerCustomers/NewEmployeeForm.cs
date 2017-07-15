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
    public partial class NewEmployeeForm : Form
    {
        private int _EmployeeID = -1;

        public int EmployeeID
        {
            get { return _EmployeeID; }
            set { _EmployeeID = value; }
        }
        public NewEmployeeForm()
        {
            InitializeComponent();
        }

        private void NewEmployeeForm_Load(object sender, EventArgs e)
        {
            if (EmployeeID != -1)
            {
                btnUpdate.Enabled = false;
                btnAdd.Visible = false;
                newEmLoadUp();
            }
            else
            {
                btnAdd.Enabled = false;
                btnUpdate.Visible = false;
            }
        }

        private void txtLastName_TextChanged(object sender, EventArgs e)
        {
            if (txtLastName.Text.Length > 0 && txtFirstName.Text.Length > 0)
            {
                btnAdd.Enabled = true;
                btnUpdate.Enabled = true;
            }
            else
            {
                btnAdd.Enabled = false;
                btnUpdate.Enabled = false;
            }
        }

        private void newEmLoadUp()
        {
            Employee newEm = Employee.Single(_EmployeeID);
            if (newEm.EmployeeID > 0)
            {
                txtEmployeeID.Text = newEm.EmployeeID.ToString();
                txtLastName.Text = newEm.LastName;
                txtFirstName.Text = newEm.FirstName;
                dtpBirthDate.Value = newEm.BirthDate;
                dtpHireDate.Value = newEm.HireDate;
                txtAddress.Text = newEm.Address;
                txtPostalCode.Text = newEm.PostalCode;
                txtCountry.Text = newEm.Country;
                txtHomePhone.Text = newEm.HomePhone;
            }
            else
            {
                MessageBox.Show("Loading data is fails!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }

        private Employee newEmLoadDownd()
        {
            Employee newEm = new Employee();
            if (txtFirstName.Text.Length > 0 && txtLastName.Text.Length > 0)
            {
                try
                {
                    newEm.EmployeeID = MyConvert.ToInt32(txtEmployeeID.Text);
                    newEm.LastName = txtLastName.Text;
                    newEm.FirstName = txtFirstName.Text;
                    newEm.BirthDate = dtpBirthDate.Value;
                    newEm.HireDate = dtpHireDate.Value;
                    newEm.Address = txtAddress.Text;
                    newEm.PostalCode = txtPostalCode.Text;
                    newEm.Country = txtCountry.Text;
                    newEm.HomePhone = txtHomePhone.Text;
                }
                catch
                {
                    MessageBox.Show("Loading data is fails!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.Close();
                }
            }
            else
            {
                MessageBox.Show("Loading data is fails!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            return newEm;
        }

        private void btnCancell_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            newEmLoadDownd();
            DialogResult dlg = MessageBox.Show("Do you want to add this Employee?", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (dlg == System.Windows.Forms.DialogResult.OK)
            {
                Employee.Add(newEmLoadDownd());
                MessageBox.Show("Adding is successful!", "Message Box", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
            }

        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            newEmLoadDownd();
            try
            {
                Employee.Update(newEmLoadDownd());
                MessageBox.Show("Updating is successful!", "Message Box", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
            }
            catch
            {
                MessageBox.Show("Updating is fails!", "Message Box", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void txtHomePhone_KeyPress(object sender, KeyPressEventArgs e)
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
