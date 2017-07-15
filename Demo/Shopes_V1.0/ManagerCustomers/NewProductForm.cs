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
    public partial class NewProductForm : Form
    {
        private int _ProductID = -1;

        public int ProductID
        {
            get { return _ProductID; }
            set { _ProductID = value; }
        }

        public NewProductForm()
        {
            InitializeComponent();
        }

        private void NewProductForm_Load(object sender, EventArgs e)
        {
            cbbSupplierLoad();
            if (ProductID == -1)
            {
                btnAdd.Visible = true;
                btnUpdate.Visible = false;
                btnAdd.Enabled = false;
                txtProductID.Text = "--New--";

            }
            else
            {
                btnAdd.Visible = false;
                btnUpdate.Visible = true;
                btnUpdate.Enabled = false;
                productLoad();
            }
        }

        private void txtUnitPrice_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (txtUnitPrice.Text.Contains("."))
            {
                if (Char.IsDigit(e.KeyChar) || Char.IsControl(e.KeyChar))
                {
                    e.Handled = false;
                }
                else
                    e.Handled = true;
            }
            else
            {
                if (Char.IsDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || e.KeyChar.ToString() == ".")
                {
                    e.Handled = false;
                }
                else
                    e.Handled = true;
            }
        }

        private void txtUnitInStoct_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (MyConvert.ToInt32(txtUnitInStoct.Text) > 255)
                txtUnitInStoct.Text = "";
            if (Char.IsDigit(e.KeyChar) || Char.IsControl(e.KeyChar))
            {
                e.Handled = false;
            }
            else
                e.Handled = true;
        }
        /// <summary>
        /// Load Suppliers in to combobox
        /// </summary>
        private void cbbSupplierLoad()
        {
            cbbSupplier.Items.Clear();
            List<Supplier> list = Supplier.All();
            foreach (Supplier item in list)
            {
                cbbSupplier.Items.Add(item);
            }
            cbbSupplier.ValueMember = "SupplierID";
            cbbSupplier.DisplayMember = "CompanyName";
        }

        private void btnNewSupplier_Click(object sender, EventArgs e)
        {
            NewSupplierForm frm = new NewSupplierForm();
            frm.ShowDialog();
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                cbbSupplierLoad();
        }

        private Product controlsLoad()
        {
            Product newPro = new Product();
            newPro.ProductID = ProductID;
            newPro.ProductName = txtProductName.Text;

            Supplier temp = cbbSupplier.SelectedItem as Supplier;

            newPro.SupplierID = temp.SupplierID;

            newPro.QuantityPerUnit = txtQuantityPerUnit.Text;
            newPro.UnitPrice = MyConvert.ToDouble(txtUnitPrice.Text);
            newPro.UnitsInStock = MyConvert.ToShort(txtUnitInStoct.Text);
            newPro.UnitsOnOrder = MyConvert.ToShort(txtUnitOnOrder.Text);
            newPro.Discontinued = ckbDiscontinued.Checked;
            return newPro;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (cbbSupplier.SelectedIndex > 0)
            {
                Supplier sup = cbbSupplier.SelectedItem as Supplier;
                List<Product> list = Product.FindByProductNameSupplierID(txtProductName.Text, sup.SupplierID);
                if (list.Count > 0)
                {
                    MessageBox.Show("This Product had been exited!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    btnAdd.Enabled = false;
                    return;
                }
                Product newPro = controlsLoad();
                Product.Add(newPro);
                MessageBox.Show("Adding is successful!", "Message Box!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
            }
            else
                MessageBox.Show("You don't choose Supplier!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void txtProductName_TextChanged(object sender, EventArgs e)
        {
            if (txtProductName.Text.Length > 0 && cbbSupplier.SelectedIndex > 0)
            {
                btnUpdate.Enabled = true;
                btnAdd.Enabled = true;
            }
            else
            {
                btnUpdate.Enabled = false;
                btnAdd.Enabled = false;
            }
        }

        private void cbbSupplier_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtProductName.Text.Length > 0 && cbbSupplier.SelectedIndex > 0)
            {
                btnUpdate.Enabled = true;
                btnAdd.Enabled = true;
            }
            else
            {
                btnUpdate.Enabled = false;
                btnAdd.Enabled = false;
            }
        }

        private void productLoad()
        {
            Product temp = Product.Single(ProductID);
            txtProductID.Text = temp.ProductID.ToString();
            txtProductName.Text = temp.ProductName;

            for (int i = 0; i < cbbSupplier.Items.Count; i++)
            {
                Supplier spl = cbbSupplier.Items[i] as Supplier;
                if (temp.SupplierID == spl.SupplierID)
                {
                    cbbSupplier.SelectedIndex = i;
                    break;
                }
            }
            txtQuantityPerUnit.Text = temp.QuantityPerUnit;
            txtUnitPrice.Text = temp.UnitPrice.ToString();
            txtUnitInStoct.Text = temp.UnitsInStock.ToString();
            txtUnitOnOrder.Text = temp.UnitsOnOrder.ToString();
            ckbDiscontinued.Checked = temp.Discontinued;
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                Product.Update(controlsLoad());
                MessageBox.Show("Updating is successfull!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Updating is fails" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancell_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        
    }
}
