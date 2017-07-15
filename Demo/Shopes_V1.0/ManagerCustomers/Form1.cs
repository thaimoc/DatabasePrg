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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            panelCustomersLoand();
            tabPageSuppliersLoad();
            tabPageProductsLoad();
            tapageEmployeesLoad();
            tabPageOrdersLoad();
        }

        int i = 5;
        private void timerCustomerSupport_Tick(object sender, EventArgs e)
        {
            lblCustomerSupport.Left += i;
            if (lblCustomerSupport.Left >= this.Width - lblCustomerSupport.Width)
            {
                i = -5;
                lblCustomerSupport.ForeColor = Color.Blue;
                lblSupplierSupport.ForeColor = Color.LightGreen;
            }
            if (lblCustomerSupport.Left <= 0)
            {
                i = 5;
                lblCustomerSupport.ForeColor = Color.Red;
                lblSupplierSupport.ForeColor = Color.BlueViolet;
            }
            Text = Text.Substring(Text.Length - 1, 1) + Text.Substring(0, Text.Length - 1);
            txtWelcome.Text = txtWelcome.Text.Substring(1, txtWelcome.Text.Length - 1) + txtWelcome.Text.Substring(0, 1);
            lblSupplierSupport.Text = lblSupplierSupport.Text.Substring(1, lblSupplierSupport.Text.Length - 1) + lblSupplierSupport.Text.Substring(0, 1);
            lblServiceProduct.Text = lblServiceProduct.Text.Substring(1, lblServiceProduct.Text.Length - 1) + lblServiceProduct.Text.Substring(0, 1);
            lblEmployeesSupport.Text = lblEmployeesSupport.Text.Substring(1, lblEmployeesSupport.Text.Length - 1) + lblEmployeesSupport.Text.Substring(0, 1);
            lblOrdersSupport.Text = lblOrdersSupport.Text.Substring(1, lblOrdersSupport.Text.Length - 1) + lblOrdersSupport.Text.Substring(0, 1);
            lblAuthor.Text = lblAuthor.Text.Substring(1, lblAuthor.Text.Length - 1) + lblAuthor.Text.Substring(0, 1);
        }

        #region Customers
        private void panelCustomersLoand()
        {
            listViewCustomersLoad(Customer.All(0));
            cbbCountryLoad();
            timerCustomerSupport.Start();
        }

        private void listViewCustomersLoad(List<Customer> list)
        {
            int count;
            listViewCustomers.Items.Clear();
            foreach (Customer item in list)
            {
                customerLoad(item);
            }
            count = listViewCustomers.Items.Count;
            for (int i = 0; i < count; i++)
            {
                if (i % 2 != 0)
                {
                    listViewCustomers.Items[i].BackColor = Color.PowderBlue;
                }
            }
            lblCountCustomers.Text = count + "/" + Customer.Count().ToString();
        }
        private void customerLoad(Customer customer)
        {
            ListViewItem item = listViewCustomers.Items.Add(customer.CustomerID);
            item.SubItems.Add(customer.CompanyName);
            item.SubItems.Add(customer.ContactName);
            item.SubItems.Add(customer.Address);
            item.SubItems.Add(customer.PostalCode);
            item.SubItems.Add(customer.Country);
            item.SubItems.Add(customer.Phone);
            item.SubItems.Add(customer.Fax);
        }
        private void listViewCustomers_Click(object sender, EventArgs e)
        {
            if (listViewCustomers.SelectedItems.Count > 0)
                listViewCustomers.SelectedItems[0].Checked = !listViewCustomers.SelectedItems[0].Checked;
        }
        private void listViewCustomers_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listViewCustomers.SelectedItems.Count > 0)
            {
                listViewCustomerLoad(listViewCustomers.SelectedItems[0]);
                listViewListProductLoad();
            }
        }

        private void listViewCustomerLoad(ListViewItem item)
        {
            try
            {
                listViewCustomer.Items[0].Text = "Customer ID: " + item.SubItems[0].Text;
                listViewCustomer.Items[0].ImageIndex = 0;
                listViewCustomer.Items[1].Text = "Company Name: " + item.SubItems[1].Text;
                listViewCustomer.Items[1].ImageIndex = 1;
                listViewCustomer.Items[2].Text = "Contact Name: " + item.SubItems[2].Text;
                listViewCustomer.Items[2].ImageIndex = 2;
                listViewCustomer.Items[3].Text = "Address: " + item.SubItems[3].Text;
                listViewCustomer.Items[3].ImageIndex = 3;
                listViewCustomer.Items[4].Text = "Postal Code: " + item.SubItems[4].Text;
                listViewCustomer.Items[4].ImageIndex = 4;
                listViewCustomer.Items[5].Text = "Country: " + item.SubItems[5].Text;
                listViewCustomer.Items[6].Text = "Phone: " + item.SubItems[6].Text;
                listViewCustomer.Items[6].ImageIndex = 5;
                listViewCustomer.Items[7].Text = "Fax: " + item.SubItems[7].Text;
                listViewCustomer.Items[7].ImageIndex = 6;
                listViewCustomer.View = View.List;
            }
            catch { }
        }

        private void toolStripButtonAdd_Click(object sender, EventArgs e)
        {
            NewCustomerForm frm = new NewCustomerForm();
            frm.ShowDialog();
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                toolStripButtonAll.PerformClick();
        }

        private void toolStripButtonUpdate_Click(object sender, EventArgs e)
        {
            if (listViewCustomers.SelectedItems.Count > 0)
            {
                NewCustomerForm frm = new NewCustomerForm();
                frm.CustomerID = listViewCustomers.SelectedItems[0].SubItems[0].Text;
                frm.ShowDialog();
                if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                    toolStripButtonAll.PerformClick();
            }
        }

        private void toolStripButtonDelete_Click(object sender, EventArgs e)
        {
            if (listViewCustomers.CheckedItems.Count > 0)
            {
                DialogResult dlg = MessageBox.Show("Do you want to delete the all customer whitch you checked?", "Message Box", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (dlg == System.Windows.Forms.DialogResult.OK)
                {
                    List<Customer> list = new List<Customer>();
                    foreach (ListViewItem item in listViewCustomers.Items)
                    {
                        if (item.Checked)
                            list.Add(Customer.Single(item.Text));
                    }

                    CustomerDelete(list);
                    MessageBox.Show("Deleting is successful!", "Message Box", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    toolStripButtonAll.PerformClick();
                }
            }
        }

        private void CustomerDelete(List<Customer> list)
        {
            foreach (Customer item in list)
            {
                Customer.Delete(item.CustomerID);
            }
        }

        private void toolStripButtonAll_Click(object sender, EventArgs e)
        {
            listViewCustomersLoad(Customer.All(0));
        }

        private void toolStripButtonFindCompanyName_Click(object sender, EventArgs e)
        {
            listViewCustomersLoad(Customer.FindByCompanyName(toolStripTextBoxCompanyName.Text));
        }

        private void cbbCountryLoad()
        {
            cbbCountry.Items.Clear();
            List<Customer> list = Customer.Countries();
            foreach (var item in list)
            {
                cbbCountry.Items.Add(item.Country);
            }
            cbbCountry.Items.Insert(0, "--- ALL ---");
            cbbCountry.SelectedIndex = 0;
        }

        private void cbbCountry_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbbCountry.SelectedIndex > 0)
            {
                listViewCustomersLoad(Customer.FindByCountry(cbbCountry.SelectedItem.ToString()));
            }
            else
                toolStripButtonAll.PerformClick();
        }

        private void listViewListProductLoad()
        {
            if(listViewCustomers.SelectedItems.Count > 0)
            {
                listViewListProducts.Items.Clear();
                List<Product> list = Product.FindByCustomerID(listViewCustomers.SelectedItems[0].SubItems[0].Text);
                foreach (var item in list)
                {
                    listViewListProducts.Items.Add(item.ProductName);
                }
            }
        }

        private void aSCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listViewCustomersLoad(Customer.All(1));
        }

        private void dESCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listViewCustomersLoad(Customer.All(-1));
        }

        private void toolStripButtonFindCustomers_Click(object sender, EventArgs e)
        {
            FindForm frm = new FindForm();
            frm.cbbColumnLoad(Customer.ColumnNames());
            frm.ShowDialog();
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                listViewCustomersLoad(Customer.Find(frm.Expression));
        }

        private void toolStripButtonSortCustomers_Click(object sender, EventArgs e)
        {
            SortForm frm = new SortForm();
            frm.cklbColumnsLoad(Customer.ColumnNames());
            frm.ShowDialog();
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                listViewCustomersLoad(Customer.Sort(frm.Expression));
        }

        #endregion Customer

        #region Suppliers

        private void tabPageSuppliersLoad()
        {
            listViewSuppliersLoad(Supplier.All());
            cbbCountriesLoad();
        }

        private void toolStripButtonAllSuppliers_Click(object sender, EventArgs e)
        {
            listViewSuppliersLoad(Supplier.All());
        }

        private void listViewSuppliersLoad(List<Supplier> list)
        {
            int count;
            listViewSuppliers.Items.Clear();
            foreach (Supplier item in list)
            {
                suppliersLoad(item);
            }
            count = listViewSuppliers.Items.Count;
            for (int i = 0; i < count; i++)
            {
                if (i % 2 != 0)
                {
                    listViewSuppliers.Items[i].BackColor = Color.PowderBlue;
                }
            }
            lblCountOfSuppliers.Text = count + "/" + Supplier.Count();
        }

        private void suppliersLoad(Supplier supplier)
        {
            ListViewItem item = listViewSuppliers.Items.Add(supplier.SupplierID.ToString());
            item.SubItems.Add(supplier.CompanyName);
            item.SubItems.Add(supplier.ContactName);
            item.SubItems.Add(supplier.ContactTitle);
            item.SubItems.Add(supplier.Address);
            item.SubItems.Add(supplier.City);
            item.SubItems.Add(supplier.Region);
            item.SubItems.Add(supplier.PostalCode);
            item.SubItems.Add(supplier.Country);
            item.SubItems.Add(supplier.Phone);
            item.SubItems.Add(supplier.Fax);
            item.SubItems.Add(supplier.HomePage);
        }

        private void listViewSuppliers_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listViewSuppliers.SelectedItems.Count > 0)
            {
                listViewSupplierLoad(listViewSuppliers.SelectedItems[0]);
                listView3Load();
            }
        }

        private void listViewSupplierLoad(ListViewItem item)
        {
            try
            {
                listViewSupplier.Items[0].Text = "Supplier ID: " + item.SubItems[0].Text;
                listViewSupplier.Items[0].ImageIndex = 0;
                listViewSupplier.Items[1].Text = "Company: " + item.SubItems[1].Text;
                listViewSupplier.Items[1].ImageIndex = 1;
                listViewSupplier.Items[2].Text = "Contact: " + item.SubItems[2].Text;
                listViewSupplier.Items[2].ImageIndex = 2;
                listViewSupplier.Items[3].Text = "Contact Title: " + item.SubItems[3].Text;
                listViewSupplier.Items[3].ImageIndex = 2;
                listViewSupplier.Items[4].Text = "Address: " + item.SubItems[4].Text;
                listViewSupplier.Items[4].ImageIndex = 3;
                listViewSupplier.Items[5].Text = "City: " + item.SubItems[5].Text;
                listViewSupplier.Items[6].Text = "Region: " + item.SubItems[6].Text;
                listViewSupplier.Items[7].Text = "Postal Code: " + item.SubItems[7].Text;
                listViewSupplier.Items[7].ImageIndex = 4;
                listViewSupplier.Items[8].Text = "Country: " + item.SubItems[8].Text;
                listViewSupplier.Items[9].Text = "Phone: " + item.SubItems[9].Text;
                listViewSupplier.Items[9].ImageIndex = 5;
                listViewSupplier.Items[10].Text = "Fax: " + item.SubItems[10].Text;
                listViewSupplier.Items[10].ImageIndex = 6;
                listViewSupplier.Items[11].Text = "Home Page: " + item.SubItems[11].Text;
                listViewSupplier.Items[11].ImageIndex = 3;
                listViewSupplier.View = View.List;
            }
            catch { }
        }

        private void listViewSuppliers_Click(object sender, EventArgs e)
        {
            if (listViewSuppliers.SelectedItems.Count > 0)
                listViewSuppliers.SelectedItems[0].Checked = !listViewSuppliers.SelectedItems[0].Checked;
        }

        private void toolStripButtonFindByCompany_Click(object sender, EventArgs e)
        {
            listViewSuppliersLoad(Supplier.FindByLikeCompany(toolStripTextBoxCompany.Text));
        }

        private void cbbCountriesLoad()
        {
            cbbCountries.Items.Clear();
            List<Supplier> list = Supplier.Countries();
            foreach (var item in list)
            {
                cbbCountries.Items.Add(item.Country);
            }
            cbbCountries.Items.Insert(0, "--- ALL ---");
            cbbCountries.SelectedIndex = 0;
        }

        private void cbbCountries_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbbCountries.SelectedIndex > 0)
            {
                listViewSuppliersLoad(Supplier.FindByCountry(cbbCountries.SelectedItem.ToString()));
            }
            else
                toolStripButtonAllSuppliers.PerformClick();
        }

        private void toolStripButtonAddSupplier_Click(object sender, EventArgs e)
        {
            NewSupplierForm frm = new NewSupplierForm();
            frm.ShowDialog();
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                listViewSuppliersLoad(Supplier.All());
        }

        private void toolStripButtonUpdateSupplier_Click(object sender, EventArgs e)
        {
            if (listViewSuppliers.SelectedItems.Count > 0)
            {
                NewSupplierForm frm = new NewSupplierForm();
                frm.SupplierID = MyConvert.ToInt32(listViewSuppliers.SelectedItems[0].SubItems[0].Text);
                frm.ShowDialog();
                if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                    listViewSuppliersLoad(Supplier.All());
            }
        }
        //Click vào nút toolStripButtonDeleteSuppliers
        private void toolStripButtonDeleteSuppliers_Click(object sender, EventArgs e)
        {
            if (listViewSuppliers.CheckedItems.Count > 0)
            {
                DialogResult dlg = MessageBox.Show("Do you want to delete the all suppliers whitch you checked?", "Message Box", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (dlg == System.Windows.Forms.DialogResult.OK)
                {
                    List<Supplier> list = new List<Supplier>();
                    foreach (ListViewItem item in listViewSuppliers.Items)
                    {
                        if (item.Checked)
                            list.Add(Supplier.Single(MyConvert.ToInt32(item.Text)));
                    }

                    SuppliersDelete(list);
                    MessageBox.Show("Deleting is successful!", "Message Box", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    toolStripButtonAllSuppliers.PerformClick();
                }
            }
        }

        private void SuppliersDelete(List<Supplier> list)
        {
            foreach (Supplier item in list)
            {
                Supplier.Delete(item.SupplierID);
            }
        }

        private void listView3Load()
        {
            if (listViewSuppliers.SelectedItems.Count > 0)
            {
                listView3.Items.Clear();
                List<Product> list = Product.FindBySupplierID(MyConvert.ToInt32(listViewSuppliers.SelectedItems[0].SubItems[0].Text));
                foreach (var item in list)
                {
                    listView3.Items.Add(item.ProductName);
                }
            }
        }

        private void toolStripButtonFindSuppliers_Click(object sender, EventArgs e)
        {
            FindForm frm = new FindForm();
            frm.cbbColumnLoad(Supplier.ColumnNames());
            frm.ShowDialog();
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                listViewSuppliersLoad(Supplier.Find(frm.Expression));
        }

        private void toolStripButtonSortSuppliers_Click(object sender, EventArgs e)
        {
            SortForm frm = new SortForm();
            frm.cklbColumnsLoad(Supplier.ColumnNames());
            frm.ShowDialog();
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                listViewSuppliersLoad(Supplier.Sort(frm.Expression));
        }

        #endregion Suppliers

        #region Products

        private void tabPageProductsLoad()
        {
            listViewProductsLoad(Product.All());
        }

        private void listViewProductsLoad(List<Product> list)
        {
            int count;
            listViewProducts.Items.Clear();
            foreach (var item in list)
            {
                ProductLoad(item);
            }
            count = listViewProducts.Items.Count;
            for (int i = 0; i < count; i++)
            {
                if (i % 2 != 0)
                {
                    listViewProducts.Items[i].BackColor = Color.PowderBlue;
                }
            }
            lblCountOfProducts.Text = count + "/" + Product.Count();
        }

        private void ProductLoad(Product p)
        {
            ListViewItem item = listViewProducts.Items.Add(p.ProductID.ToString());
            item.SubItems.Add(p.ProductName);
            if (p.CompanyName != null)
                item.SubItems.Add(p.CompanyName);
            else
                item.SubItems.Add(p.SupplierID.ToString());
            item.SubItems.Add(p.QuantityPerUnit);
            item.SubItems.Add(p.UnitPrice.ToString());
            item.SubItems.Add(p.UnitsInStock.ToString());
            item.SubItems.Add(p.UnitsOnOrder.ToString());
            if (p.Discontinued)
                item.SubItems.Add("Yes");
            else
                item.SubItems.Add("No");
        }

        private void listViewProducts_Click(object sender, EventArgs e)
        {
            if (listViewProducts.SelectedItems.Count > 0)
                listViewProducts.SelectedItems[0].Checked = !listViewProducts.SelectedItems[0].Checked;
        }

        private void listViewProducts_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listViewProducts.SelectedItems.Count > 0)
            {
                listViewProductLoad(listViewProducts.SelectedItems[0]);
                listView6Load(listViewProducts.SelectedItems[0]);
            }
        }

        private void listViewProductLoad(ListViewItem item)
        {
            try
            {
                listViewProduct.Items[0].Text = "ID: " + item.SubItems[0].Text;
                listViewProduct.Items[1].Text = "Product: " + item.SubItems[1].Text;
                listViewProduct.Items[2].Text = "Supplier:" + item.SubItems[2].Text;
                listViewProduct.Items[3].Text = "Quantity Per Unit: " + item.SubItems[3].Text;
                listViewProduct.Items[4].Text = "Unit Price: " + item.SubItems[4].Text;
                listViewProduct.Items[5].Text = "Units In Stock: " + item.SubItems[5].Text;
                listViewProduct.Items[6].Text = "Units On Order: " + item.SubItems[6].Text;
                listViewProduct.Items[7].Text = "Discontinued: " + item.SubItems[7].Text;
                listViewProduct.View = View.List;
            }
            catch { }
        }

        private void listView6Load(ListViewItem item)
        {
            Product p = Product.Single(MyConvert.ToInt32(item.SubItems[0].Text));
            Supplier sp = Supplier.Single(p.SupplierID);
            try
            {
                listView6.Items[0].Text = "Supplier ID: " + sp.SupplierID.ToString();
                listView6.Items[1].Text = "Company: " + sp.CompanyName;
                listView6.Items[2].Text = "Contact: " + sp.ContactName;
                listView6.Items[3].Text = "Contact Title: " + sp.ContactTitle;
                listView6.Items[4].Text = "Address: " + sp.Address;
                listView6.Items[5].Text = "City: " + sp.City;
                listView6.Items[6].Text = "Region: " + sp.Region;
                listView6.Items[7].Text = "Postal Code: " + sp.PostalCode;
                listView6.Items[8].Text = "Country: " + sp.Country;
                listView6.Items[9].Text = "Phone: " + sp.Phone;
                listView6.Items[10].Text = "Fax: " + sp.Fax;
                listView6.Items[11].Text = "Home Page: " + sp.HomePage;
                listView6.View = View.List;
            }
            catch { }
        }

        private void toolStripButtonProductsAll_Click(object sender, EventArgs e)
        {
            listViewProductsLoad(Product.All());
        }

        private void toolStripButtonProductFindByProduct_Click(object sender, EventArgs e)
        {
            listViewProductsLoad(Product.FindByLikeProductName(toolStripTextBoxFindByProduct.Text));
        }

        private void cbbDiscontinued_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbbDiscontinued.SelectedIndex > 0)
            {
                listViewProductsLoad(Product.FindByDiscontinued(cbbDiscontinued.SelectedItem.ToString()));
            }
            else
                toolStripButtonProductsAll.PerformClick();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            NewProductForm frm = new NewProductForm();
            frm.ShowDialog();
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                listViewProductsLoad(Product.All());
            toolStripButtonAllSuppliers.PerformClick();
        }

        private void toolStripButtonProductUpdate_Click(object sender, EventArgs e)
        {
            if (listViewProducts.SelectedItems.Count > 0)
            {
                NewProductForm frm = new NewProductForm();
                frm.ProductID = MyConvert.ToInt32(listViewProducts.SelectedItems[0].SubItems[0].Text);
                frm.ShowDialog();
                if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                    listViewProductsLoad(Product.All());
                toolStripButtonAllSuppliers.PerformClick();
            }
        }

        private void toolStripButtonDeleteProduct_Click(object sender, EventArgs e)
        {
            if (listViewProducts.CheckedItems.Count > 0)
            {
                DialogResult dlg = MessageBox.Show("Do you want to delete the all products which were had been selected?", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dlg == System.Windows.Forms.DialogResult.OK)
                {
                    foreach (ListViewItem item in listViewProducts.CheckedItems)
                    {
                        Product.Delete(MyConvert.ToInt32(item.SubItems[0].Text));
                    }
                    MessageBox.Show("Deleting is successfull", "Message Box", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    listViewProductsLoad(Product.All());
                }
            }
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            FindForm frm = new FindForm();
            frm.cbbColumnLoad(Product.ColumnNames());
            frm.ShowDialog();
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                listViewProductsLoad(Product.Find(frm.Expression));
        }

        private void toolStripButtonSortProducts_Click(object sender, EventArgs e)
        {
            SortForm frm = new SortForm();
            frm.cklbColumnsLoad(Product.ColumnNames());
            frm.ShowDialog();
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                listViewProductsLoad(Product.Sort(frm.Expression));
        }

        #endregion Products

        #region Employees

        private void tapageEmployeesLoad()
        {
            listViewEmployeesLoad(Employee.All());
            listViewBirdayCurrentMothLoad();
        }

        private void listViewEmployeesLoad(List<Employee> list)
        {
            int count;
            listViewEmployees.Items.Clear();
            foreach (Employee item in list)
            {
                EmployeeLoad(item);
            }
            count = listViewEmployees.Items.Count;
            for (int i = 0; i < count; i++)
            {
                if (i % 2 == 0)
                    listViewEmployees.Items[i].ForeColor = Color.Chocolate;
            }
        }

        private void EmployeeLoad(Employee employee)
        {
            ListViewItem item = listViewEmployees.Items.Add(employee.EmployeeID.ToString());
            item.SubItems.Add(employee.EmployeeName);
            item.SubItems.Add(employee.BirthDate.ToShortDateString());
            item.SubItems.Add(employee.HireDate.ToShortDateString());
            item.SubItems.Add(employee.Address);
            item.SubItems.Add(employee.PostalCode);
            item.SubItems.Add(employee.Country);
            item.SubItems.Add(employee.HomePhone);
        }

        private void listViewEmployees_Click(object sender, EventArgs e)
        {
            if (listViewEmployees.SelectedItems.Count > 0)
                listViewEmployees.SelectedItems[0].Checked = !listViewEmployees.SelectedItems[0].Checked;
        }

        private void listViewEmployees_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listViewEmployees.SelectedItems.Count > 0)
            {
                listViewEmployeeLoad(listViewEmployees.SelectedItems[0]);
                treeViewOrdersOfEmployeeLoad(MyConvert.ToInt32(listViewEmployees.SelectedItems[0].Text));
            }
        }

        private void listViewEmployeeLoad(ListViewItem item)
        {
            try
            {
                listViewEmployee.Items[0].Text = "ID: " + item.SubItems[0].Text;
                listViewEmployee.Items[1].Text = "Name: " + item.SubItems[1].Text;
                listViewEmployee.Items[2].Text = "Birth Date:" + item.SubItems[2].Text;
                listViewEmployee.Items[3].Text = "Hire Date: " + item.SubItems[3].Text;
                listViewEmployee.Items[4].Text = "Address: " + item.SubItems[4].Text;
                listViewEmployee.Items[5].Text = "Postal Code: " + item.SubItems[5].Text;
                listViewEmployee.Items[6].Text = "Country: " + item.SubItems[6].Text;
                listViewEmployee.Items[7].Text = "Home Phone: " + item.SubItems[7].Text;
                listViewEmployee.View = View.Tile;
            }
            catch { }
        }

        private void treeViewOrdersOfEmployeeLoad(int employeeID)
        {
            treeViewOrdersOfEmployee.Nodes.Clear();
            if (listViewEmployees.SelectedItems.Count > 0)
            {
                List<Order> list = Order.FindByEmployeeID(MyConvert.ToInt32(listViewEmployees.SelectedItems[0].Text));
                foreach (Order item in list)
                {
                    TreeNode node = treeViewOrdersOfEmployee.Nodes.Add("Order Date: " + item.OrderDate.ToShortDateString());
                    node.Nodes.Add("Customer: " + item.Customer);
                    node.Nodes.Add("Required Date: " + item.RequiredDate.ToShortDateString());
                    node.Nodes.Add("Ship Address: " + item.ShipAddress);
                    node.Nodes.Add("Status: " + item.Status);
                    node.Expand();
                }
            }
            
        }

        private void listViewBirdayCurrentMothLoad()
        {
            listViewBirdayCurrentMoth.Items.Clear();
            List<Employee> list = Employee.FindByMonthBirthDate(DateTime.Now);
            foreach (var item in list)
            {
                ListViewItem i = listViewBirdayCurrentMoth.Items.Add(item.EmployeeName);
                i.SubItems.Add(item.BirthDate.ToShortDateString());
            }
        }

        private void toolStripButtonFindByName_Click(object sender, EventArgs e)
        {
            listViewEmployeesLoad(Employee.FindByName(toolStripTextBoxFindByName.Text));
        }

        private void toolStripButtonAddNewEmployee_Click(object sender, EventArgs e)
        {
            NewEmployeeForm frm = new NewEmployeeForm();
            frm.ShowDialog();
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                listViewEmployeesLoad(Employee.All());
        }

        private void toolStripButtonUpdateEmployee_Click(object sender, EventArgs e)
        {
            if (listViewEmployees.SelectedItems.Count > 0)
            {
                NewEmployeeForm frm = new NewEmployeeForm();
                frm.EmployeeID = MyConvert.ToInt32(listViewEmployees.SelectedItems[0].Text);
                frm.ShowDialog();
                if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                    listViewEmployeesLoad(Employee.All());
            }
        }

        private void toolStripButtonDeleteEmployees_Click(object sender, EventArgs e)
        {
            if (listViewEmployees.CheckedItems.Count > 0)
            {
                DialogResult dlg = MessageBox.Show("Do you want to delete the employees which was been checked!", "Quession", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dlg == System.Windows.Forms.DialogResult.OK)
                {
                    foreach (ListViewItem item in listViewEmployees.CheckedItems)
                    {
                        Employee.Delete(MyConvert.ToInt32(item.SubItems[0].Text));
                    }
                    listViewEmployeesLoad(Employee.All());
                }
            }
        }

        private void toolStripButtonFindEmployees_Click(object sender, EventArgs e)
        {
            FindForm frm = new FindForm();
            frm.cbbColumnLoad(Employee.ColumnNames());
            frm.ShowDialog();
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                listViewEmployeesLoad(Employee.Find(frm.Expression));
        }

        private void toolStripButtonSortEmployees_Click(object sender, EventArgs e)
        {
            SortForm frm = new SortForm();
            frm.cklbColumnsLoad(Employee.ColumnNames());
            frm.ShowDialog();
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                listViewEmployeesLoad(Employee.Sort(frm.Expression));
        }

        private void toolStripButton12_Click(object sender, EventArgs e)
        {
            listViewEmployeesLoad(Employee.All());
        }

        #endregion Employees

        #region Orders

        private void tabPageOrdersLoad()
        {
            listViewOrdersLoad(Order.All());
            cbbStatusLoad();
        }

        private void listViewOrdersLoad(List<Order> list)
        {
            int count;
            listViewOrders.Items.Clear();
            foreach (Order item in list)
            {
                OrderLoad(item);
            }
            count = listViewOrders.Items.Count;
            for (int i = 0; i < count; i++)
            {
                if (i % 2 != 0)
                    listViewOrders.Items[i].BackColor = Color.PowderBlue;
            }
            lblCoutOfOrders.Text = count + "/" + Order.Count();
        }

        private void OrderLoad(Order order)
        {
            ListViewItem item = listViewOrders.Items.Add(order.OrderID.ToString());
            if (order.Customer != null)
                item.SubItems.Add(order.Customer);
            else
                item.SubItems.Add(order.CustomerID);
            if (order.Employee != null)
                item.SubItems.Add(order.Employee);
            else
                item.SubItems.Add(order.EmployeeID.ToString());

            item.SubItems.Add(order.OrderDate.ToShortDateString());
            item.SubItems.Add(order.RequiredDate.ToShortDateString());
            item.SubItems.Add(order.ShippedDate.ToShortDateString());
            item.SubItems.Add(order.Freight.ToString());
            item.SubItems.Add(order.ShipAddress);
            item.SubItems.Add(order.ShipPostalCode);
            item.SubItems.Add(order.Status);
        }

        private void listViewOrders_Click(object sender, EventArgs e)
        {
            if (listViewOrders.SelectedItems.Count > 0)
                listViewOrders.SelectedItems[0].Checked = !listViewOrders.SelectedItems[0].Checked;
        }

        private void listViewOrders_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listViewOrders.SelectedItems.Count > 0)
            {
                treeViewOrderLoad(listViewOrders.SelectedItems[0]);
            }
        }

        private void treeViewOrderLoad(ListViewItem itemlv)
        {
            treeViewOrder.Nodes.Clear();
            List<OrderDetail> list = OrderDetail.FindByOrderID(MyConvert.ToInt32(itemlv.Text));
            foreach (var item in list)
            {
                TreeNode node = treeViewOrder.Nodes.Add("Product: " + item.Product);
                node.Nodes.Add("Unit Price: " + item.UnitPrice.ToString());
                node.Nodes.Add("Quantity: " + item.Quantity.ToString());
                node.Nodes.Add("Discount: " + item.Discount.ToString());
                node.Expand();
            }
        }

        private void toolStripButtonFindByCustomer_Click(object sender, EventArgs e)
        {
            listViewOrdersLoad(Order.FindByCustomer(FindByCustomer.Text));
        }

        private void toolStripButtonAllOrders_Click(object sender, EventArgs e)
        {
            listViewOrdersLoad(Order.All());
        }

        private void cbbStatusLoad()
        {
            List<Order> list = Order.Statuses();
            cbbStatus.Items.Clear();
            foreach (Order item in list)
            {
                cbbStatus.Items.Add(item.Status);
            }
            cbbStatus.Items.Insert(0, "-- All --");
            cbbStatus.SelectedIndex = 0;
        }

        private void cbbStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbbStatus.SelectedIndex > 0)
            {
                listViewOrdersLoad(Order.FindByStatus(cbbStatus.SelectedItem.ToString()));
            }
            else
                listViewOrdersLoad(Order.All());
        }

        private void toolStripButtonAddNewOrder_Click(object sender, EventArgs e)
        {
            OrderDetailForm frm = new OrderDetailForm();
            frm.ShowDialog();
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                listViewOrdersLoad(Order.All());
            listViewCustomersLoad(Customer.All(0));
            listViewEmployeesLoad(Employee.All());
        }

        private void toolStripButtonUpdateOrder_Click(object sender, EventArgs e)
        {
            if (listViewOrders.SelectedItems.Count > 0)
            {
                OrderDetailForm frm = new OrderDetailForm();
                frm.OrderID = MyConvert.ToInt32(listViewOrders.SelectedItems[0].SubItems[0].Text);
                frm.ShowDialog();
                if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                    listViewOrdersLoad(Order.All());
                listViewCustomersLoad(Customer.All(0));
                listViewEmployeesLoad(Employee.All());
            }
        }

        private void toolStripButtonDeleteOrder_Click(object sender, EventArgs e)
        {
            if (listViewOrders.CheckedItems.Count > 0)
            {
                DialogResult dlg = MessageBox.Show("Do you want to delete the all order which was checked?", "Warrning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dlg == System.Windows.Forms.DialogResult.OK)
                {
                    foreach (ListViewItem item in listViewOrders.CheckedItems)
                    {
                        Order.Delete(Convert.ToInt32(item.SubItems[0].Text));
                    }
                    listViewOrdersLoad(Order.All());
                    cbbStatusLoad();
                }
            }
        }

        private void toolStripButtonFindOrders_Click(object sender, EventArgs e)
        {
            FindForm frm = new FindForm();
            frm.cbbColumnLoad(Order.ColumnNames());
            frm.ShowDialog();
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                listViewOrdersLoad(Order.Find(frm.Expression));
        }

        private void toolStripButtonSortOrders_Click(object sender, EventArgs e)
        {
            SortForm frm = new SortForm();
            frm.cklbColumnsLoad(Order.ColumnNames());
            frm.ShowDialog();
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
                listViewOrdersLoad(Order.Sort(frm.Expression));
        }

        #endregion Orders
        
    }
}
