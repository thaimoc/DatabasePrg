using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Linq;
using System.Text;
using System.Windows.Forms;
using DataAccess;
using Microsoft.Office.Interop.Excel;
using System.Threading;

namespace ManagerCustomers
{
    public partial class Form1 : Form
    {
        public ListViewColumnSorter lvwColumnSorter;
        public Form1()
        {
            InitializeComponent();
            lvwColumnSorter = new ListViewColumnSorter();
            this.listViewCustomers.ListViewItemSorter = lvwColumnSorter;
            this.listViewOrders.ListViewItemSorter = lvwColumnSorter;
            this.listViewSuppliers.ListViewItemSorter = lvwColumnSorter;
            this.listViewProducts.ListViewItemSorter = lvwColumnSorter;
            this.listViewEmployees.ListViewItemSorter = lvwColumnSorter;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            menuLogin.Enabled = false;
            tabControl.TabPages.Clear();
            tabControl.TabPages.Add(tabPage1);
            //tabControl.TabPages.Add(tabPageEmployees);
            //tabControl.TabPages.Add(tabPageOrders);
            //tabControl.TabPages.Add(tabPageProducts);
            //tabControl.TabPages.Add(tabPageSuppliers);
            //tabControl.TabPages.Add(tpCustomers);
            panelCustomersLoand();
            tabPageSuppliersLoad();
            tabPageProductsLoad();
            tapageEmployeesLoad();
            tabPageOrdersLoad();
        }

        //int i = 5;

        #region Customers
        private void panelCustomersLoand()
        {
            listViewCustomersLoad(Customer.All(0));
            // cbbCountryLoad();
            //  timerCustomerSupport.Start();
            // cbbColumnCustomersLoad(Customer.ColumnNames());

        }
        //private void cbbColumnCustomersLoad(object[] items)
        //{
        //    toolStripCbbCustomerFind.Items.Clear();
        //    toolStripCbbCustomerFind.Items.AddRange(items);
        //    if (items.Length > 0)
        //        toolStripCbbCustomerFind.SelectedIndex = 0;
        //}
        private void listViewCustomersLoad(List<Customer> list)
        {
            int count;
            listViewCustomers.Items.Clear();
            foreach (Customer item in list)
            {
                customerLoad(item);
            }
            count = listViewCustomers.Items.Count;
            //for (int i = 0; i < count; i++)
            //{
            //    if (i % 2 != 0)
            //    {
            //        listViewCustomers.Items[i].BackColor = Color.PowderBlue;
            //    }
            //}
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
            if (listViewCustomers.SelectedItems.Count > 0)
            {
                DialogResult dlg = MessageBox.Show("Do you want to delete the all customer whitch you select?", "Message Box", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (dlg == System.Windows.Forms.DialogResult.OK)
                {
                    List<Customer> list = new List<Customer>();
                    foreach (ListViewItem item in listViewCustomers.Items)
                    {
                        if (item.Selected)
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

        //private void cbbCountryLoad()
        //{
        //    cbbCountry.Items.Clear();
        //    List<Customer> list = Customer.Countries();
        //    foreach (var item in list)
        //    {
        //        cbbCountry.Items.Add(item.Country);
        //    }
        //    cbbCountry.Items.Insert(0, "--- ALL ---");
        //    cbbCountry.SelectedIndex = 0;
        //}

        //private void cbbCountry_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    if (cbbCountry.SelectedIndex > 0)
        //    {
        //        listViewCustomersLoad(Customer.FindByCountry(cbbCountry.SelectedItem.ToString()));
        //    }
        //    else
        //        toolStripButtonAll.PerformClick();
        //}

        private void listViewListProductLoad()
        {
            if (listViewCustomers.SelectedItems.Count > 0)
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
            string Value = toolStripTxtCustomersFind.Text;
            string CbbValue = toolStripCbbCustomerFind.Text;
            List<Customer> list = Customer.FindBy(Value, CbbValue);
            listViewCustomersLoad(list);
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
            //for (int i = 0; i < count; i++)
            //{
            //    if (i % 2 != 0)
            //    {
            //        listViewSuppliers.Items[i].BackColor = Color.PowderBlue;
            //    }
            //}
            toolStripLabel4lblCountOfSuppliers.Text = count + "/" + Supplier.Count();
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
            listViewSuppliersLoad(Supplier.FindByLikeCompany(toolStripTextBoxSupplersFind.Text));
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

        private void toolStripButtonDeleteSuppliers_Click(object sender, EventArgs e)
        {
            if (listViewSuppliers.SelectedItems.Count > 0)
            {
                DialogResult dlg = MessageBox.Show("Do you want to delete the all suppliers whitch you select?", "Message Box", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (dlg == System.Windows.Forms.DialogResult.OK)
                {
                    List<Supplier> list = new List<Supplier>();
                    foreach (ListViewItem item in listViewSuppliers.Items)
                    {
                        if (item.Selected)
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

        }
        private void toolStripButtonSuppliersFind_Click(object sender, EventArgs e)
        {
            string value = toolStripTextBoxSupplersFind.Text;
            string cbbvalue = toolStripCbbSuppliersFind.Text;
            List<Supplier> list = Supplier.FindBy(value, cbbvalue);
            listViewSuppliersLoad(list);
        }
        #endregion Suppliers

        #region Products

        private void tabPageProductsLoad()
        {
            listViewProductsLoad(Product.All());
        }
        public void cbbColumnProductsLoad(object[] items)
        {
            TsCbbProductFind.Items.Clear();
            TsCbbProductFind.Items.AddRange(items);
            if (items.Length > 0)
                TsCbbProductFind.SelectedIndex = 0;
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
            //for (int i = 0; i < count; i++)
            //{
            //    if (i % 2 != 0)
            //    {
            //        listViewProducts.Items[i].BackColor = Color.PowderBlue;
            //    }
            //}
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
            listViewProductsLoad(Product.FindByLikeProductName(tsTxtProductFind.Text));
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
            if (listViewProducts.SelectedItems.Count > 0)
            {
                DialogResult dlg = MessageBox.Show("Do you want to delete the all products which were had been selected?", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dlg == System.Windows.Forms.DialogResult.OK)
                {
                    foreach (ListViewItem item in listViewProducts.SelectedItems)
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
            string Value = tsTxtProductFind.Text;
            string ValueCbbox = TsCbbProductFind.Text;
            List<Product> list = Product.FindBy(Value, ValueCbbox);
            listViewProductsLoad(list);
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
            //for (int i = 0; i < count; i++)
            //{
            //    if (i % 2 == 0)
            //        listViewEmployees.Items[i].ForeColor = Color.Chocolate;
            //}
            lblCountOfEmployee.Text = count + "/" + Order.Count();
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
            listViewEmployeesLoad(Employee.FindByName(toolStripTextBoxEmployeeFind.Text));
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
            if (listViewEmployees.SelectedItems.Count > 0)
            {
                DialogResult dlg = MessageBox.Show("Do you want to delete the employees which was been checked!", "Quession", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dlg == System.Windows.Forms.DialogResult.OK)
                {
                    foreach (ListViewItem item in listViewEmployees.SelectedItems)
                    {
                        Employee.Delete(MyConvert.ToInt32(item.SubItems[0].Text));
                    }
                    listViewEmployeesLoad(Employee.All());
                }
            }
        }

        private void toolStripButtonFindEmployees_Click(object sender, EventArgs e)
        {
            string value = toolStripTextBoxEmployeeFind.Text;
            string cbbvalue = toolStripcbbEmployeeFind.Text;
            List<Employee> list = Employee.FindBy(value, cbbvalue);
            listViewEmployeesLoad(list);
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
            //  cbbStatusLoad();
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
            //for (int i = 0; i < count; i++)
            //{
            //    if (i % 2 != 0)
            //        tsbOrExportToExcel.Items[i].BackColor = Color.PowderBlue;
            //}
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
            listViewOrdersLoad(Order.FindByCustomer(toolStripTxtOrdersFind.Text));
        }

        private void toolStripButtonAllOrders_Click(object sender, EventArgs e)
        {
            listViewOrdersLoad(Order.All());
        }

        //private void cbbStatusLoad()
        //{
        //    List<Order> list = Order.Statuses();
        //    cbbStatus.Items.Clear();
        //    foreach (Order item in list)
        //    {
        //        cbbStatus.Items.Add(item.Status);
        //    }
        //    cbbStatus.Items.Insert(0, "-- All --");
        //    cbbStatus.SelectedIndex = 0;
        //}

        //private void cbbStatus_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    if (cbbStatus.SelectedIndex > 0)
        //    {
        //        listViewOrdersLoad(Order.FindByStatus(cbbStatus.SelectedItem.ToString()));
        //    }
        //    else
        //        listViewOrdersLoad(Order.All());
        //}

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
            if (listViewOrders.SelectedItems.Count > 0)
            {
                DialogResult dlg = MessageBox.Show("Do you want to delete the all order which was selected?", "Warrning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dlg == System.Windows.Forms.DialogResult.OK)
                {
                    foreach (ListViewItem item in listViewOrders.SelectedItems)
                    {
                        Order.Delete(Convert.ToInt32(item.SubItems[0].Text));
                    }
                    listViewOrdersLoad(Order.All());
                    //  cbbStatusLoad();
                }
            }
        }

        private void toolStripButtonFindOrders_Click(object sender, EventArgs e)
        {
            string Value = toolStripTxtOrdersFind.Text;
            string CbbValue = toolStripCbbOrdersFind.Text;
            List<Order> list = Order.FindBy(Value, CbbValue);
            listViewOrdersLoad(list);
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

        #region SapXep
        private void listViewCustomers_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            for (int i = 0; i < listViewCustomers.Columns.Count; i++)
            {
                int x = listViewCustomers.Columns[i].ImageIndex;
                if (x == 0 || x == 1)
                {
                    listViewCustomers.Columns[i].ImageIndex = 2;
                }

            }
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order == SortOrder.Ascending)
                {
                    listViewCustomers.Columns[e.Column].ImageIndex = 1;
                    lvwColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    listViewCustomers.Columns[e.Column].ImageIndex = 0;
                    lvwColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
                listViewCustomers.Columns[e.Column].ImageIndex = 0;
            }
            // Perform the sort with these new sort options.
            this.listViewCustomers.Sort();
        }

        private void listViewOrders_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            for (int i = 0; i < listViewOrders.Columns.Count; i++)
            {
                int x = listViewOrders.Columns[i].ImageIndex;
                if (x == 0 || x == 1)
                {
                    listViewOrders.Columns[i].ImageIndex = 2;
                }

            }
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order == SortOrder.Ascending)
                {
                    listViewOrders.Columns[e.Column].ImageIndex = 1;
                    lvwColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    listViewOrders.Columns[e.Column].ImageIndex = 0;
                    lvwColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
                listViewOrders.Columns[e.Column].ImageIndex = 0;
            }
            // Perform the sort with these new sort options.
            this.listViewOrders.Sort();
        }

        private void listViewSuppliers_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            for (int i = 0; i < listViewSuppliers.Columns.Count; i++)
            {
                int x = listViewSuppliers.Columns[i].ImageIndex;
                if (x == 0 || x == 1)
                {
                    listViewSuppliers.Columns[i].ImageIndex = 2;
                }

            }
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order == SortOrder.Ascending)
                {
                    listViewSuppliers.Columns[e.Column].ImageIndex = 1;
                    lvwColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    listViewSuppliers.Columns[e.Column].ImageIndex = 0;
                    lvwColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
                listViewSuppliers.Columns[e.Column].ImageIndex = 0;
            }
            // Perform the sort with these new sort options.
            this.listViewSuppliers.Sort();
        }

        private void listViewProducts_ColumnClick(object sender, ColumnClickEventArgs e)
        {

            for (int i = 0; i < listViewProducts.Columns.Count; i++)
            {
                int x = listViewProducts.Columns[i].ImageIndex;
                if (x == 0 || x == 1)
                {
                    listViewProducts.Columns[i].ImageIndex = 2;
                }

            }
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order == SortOrder.Ascending)
                {
                    listViewProducts.Columns[e.Column].ImageIndex = 1;
                    lvwColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    listViewProducts.Columns[e.Column].ImageIndex = 0;
                    lvwColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
                listViewProducts.Columns[e.Column].ImageIndex = 0;
            }
            // Perform the sort with these new sort options.
            this.listViewProducts.Sort();
        }

        private void listViewEmployees_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            for (int i = 0; i < listViewEmployees.Columns.Count; i++)
            {
                int x = listViewEmployees.Columns[i].ImageIndex;
                if (x == 0 || x == 1)
                {
                    listViewEmployees.Columns[i].ImageIndex = 2;
                }

            }
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order == SortOrder.Ascending)
                {
                    listViewEmployees.Columns[e.Column].ImageIndex = 1;
                    lvwColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    listViewEmployees.Columns[e.Column].ImageIndex = 0;
                    lvwColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
                listViewEmployees.Columns[e.Column].ImageIndex = 0;
            }
            // Perform the sort with these new sort options.
            this.listViewEmployees.Sort();
        }
        #endregion





        private void tabControl_MouseDown(object sender, MouseEventArgs e)
        {
            for (int i = 0; i < this.tabControl.TabPages.Count; i++)
            {
                System.Drawing.Rectangle rPage = tabControl.GetTabRect(i);
                System.Drawing.Rectangle closeButton = new System.Drawing.Rectangle(rPage.Right - 20, rPage.Top + 15, 20, 20);
                if (closeButton.Contains(e.Location))
                {
                    if (MessageBox.Show("Close this Tab?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        this.tabControl.TabPages.RemoveAt(i);
                        break;
                    }
                }
            }
        }

        private void menuLogin_Click(object sender, EventArgs e)
        {
            frmLogin login = new frmLogin();
            login.ShowDialog();


        }

        private void menuLogout_Click(object sender, EventArgs e)
        {
            this.Close();
            MessageBox.Show("Bạn có chắc chắn Logout không ?", "Hỏi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            frmLogin login = new frmLogin();
            login.ShowDialog();

        }

        private void menuExit_Click_1(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có chắc chắn Thoát không ?", "Hỏi", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                System.Windows.Forms.Application.Exit();
        }

        private void menuAbout_Click(object sender, EventArgs e)
        {
            if (!tabControl.TabPages.Contains(tabPageAboutUs))
                tabControl.TabPages.Add(tabPageAboutUs);
            else
                tabControl.SelectTab(tabPageAboutUs);
        }
        private void menuCustomers_Click(object sender, EventArgs e)
        {
            if (!tabControl.TabPages.Contains(tpCustomers))
                tabControl.TabPages.Add(tpCustomers);
            else
                tabControl.SelectTab(tpCustomers);
        }
        private void menuProducts_Click(object sender, EventArgs e)
        {
            if (!tabControl.TabPages.Contains(tabPageProducts))
                tabControl.TabPages.Add(tabPageProducts);
            else
                tabControl.SelectTab(tabPageProducts);
        }

        private void menuEmployees_Click(object sender, EventArgs e)
        {
            if (!tabControl.TabPages.Contains(tabPageEmployees))
                tabControl.TabPages.Add(tabPageEmployees);
            else
                tabControl.SelectTab(tabPageEmployees);
        }

        private void menuOrders_Click(object sender, EventArgs e)
        {
            if (!tabControl.TabPages.Contains(tabPageOrders))
                tabControl.TabPages.Add(tabPageOrders);
            else
                tabControl.SelectTab(tabPageOrders);
        }

        private void menuSuppliers_Click(object sender, EventArgs e)
        {
            if (!tabControl.TabPages.Contains(tabPageSuppliers))
                tabControl.TabPages.Add(tabPageSuppliers);
            else
                tabControl.SelectTab(tabPageSuppliers);
        }
        private void menuHome_Click(object sender, EventArgs e)
        {
            if (!tabControl.TabPages.Contains(tabPage1))
                tabControl.TabPages.Add(tabPage1);
            else
                tabControl.SelectTab(tabPage1);
        }
        private void MenuStripOrdersDetail_Click(object sender, EventArgs e)
        {
            OrderDetailForm od = new OrderDetailForm();
            od.ShowDialog();
        }

        //==========================================================================
  

        private static void OpenFile(out bool hasHeaders, out OpenFileDialog open)
        {
            hasHeaders = true;
            open = new OpenFileDialog();
            open.FileName = "Import file excel";
            open.Filter = "Excel|*.xls";
            open.Multiselect = false;
        }
        //////////////////////////////IMPORT - EXPORT Employees////////////////////////////////////////////
        private void DataBinEmployees(List<Employee> list)
        {
            listViewEmployees.Items.Clear();
            foreach (Employee item in list)
            {
                ListViewItem lvitem = new ListViewItem(item.EmployeeID.ToString());
                lvitem.SubItems.Add(item.EmployeeName.ToString());
                lvitem.SubItems.Add(item.BirthDate.ToShortDateString());
                lvitem.SubItems.Add(item.HireDate.ToShortDateString());
                lvitem.SubItems.Add(item.Address);
                lvitem.SubItems.Add(item.PostalCode.ToString());
                lvitem.SubItems.Add(item.Country.ToString());
                lvitem.SubItems.Add(item.HomePhone.ToString());
                listViewEmployees.Items.Add(lvitem);
            }
        }
        private void tsbExportToExcel_Click(object sender, EventArgs e)
        {
            Worksheet ws = Import_Export.ExportExcel();
            ws.Cells[1, 1] = "EmployeeID";
            ws.Cells[1, 2] = "Name";
            ws.Cells[1, 3] = "BirthDate";
            ws.Cells[1, 4] = "HireDate";
            ws.Cells[1, 5] = "Address";
            ws.Cells[1, 6] = "PostalCode";
            ws.Cells[1, 7] = "Country";
            ws.Cells[1, 8] = "HomePhone";
            for (int i = 0; i < listViewEmployees.Items.Count; i++)
            {
                ListViewItem lvitem = listViewEmployees.Items[i];
                ws.Cells[i + 2, 1] = lvitem.SubItems[0].Text;
                ws.Cells[i + 2, 2] = lvitem.SubItems[1].Text;
                ws.Cells[i + 2, 3] = lvitem.SubItems[2].Text;
                ws.Cells[i + 2, 4] = lvitem.SubItems[3].Text;
                ws.Cells[i + 2, 5] = lvitem.SubItems[4].Text;
                ws.Cells[i + 2, 6] = lvitem.SubItems[5].Text;
                ws.Cells[i + 2, 7] = lvitem.SubItems[6].Text;
                ws.Cells[i + 2, 8] = lvitem.SubItems[7].Text;
            }
        }
        private void tsbImportToExcel_Click(object sender, EventArgs e)
        {
            bool hasHeaders = true;
            OpenFileDialog open = new OpenFileDialog();
            open.FileName = "Import file excel";
            open.Filter = "Excel|*.xls";
            open.Multiselect = false;
            if (open.ShowDialog() == DialogResult.OK)
            {
                //List<Software> list = Import_Export.ImportExcel(openFileDialog1.FileName, hasHeaders);
                DataBinEmployees(Import_Export.ImportExcel_Employee(open.FileName, hasHeaders));
            }
        }
        //////////////////////////////IMPORT - EXPORT Customers////////////////////////////////////////////
        private void DataBinCustomers(List<Customer> list)
        {
            listViewCustomers.Items.Clear();
            foreach (Customer item in list)
            {
                ListViewItem lvitem = new ListViewItem(item.CustomerID.ToString());
                lvitem.SubItems.Add(item.CompanyName.ToString());
                lvitem.SubItems.Add(item.ContactName);
                lvitem.SubItems.Add(item.Address);
                lvitem.SubItems.Add(item.PostalCode);
                lvitem.SubItems.Add(item.Country.ToString());
                lvitem.SubItems.Add(item.Phone.ToString());
                lvitem.SubItems.Add(item.Fax.ToString());
                listViewCustomers.Items.Add(lvitem);
            }
        }
        private void tsbCusExportToExcel_Click(object sender, EventArgs e)
        {
            Worksheet ws = Import_Export.ExportExcel();
            ws.Cells[1, 1] = "ID";
            ws.Cells[1, 2] = "CompanyName";
            ws.Cells[1, 3] = "ContactName";
            ws.Cells[1, 4] = "Address";
            ws.Cells[1, 5] = "PostalCode";
            ws.Cells[1, 6] = "Country";
            ws.Cells[1, 7] = "Phone";
            ws.Cells[1, 8] = "Fax";
            for (int i = 0; i < listViewCustomers.Items.Count; i++)
            {
                ListViewItem lvitem = listViewCustomers.Items[i];
                ws.Cells[i + 2, 1] = lvitem.SubItems[0].Text;
                ws.Cells[i + 2, 2] = lvitem.SubItems[1].Text;
                ws.Cells[i + 2, 3] = lvitem.SubItems[2].Text;
                ws.Cells[i + 2, 4] = lvitem.SubItems[3].Text;
                ws.Cells[i + 2, 5] = lvitem.SubItems[4].Text;
                ws.Cells[i + 2, 6] = lvitem.SubItems[5].Text;
                ws.Cells[i + 2, 7] = lvitem.SubItems[6].Text;
                ws.Cells[i + 2, 8] = lvitem.SubItems[7].Text;
            }
        }
        private void tsbCusImportToExcel_Click(object sender, EventArgs e)
        {
            bool hasHeaders;
            OpenFileDialog open;
            OpenFile(out hasHeaders, out open);
            if (open.ShowDialog() == DialogResult.OK)
            {
                DataBinCustomers(Import_Export.ImportExcel_Customer(open.FileName, hasHeaders));
                // DataBinCustomers(Import_Export.ImportExcel(open.FileName, hasHeaders));
            }
        }
        //////////////////////////////IMPORT - EXPORT Order////////////////////////////////////////////
        private void DataBinOrder(List<Order> list)
        {
            listViewOrders.Items.Clear();
            foreach (Order item in list)
            {
                ListViewItem lvitem = new ListViewItem(item.OrderID.ToString());
                lvitem.SubItems.Add(item.Customer);
                lvitem.SubItems.Add(item.Employee);
                lvitem.SubItems.Add(item.OrderDate.ToShortDateString());
                lvitem.SubItems.Add(item.RequiredDate.ToShortDateString());
                lvitem.SubItems.Add(item.ShippedDate.ToShortDateString());
                lvitem.SubItems.Add(item.Freight.ToString());
                lvitem.SubItems.Add(item.ShipAddress.ToString());
                lvitem.SubItems.Add(item.ShipPostalCode.ToString());
                lvitem.SubItems.Add(item.Status.ToString());
                listViewOrders.Items.Add(lvitem);
            }
        }
        private void tsbOrExportToExcel_Click(object sender, EventArgs e)
        {
            Worksheet ws = Import_Export.ExportExcel();
            ws.Cells[1, 1] = "OrderID";
            ws.Cells[1, 2] = "CustomerID";
            ws.Cells[1, 3] = "EmployeeID";
            ws.Cells[1, 4] = "OrderDate";
            ws.Cells[1, 5] = "RequiredDate";
            ws.Cells[1, 6] = "ShippedDate";
            ws.Cells[1, 7] = "Freight";
            ws.Cells[1, 8] = "ShipAddress";
            ws.Cells[1, 9] = "ShipPostalCode";
            ws.Cells[1, 10] = "Status";
            for (int i = 0; i < listViewOrders.Items.Count; i++)
            {
                ListViewItem lvitem = listViewOrders.Items[i];
                ws.Cells[i + 2, 1] = lvitem.SubItems[0].Text;
                ws.Cells[i + 2, 2] = lvitem.SubItems[1].Text;
                ws.Cells[i + 2, 3] = lvitem.SubItems[2].Text;
                ws.Cells[i + 2, 4] = lvitem.SubItems[3].Text;
                ws.Cells[i + 2, 5] = lvitem.SubItems[4].Text;
                ws.Cells[i + 2, 6] = lvitem.SubItems[5].Text;
                ws.Cells[i + 2, 7] = lvitem.SubItems[6].Text;
                ws.Cells[i + 2, 8] = lvitem.SubItems[7].Text;
                ws.Cells[i + 2, 9] = lvitem.SubItems[8].Text;
                ws.Cells[i + 2, 10] = lvitem.SubItems[9].Text;
            }
        }
        private void tsbOrImportToExcel_Click(object sender, EventArgs e)
        {
            bool hasHeaders;
            OpenFileDialog open;
            OpenFile(out hasHeaders, out open);
            if (open.ShowDialog() == DialogResult.OK)
            {
                DataBinOrder(Import_Export.ImportExcel_Order(open.FileName, hasHeaders));
                // DataBinCustomers(Import_Export.ImportExcel(open.FileName, hasHeaders));
            }
        }
        ///////////////////////////////IMPORT - EXPORT Supplier//////////////////////////////////////
        private void DataBinSupplier(List<Supplier> list)
        {
            listViewSuppliers.Items.Clear();
            foreach (Supplier item in list)
            {
                ListViewItem lvitem = new ListViewItem(item.SupplierID.ToString());
                lvitem.SubItems.Add(item.CompanyName);
                lvitem.SubItems.Add(item.ContactName);
                lvitem.SubItems.Add(item.ContactTitle);
                lvitem.SubItems.Add(item.Address);
                lvitem.SubItems.Add(item.City);
                lvitem.SubItems.Add(item.Region);
                lvitem.SubItems.Add(item.PostalCode);
                lvitem.SubItems.Add(item.Country);
                lvitem.SubItems.Add(item.Phone);
                lvitem.SubItems.Add(item.Fax);
                lvitem.SubItems.Add(item.HomePage);
                listViewSuppliers.Items.Add(lvitem);
            }
        }
        private void tsbSupExportToExcel_Click(object sender, EventArgs e)
        {
            Worksheet ws = Import_Export.ExportExcel();
            ws.Cells[1, 1] = "SupplierID";
            ws.Cells[1, 2] = "CompanyName";
            ws.Cells[1, 3] = "ContactName";
            ws.Cells[1, 4] = "ContactTitle";
            ws.Cells[1, 5] = "Address";
            ws.Cells[1, 6] = "City";
            ws.Cells[1, 7] = "Region";
            ws.Cells[1, 8] = "PostalCode";
            ws.Cells[1, 9] = "Country";
            ws.Cells[1, 10] = "Phone";
            ws.Cells[1, 11] = "Fax";
            ws.Cells[1, 12] = "HomePage";
            for (int i = 0; i < listViewSuppliers.Items.Count; i++)
            {
                ListViewItem lvitem = listViewSuppliers.Items[i];
                ws.Cells[i + 2, 1] = lvitem.SubItems[0].Text;
                ws.Cells[i + 2, 2] = lvitem.SubItems[1].Text;
                ws.Cells[i + 2, 3] = lvitem.SubItems[2].Text;
                ws.Cells[i + 2, 4] = lvitem.SubItems[3].Text;
                ws.Cells[i + 2, 5] = lvitem.SubItems[4].Text;
                ws.Cells[i + 2, 6] = lvitem.SubItems[5].Text;
                ws.Cells[i + 2, 7] = lvitem.SubItems[6].Text;
                ws.Cells[i + 2, 8] = lvitem.SubItems[7].Text;
                ws.Cells[i + 2, 9] = lvitem.SubItems[8].Text;
                ws.Cells[i + 2, 10] = lvitem.SubItems[9].Text;
                ws.Cells[i + 2, 11] = lvitem.SubItems[10].Text;
                ws.Cells[i + 2, 12] = lvitem.SubItems[11].Text;
            }
        }
        private void tsbSupImportToExcel_Click(object sender, EventArgs e)
        {
            bool hasHeaders;
            OpenFileDialog open;
            OpenFile(out hasHeaders, out open);
            if (open.ShowDialog() == DialogResult.OK)
            {
                DataBinSupplier(Import_Export.ImportExcel_Supplier(open.FileName, hasHeaders));
                // DataBinCustomers(Import_Export.ImportExcel(open.FileName, hasHeaders));
            }
        }
        //////////////////////////////IMPORT - EXPORT Product////////////////////////////////////////////
        private void DataBinProduct(List<Product> list)
        {
            listViewProducts.Items.Clear();
            foreach (Product item in list)
            {
                ListViewItem lvitem = new ListViewItem(item.ProductID.ToString());
                lvitem.SubItems.Add(item.ProductName);
                lvitem.SubItems.Add(item.SupplierID.ToString());
                lvitem.SubItems.Add(item.QuantityPerUnit);
                lvitem.SubItems.Add(item.UnitPrice.ToString());
                lvitem.SubItems.Add(item.UnitsInStock.ToString());
                lvitem.SubItems.Add(item.UnitsOnOrder.ToString());
                lvitem.SubItems.Add(item.Discontinued.ToString());
                listViewProducts.Items.Add(lvitem);
            }
        }
        private void tsbProExportToExcel_Click(object sender, EventArgs e)
        {
            Worksheet ws = Import_Export.ExportExcel();
            ws.Cells[1, 1] = "ProductID";
            ws.Cells[1, 2] = "ProductName";
            ws.Cells[1, 3] = "SupplierID";
            ws.Cells[1, 4] = "QuantityPerUnit";
            ws.Cells[1, 5] = "UnitPrice";
            ws.Cells[1, 6] = "UnitsInStock";
            ws.Cells[1, 7] = "UnitsOnOrder";
            ws.Cells[1, 8] = "Discontinued";
            for (int i = 0; i < listViewProducts.Items.Count; i++)
            {
                ListViewItem lvitem = listViewProducts.Items[i];
                ws.Cells[i + 2, 1] = lvitem.SubItems[0].Text;
                ws.Cells[i + 2, 2] = lvitem.SubItems[1].Text;
                ws.Cells[i + 2, 3] = lvitem.SubItems[2].Text;
                ws.Cells[i + 2, 4] = lvitem.SubItems[3].Text;
                ws.Cells[i + 2, 5] = lvitem.SubItems[4].Text;
                ws.Cells[i + 2, 6] = lvitem.SubItems[5].Text;
                ws.Cells[i + 2, 7] = lvitem.SubItems[6].Text;
                ws.Cells[i + 2, 8] = lvitem.SubItems[7].Text;
            }
        }
        private void tsbProImportToExcel_Click(object sender, EventArgs e)
        {
            bool hasHeaders;
            OpenFileDialog open;
            OpenFile(out hasHeaders, out open);
            if (open.ShowDialog() == DialogResult.OK)
            {
                DataBinProduct(Import_Export.ImportExcel_Product(open.FileName, hasHeaders));
                // DataBinCustomers(Import_Export.ImportExcel(open.FileName, hasHeaders));
            }
        }

        private void menuHide_Click(object sender, EventArgs e)
        {
            tabControl.Dock = DockStyle.Fill;
        }




        //////////////////////////////////////////////////////////////////////////

    }
}














