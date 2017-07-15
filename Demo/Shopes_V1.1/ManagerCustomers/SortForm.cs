using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ManagerCustomers
{
    public partial class SortForm : Form
    {
        private string expression = "";

        public string Expression
        {
            get { return expression; }
        }

        public SortForm()
        {
            InitializeComponent();
        }

        private void SortForm_Load(object sender, EventArgs e)
        {

        }

        public void cklbColumnsLoad(object[] items)
        {
            cklbColumns.Items.Clear();
            cklbColumns.Items.AddRange(items);
        }

        private void cklbColumns_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (e.NewValue == CheckState.Unchecked)
            {
                lblExpressionCurrent.Text = "";
                foreach (var item in cklbColumns.CheckedItems)
                {
                    if (cklbColumns.Items[e.Index].ToString() != item.ToString())
                        lblExpressionCurrent.Text += item.ToString() + ", ";
                }
                if (lblExpressionCurrent.Text.Length > 0)
                    lblExpressionCurrent.Text = lblExpressionCurrent.Text.Substring(0, lblExpressionCurrent.Text.Length - 2);
            }
            else
            {
                if (lblExpressionCurrent.Text.Length > 0)
                    lblExpressionCurrent.Text = lblExpressionCurrent.Text + ", "
                    + cklbColumns.Items[e.Index].ToString();
                else
                    lblExpressionCurrent.Text = cklbColumns.Items[e.Index].ToString();
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (cklbColumns.Items.Count > 0)
            {
                string orderby = "ASC";
                if (radDESC.Checked) orderby = "DESC";
                lvExpressions.Items.Add(lblExpressionCurrent.Text + ' ' + orderby);
                for (int i = 0; i < cklbColumns.Items.Count; i++)
                {
                    if (cklbColumns.GetItemChecked(i))
                        cklbColumns.Items.RemoveAt(i--);
                }
                lblExpressionCurrent.Text = "";
            }
            else
            {
                grbSortType.Enabled = false;
            }
            btnAdd.Enabled = false;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            if (lvExpressions.Items.Count > 0)
            {
                expression = "";
                foreach (ListViewItem item in lvExpressions.Items)
                {
                    expression += item.SubItems[0].Text + ", ";
                }
                expression = expression.Substring(0, expression.Length - 2);
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
            }
            this.Close();
        }

        private void lvExpressions_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvExpressions.SelectedItems.Count > 0)
                btnRemove.Enabled = true;
            else
                btnRemove.Enabled = false;
        }

        private void btnRemove_Click(object sender, EventArgs e)
        {
            if (lvExpressions.SelectedItems.Count > 0)
            {
                lvExpressions.SelectedItems[0].Text = lvExpressions.SelectedItems[0].Text.Substring(0, lvExpressions.SelectedItems[0].Text.LastIndexOf(' '));
                string[] s = lvExpressions.SelectedItems[0].Text.Split(',');
                cklbColumns.Items.Add(s[0]);
                for (int i = 1; i < s.Length; ++i)
                {
                    cklbColumns.Items.Add(s[i].Substring(1));
                }
                lvExpressions.Items.Remove(lvExpressions.SelectedItems[0]);
            }
            btnRemove.Enabled = false;
        }

        private void cklbColumns_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cklbColumns.CheckedItems.Count > 0)
                btnAdd.Enabled = true;
            else
                btnAdd.Enabled = false;
        }
    }
}
