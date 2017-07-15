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
    public partial class FindForm : Form
    {
        private string expression = "";

        public string Expression
        {
            get { return expression; }
        }
        public FindForm()
        {
            InitializeComponent();
        }

        private void FindForm_Load(object sender, EventArgs e)
        {
            cbbperator.SelectedIndex = 0;
            RadiosCheckState();
        }

        private void RadiosCheckState()
        {
            grbRelationExpression.Enabled = (lv.Items.Count > 0);
        }

        public void cbbColumnLoad(object[] items)
        {
            cbbColumns.Items.Clear();
            cbbColumns.Items.AddRange(items);
            if (items.Length > 0)
                cbbColumns.SelectedIndex = 0;
        }

        private void txtValue_TextChanged(object sender, EventArgs e)
        {
            btnAdd.Enabled = txtValue.Text.Length > 0;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            string andor = "";
            if (grbRelationExpression.Enabled)
            {
                if (radAnd.Checked) andor = "AND";
                else andor = "OR";
            }

            ListViewItem item = new ListViewItem(andor);
            item.SubItems.Add(cbbColumns.Text + ' ' + cbbperator.Text + ' ' + txtValue.Text);
            lv.Items.Add(item);
            this.RadiosCheckState();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            if (lv.Items.Count > 0)
            {
                StringBuilder sb = new StringBuilder();

                foreach (ListViewItem item in lv.Items)
                {
                    sb.Append(item.Text + ' ' + item.SubItems[1].Text + ' ');
                }
                expression = sb.ToString();
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
            }
            this.Close();
        }
    }
}
