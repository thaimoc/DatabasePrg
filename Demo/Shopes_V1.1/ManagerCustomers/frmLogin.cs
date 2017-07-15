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
    public partial class frmLogin : Form
    {
        public string username;
        public string password;
        public frmLogin()
        {
            InitializeComponent();
            txtUsername.Select();
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            string un = txtUsername.Text.Trim();
            string pw = txtPassword.Text.Trim();
            int kt = Login.UserLogin(un, pw);
            try
            {
                if (un == "")
                {
                    MessageBox.Show("Thiếu tên đăng nhập!", "Chú ý!");
                    txtUsername.Select();
                    return;
                }
                if (pw == "")
                {
                    MessageBox.Show("Thiếu mật khẩu đăng nhập!", "Chú ý!");
                    txtPassword.Select();
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            if(kt > 0)
            {
                username = un;
                password = pw;
                this.Visible = false;
                Form1 f1 = new Form1();
                f1.Show();
            }
            else
            {
                MessageBox.Show("Đăng nhập thất bại !");
                txtUsername.Select();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
