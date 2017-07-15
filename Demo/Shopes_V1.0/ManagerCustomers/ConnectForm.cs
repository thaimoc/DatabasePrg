using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using DataAccess;

namespace ManagerCustomers
{
    public partial class ConnectForm : Form
    {
        public ConnectForm()
        {
            InitializeComponent();
        }

        private void ConnectForm_Load(object sender, EventArgs e)
        {
            cbbAuntication.SelectedIndex = 1;
            KimNguyenReadFile();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            string s = "";
            if (cbbAuntication.SelectedItem.ToString() == "SQL Sever Authentication")
            {
                string log = "";
                string pw = "";
                try
                {
                    if (txtUser.Text.Length > 0)
                        log = txtUser.Text;
                    if (txtPassword.Text.Length > 0)
                        pw = txtPassword.Text;
                }
                catch
                { }
                if (log.Length > 0)
                    s = string.Format("Server={0};UID={1};PWD={2};Database=Shopes", txtServer.Text, log, pw);
            }
            else
            {
                s = string.Format("server={0};dataBase=Shopes;integrated security=true", txtServer.Text);
            }            
            try
            {
                Helper.connectionString = s;
                if (Helper.TestConnect())
                {
                    MessageBox.Show("Connection is successful!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    WrritenFile(s);
                    this.DialogResult = System.Windows.Forms.DialogResult.OK;
                }
                else
                {
                    MessageBox.Show("Connection is fails!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);                    
                }
            }
            catch
            {
                Application.Exit();
            }
        }

        private void txtServer_TextChanged(object sender, EventArgs e)
        {
            if (txtServer.Text.Length > 0)
            {
                btnConnect.Enabled = true;
            }
            else
                btnConnect.Enabled = false;
        }

        private void cbbAuntication_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbbAuntication.SelectedItem.ToString() == "Windows Authentication")
            {
                txtUser.Text = "sa";
                txtPassword.Text = "";
                grblogin.Enabled = false;
            }
            else
            {
                grblogin.Enabled = true;
            }
        }

        private void WrritenFile(string connectionString)
        {
            try
            {
                FileStream myfile = new FileStream(@"app.iso", FileMode.CreateNew, FileAccess.ReadWrite, FileShare.Write);
                StreamWriter tex = new StreamWriter(myfile);
                tex.WriteLine(connectionString);                
                tex.Close();
                myfile.Close();
            }
            catch
            { }            
        }

        private void KimNguyenWrritenFile()
        {
            try
            {
                FileStream myfile = new FileStream(@"Login.iso", FileMode.CreateNew, FileAccess.ReadWrite, FileShare.Write);
                StreamWriter tex = new StreamWriter(myfile);
                tex.WriteLine(txtUser.Text);
                tex.WriteLine(txtPassword.Text);
                tex.Close();
                myfile.Close();
            }
            catch
            { }
        }
        private void KimNguyenDeleteFile()
        {
            try
            {
                File.Delete(@"Login.iso");
            }
            catch { }
        }
        private void KimNguyenReadFile()
        {
            try
            {
                FileStream myfile = new FileStream(@"Login.iso", FileMode.Open, FileAccess.ReadWrite, FileShare.Write);
                StreamReader txt = new StreamReader(myfile);
                txtUser.Text = txt.ReadLine();
                txtPassword.Text = txt.ReadLine();
                txt.Close();
                myfile.Close();
                ckbRemember.Checked = true;
            }
            catch
            {
                ckbRemember.Checked = false;
            }

            try
            {
                FileStream myfile = new FileStream(@"app.iso", FileMode.Open, FileAccess.ReadWrite, FileShare.Write);
                StreamReader txt = new StreamReader(myfile);
                string s = txt.ReadLine();
                string[] ss = s.Split(';');
                string[] sss = ss[0].Split('=');
                txtServer.Text = sss[1];
            }
            catch { }

        }
        private void ConnectForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ckbRemember.Checked)
                KimNguyenWrritenFile();
            else
                KimNguyenDeleteFile();
        }


    }
}
