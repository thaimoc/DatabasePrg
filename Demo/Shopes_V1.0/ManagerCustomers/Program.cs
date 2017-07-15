using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DataAccess;

namespace ManagerCustomers
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new ConnectForm());
            ConnectForm frm = new ConnectForm();
            frm.ShowDialog();
            if (frm.DialogResult == DialogResult.OK)
            {
                new Form1().ShowDialog();
            }
        }
    }
}
