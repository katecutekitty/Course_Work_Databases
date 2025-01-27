using cursovaya;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BDKursovaya
{
    public partial class AuthForm : Form
    {

        //public static string s = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BeautySalon2.mdb;";
        //private OleDbConnection mC;

        //public AuthForm()
        //{
        //    InitializeComponent();

        //    mC = new OleDbConnection(s);
        //    mC.Open();

        public static string s = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BeautySalon2.mdb;";
        private OleDbConnection mC;

        public AuthForm()
        {
            InitializeComponent();
            mC = new OleDbConnection(s);
            mC.Open();
        }

        public string returnLogin()
        {
            return loginBox.Text.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string login = loginBox.Text;
            string password = passwordBox.Text;

            // Хэшируем пароль
            string hashedPassword;
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] bytes = Encoding.UTF8.GetBytes(password);
                byte[] hash = sha256.ComputeHash(bytes);
                hashedPassword = BitConverter.ToString(hash).Replace("-", "").ToLower();
            }

            OleDbCommand command = new OleDbCommand($"SELECT COUNT(*) FROM AuthorizationInfo WHERE login = \"{login}\"", mC);
            string existUser = command.ExecuteScalar().ToString();
            command = new OleDbCommand($"SELECT COUNT(*) FROM AuthorizationInfo WHERE login = \"{login}\" AND password = \"{hashedPassword}\"", mC);
            string existUserWithPass = command.ExecuteScalar().ToString();

            if (existUser == "0")
            {
                MessageBox.Show("Вы не зарегистрированы!");
                return;
            }
            else if (existUserWithPass == "0")
            {
                MessageBox.Show("Неверный пароль!");
                return;
            }
            else
            {
                command = new OleDbCommand($"SELECT status FROM AuthorizationInfo WHERE login = \"{login}\"", mC);
                OleDbCommand command1 = new OleDbCommand($"SELECT user_id FROM AuthorizationInfo WHERE login = \"{login}\"", mC);
                int root_id = Convert.ToInt32(command.ExecuteScalar());
                int user_id = Convert.ToInt32(command1.ExecuteScalar());

                if (root_id == 0)
                {
                    MessageBox.Show("Вы вошли в систему как администратор салона");
                    ClientView adminForm = new ClientView(1);
                    adminForm.ShowDialog();
                }
                else if (root_id == 1)
                {
                    MessageBox.Show("Вы вошли в систему как администратор информационной системы");
                    Admin_window adminPrimeForm = new Admin_window();
                    adminPrimeForm.ShowDialog();
                }
            }
        }

        private void loginBox_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
