using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace cursovaya
{
    public partial class Add_TABLE : Form
    {
        Admin_window admin_Window;
        private OleDbConnection DbCon;
        List<string> list_tables = new List<string>();

        public Add_TABLE(OleDbConnection dbCon, Admin_window admin_w)
        {
            admin_Window = admin_w;
            DbCon = dbCon;
            InitializeComponent();
            comboBox1.Enabled = false;
            update_tables();



        }
        public void update_tables()
        {
            list_tables.Clear();
            DataTable dt = DbCon.GetSchema("Tables");
            foreach (DataRow row in dt.Rows)
            {
                list_tables.Add(row["TABLE_NAME"].ToString());
                if (!row["TABLE_NAME"].ToString().Contains("MSys"))
                    comboBox1.Items.Add(row["TABLE_NAME"].ToString());
            }


        }





        //public List<string> GetTableNamesSQLServer(string connectionString)
        //{
        //    List<string> tableNames = new List<string>();
        //    string query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'";


        //    using (SqlConnection connection = new SqlConnection(connectionString))
        //    {
        //        using (SqlCommand command = new SqlCommand(query, connection))
        //        {
        //            connection.Open();
        //            using (SqlDataReader reader = command.ExecuteReader())
        //            {
        //                while (reader.Read())
        //                {
        //                    tableNames.Add(reader.GetString(0));
        //                }
        //            }
        //        }
        //    }
        //    return tableNames;
        //}


        






        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Название таблицы не может быть пустым.", "Ошибка");
            }
            else
            {
                try
                {
                    string dbname = textBox1.Text;
                    string s = $"CREATE TABLE {dbname}";

                    OleDbCommand cmd = new OleDbCommand(s, DbCon);
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Таблица с таким названием уже существует.", "Ошибка");
                }
            }
            update_tables();

        }

        // Добавить поле
        private void button2_Click(object sender, EventArgs e)
        {
            string tablename = textBox1.Text;
            string fieldname = textBox2.Text;
            string sql_add_back = "";


            string type = "NONE";
            if (radioButton1.Checked == true)
            {
                type = "TEXT";
            }
            else
            {
                type = "INT";
            }

            if (list_tables.Contains(tablename))
            {



                string sql = $"ALTER TABLE [{tablename}] ADD COLUMN [{fieldname}] TEXT ";
                if (radioButton1.Checked == true && checkBox1.Checked == false)//поле числовое не ключевое
                {
                    sql = $"ALTER TABLE [{tablename}] ADD COLUMN [{fieldname}] INT";
                    OleDbCommand cmd = new OleDbCommand(sql, DbCon);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("1");
                    listBox1.Items.Add($"{fieldname} {type}");

                }
                if (radioButton1.Checked == true && checkBox1.Checked == true) //поле числовое ключевое
                {

                    sql = $"ALTER TABLE [{tablename}] ADD COLUMN [{fieldname}] INT PRIMARY KEY";
                    sql_add_back = $"ALTER TABLE [{comboBox1.Text}] ADD COLUMN [{fieldname}] INT";
                    OleDbCommand cmd = new OleDbCommand(sql, DbCon);
                    cmd.ExecuteNonQuery();

                    OleDbCommand cmd2 = new OleDbCommand(sql_add_back, DbCon);
                    cmd2.ExecuteNonQuery();

                    string foreign_sql = $"ALTER TABLE {comboBox1.Text} ADD FOREIGN KEY({fieldname}) REFERENCES {tablename}({fieldname})";
                    OleDbCommand cmd3 = new OleDbCommand(foreign_sql, DbCon);
                    cmd3.ExecuteNonQuery();

                    listBox1.Items.Add($"{fieldname} {type}");
                }
                if (radioButton2.Checked == true && checkBox1.Checked == false) //поле текстовое не ключевое
                {
                    sql = $"ALTER TABLE [{tablename}] ADD COLUMN [{fieldname}] TEXT";
                    OleDbCommand cmd = new OleDbCommand(sql, DbCon);
                    cmd.ExecuteNonQuery();

                    listBox1.Items.Add($"{fieldname} {type}");
                }
                if (radioButton2.Checked == true && checkBox1.Checked == true) //поле текстовое ключевое
                {
                    sql = $"ALTER TABLE [{tablename}] ADD COLUMN [{fieldname}] TEXT PRIMARY KEY";
                    sql_add_back = $"ALTER TABLE [{comboBox1.Text}] ADD COLUMN [{fieldname}] TEXT";

                    OleDbCommand cmd = new OleDbCommand(sql, DbCon);
                    cmd.ExecuteNonQuery();


                    OleDbCommand cmd2 = new OleDbCommand(sql_add_back, DbCon);
                    cmd2.ExecuteNonQuery();


                    string foreign_sql = $"ALTER TABLE {comboBox1.Text} ADD FOREIGN KEY({fieldname}) REFERENCES {tablename}({fieldname})";
                    OleDbCommand cmd3 = new OleDbCommand(foreign_sql, DbCon);
                    cmd3.ExecuteNonQuery();

                    listBox1.Items.Add($"{fieldname} {type}");
                }
            }
            else
            {
                MessageBox.Show("Таблицы с таким названием не найдено");
            }
        }



        // Обработка CheckBox
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            comboBox1.Enabled = checkBox1.Checked;

        }

        // Переключение RadioButton
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked) radioButton2.Checked = false;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked) radioButton1.Checked = false;
        }

        private void button3_Click(object sender, EventArgs e)//назад
        {
            this.Close();
            admin_Window.Show();
        }

        private void button4_Click(object sender, EventArgs e)//очистить информацию
        {
            listBox1.Items.Clear();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
