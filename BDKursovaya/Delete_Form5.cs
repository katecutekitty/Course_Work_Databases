using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Drawing.Text;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace cursovaya
{
    public partial class Delete_Form5 : Form
    {
        Admin_window admin_Window;
        private OleDbConnection DbCon;
        public Delete_Form5(OleDbConnection dbCon, Admin_window admin_w)
        {
            DbCon = dbCon;
            admin_Window = admin_w;
            InitializeComponent();
            Add_Tables_names();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked) { Add_Field_values(); }
            else { listView1.Items.Clear(); }
        }
        public void Add_Tables_names()
        {
            comboBox1.Text = "";
            comboBox1.Items.Clear();
            DataTable tables = DbCon.GetSchema("Tables");
            foreach (DataRow row in tables.Rows)
            {
                string tableName = row["TABLE_NAME"].ToString();
                if (!tableName.StartsWith("MSys"))
                    if (!tableName.StartsWith("f_6F231996"))
                        if (!tableName.StartsWith("fk_special_table"))
                            if (!tableName.StartsWith("~TM"))
                                comboBox1.Items.Add(tableName);
            }
        }
        private void Add_Field_values()
        {
            listView1.Clear();
            if (comboBox2.SelectedItem != null)
            {
                listView1.Columns.Add("");
                string table_name = comboBox1.SelectedItem.ToString();
                string field_name = comboBox2.SelectedItem.ToString();
                string values = $"SELECT {field_name} FROM {table_name}";
                OleDbCommand cmd = new OleDbCommand(values, DbCon);
                OleDbDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    ListViewItem item = new ListViewItem(reader[0].ToString());
                    listView1.Items.Add(item);
                }
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
            {
                string tableName = comboBox1.SelectedItem.ToString();

                // Подтверждение удаления
                DialogResult confirm = MessageBox.Show($"Вы действительно хотите удалить таблицу {tableName}?",
                                                       "Подтверждение удаления",
                                                       MessageBoxButtons.YesNo);
                if (confirm == DialogResult.Yes)
                {

                    string s = $"SELECT* FROM MSysRelationships WHERE tblOne = {tableName} OR tblTwo = {tableName};";
                    OleDbCommand cmd = new OleDbCommand(s, DbCon);
                    var result = cmd.ExecuteReader();
                    while (result.Read())
                    {
                        MessageBox.Show(result[0].ToString());
                    }

                    //// Удаление внешних ключей, если есть
                    //if (RowExists(tableName))
                    //{
                    //    delete_keys(tableName);
                    //}

                    //// Удаляем таблицу
                    //try
                    //{
                    //    string sqlDropTable = $"DROP TABLE [{tableName}]";
                    //    using (OleDbCommand cmd = new OleDbCommand(sqlDropTable, DbCon))
                    //    {
                    //        cmd.ExecuteNonQuery();
                    //    }
                    //    MessageBox.Show($"Таблица {tableName} успешно удалена.", "Успех");
                    //}
                    //catch (Exception ex)
                    //{
                    //    MessageBox.Show($"Ошибка при удалении таблицы {tableName}: {ex.Message}", "Ошибка");
                    //}

                    //// Обновляем список таблиц
                    //Add_Tables_names();
                }
            }

        }

        


        // Удаление внешнего ключа
        private void DropForeignKey(string tableName, string fkName)
        {
            try
            {
                string dropConstraintSql = $"ALTER TABLE [{tableName}] DROP CONSTRAINT [{fkName}]";
                using (OleDbCommand cmd = new OleDbCommand(dropConstraintSql, DbCon))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении внешнего ключа {fkName}: {ex.Message}", "Ошибка");
            }
        }

        // Удаление индекса
        private void DropIndex(string tableName, string indexName)
        {
            try
            {
                string dropIndexSql = $"DROP INDEX [{indexName}] ON [{tableName}]";
                using (OleDbCommand cmd = new OleDbCommand(dropIndexSql, DbCon))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении индекса {indexName}: {ex.Message}", "Ошибка");
            }
        }

        // Функция для удаления ключей и связанных данных
        private void delete_keys(string tableName)
        {
            DataTable dt = GetRowData(tableName);

            foreach (DataRow drrows in dt.Rows)
            {
                string keyTable = drrows["key_table"].ToString();
                string linkTable = drrows["link_table"].ToString();
                string fkName = drrows["key_name"].ToString();
                string fieldName = drrows["field_name"].ToString();

                try
                {
                    // Удаляем внешний ключ
                    if (linkTable == tableName)
                    {
                        DropForeignKey(linkTable, fkName);
                    }
                    else if (keyTable == tableName)
                    {
                        DropForeignKey(keyTable, fkName);
                    }

                    // Удаляем уникальный индекс
                    string indexName = $"idx_{keyTable}_{fieldName}";
                    DropIndex(keyTable, indexName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при удалении зависимостей таблицы {tableName}: {ex.Message}", "Ошибка");
                }
            }

            // Удаляем данные из fk_special_table
            DeleteRow(tableName);
        }

        // Удаление строки из fk_special_table
        private void DeleteRow(string tableName)
        {
            try
            {
                string deleteSql = $"DELETE FROM fk_special_table WHERE key_table = '{tableName}' OR link_table = '{tableName}'";
                using (OleDbCommand cmd = new OleDbCommand(deleteSql, DbCon))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении строки из fk_special_table: {ex.Message}", "Ошибка");
            }
        }

        // Удаление таблицы
        private void DropTable(string tableName)
        {
            try
            {
                string dropTableSql = $"DROP TABLE [{tableName}]";
                using (OleDbCommand cmd = new OleDbCommand(dropTableSql, DbCon))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении таблицы {tableName}: {ex.Message}", "Ошибка");
            }
        }



        // Проверка существования индекса
        private bool IndexExists(string tableName, string indexName)
        {
            string query = $"SELECT COUNT(*) FROM sys.indexes WHERE object_id = OBJECT_ID('{tableName}') AND name = '{indexName}'";
            using (OleDbCommand cmd = new OleDbCommand(query, DbCon))
            {
                object result = cmd.ExecuteScalar();
                return result != null && Convert.ToInt32(result) > 0;
            }
        }




        //Проверка наличия ключа у таблицы
        private bool RowExists(string tableName)
        {
            string sqlQuery = $"SELECT COUNT(*) FROM fk_special_table WHERE link_table = '{tableName}' OR key_table = '{tableName}'";
            using (OleDbCommand cmd = new OleDbCommand(sqlQuery, DbCon))
            {
                object result = cmd.ExecuteScalar();
                return result != null && Convert.ToInt32(result) > 0;
            }
        }


        //получение строк с удаляемой таблицей
        private DataTable GetRowData(string tableName)
        {
            DataTable result_table = new DataTable();
            string sql_get = $"SELECT * FROM fk_special_table WHERE link_table = '{tableName}' OR key_table = '{tableName}'";
            OleDbDataAdapter adapter = new OleDbDataAdapter(sql_get, DbCon);
            adapter.Fill(result_table);
            return result_table;
        }
        public void Add_Fields_names()
        {
            comboBox2.Text = "";
            comboBox2.Items.Clear();
            if (comboBox1.SelectedItem != null)
            {
                string tableName = comboBox1.SelectedItem.ToString();
                DataTable columns = DbCon.GetSchema("Columns", new string[] { null, null, tableName, null });
                foreach (DataRow row in columns.Rows)
                {
                    string columnName = row["COLUMN_NAME"].ToString();
                    comboBox2.Items.Add(columnName);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex > -1)
            {
                comboBox2.Items.Clear();
                Add_Fields_names();
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex > -1)
            {
                //очистка listview
                listView1.Items.Clear();
                if (checkBox1.Checked) { Add_Field_values(); }
            }
        }


    }
}
