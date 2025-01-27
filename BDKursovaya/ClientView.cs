using System;
using System.Data.OleDb;
using System.Net.Mail;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Data.SqlClient;
using System.Data;

namespace BDKursovaya
{
    public partial class ClientView : Form
    {
        public static string s = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BeautySalon2.mdb;";
        private OleDbConnection mC;
        private int currentId;

        bool Is_it_start = true;

        public ClientView(int cntClientId)
        {
            
            InitializeComponent();
            mC = new OleDbConnection(s);
            mC.Open();

            currentId = cntClientId;

            OleDbCommand cmd = new OleDbCommand("SELECT * FROM Services", mC);
            var l = cmd.ExecuteReader();
            comboBox1.Items.Clear();
            while (l.Read()) comboBox1.Items.Add(l[0].ToString());
            l.Close();
            RefreshList();

            OleDbCommand cmd1 = new OleDbCommand("SELECT * FROM Services", mC);
            var l1 = cmd1.ExecuteReader();
            comboBox4.Items.Clear();
            while (l1.Read()) comboBox4.Items.Add(l1[0].ToString());
            l1.Close();



            dateTimePicker1.Value = DateTime.Now.AddDays(1);
        }

        private int getPrice(string serviceTitle)
        {
            int price = Convert.ToInt32(new OleDbCommand($"SELECT price from Services WHERE title = '{serviceTitle}'", mC).ExecuteScalar());
            return price;
        }

        private Tuple<int, string> getIdAndCat(string master)
        {
            OleDbCommand c = new OleDbCommand($"SELECT id, category FROM Masters WHERE fullName = '{master}'", mC);
            var l = c.ExecuteReader();
            Tuple<int, string> result = Tuple.Create(0, "");
            while (l.Read())
            {
                result = Tuple.Create(Convert.ToInt32(l[0]), l[1].ToString());
            }
            l.Close();
            return result;
        }

        private string[] Get_list_of_bool_from_base(System.Data.OleDb.OleDbDataReader l)
        {
            string[] splited_text = l[4].ToString().Split(new char[] { ',' });
            return splited_text;
        }

        private bool FreeDayCheck(System.Data.OleDb.OleDbDataReader l)
        {
            int dateTimePicker1_day = DateTime.Parse(dateTimePicker1.Text.ToString()).Day;
            int between_today_and_date = dateTimePicker1_day-DateTime.Now.Day-1;
            string[] days = Get_list_of_bool_from_base(l);
            if (days[between_today_and_date] == "true")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem == null || comboBox2.SelectedItem == null || textBox2.Text == String.Empty)
            {
                MessageBox.Show("Не заполнено одно или несколько полей");
            }
            else
            {
                var selectedService = comboBox1.SelectedItem.ToString();
                var selectedMaster = comboBox2.SelectedItem.ToString();
                string telephoneNumber = textBox2.Text;

                var masterArgs = getIdAndCat(selectedMaster);

                int selectedMasterId = masterArgs.Item1;
                string selectedMasterCategory = masterArgs.Item2;

                int servicePrice = getPrice(selectedService);

                var selectedDate = dateTimePicker1.Value.Date;

                OleDbCommand cmd = new OleDbCommand($"SELECT id FROM Clients WHERE phoneNumber = '{telephoneNumber}'", mC);
                int addedClientId = Convert.ToInt16(cmd.ExecuteScalar());


                OleDbCommand cd = new OleDbCommand($"SELECT COUNT (*) FROM Bookings LEFT JOIN Clients ON Bookings.client_id = Clients.id WHERE Clients.phoneNumber = '{telephoneNumber}' AND appointment_date_time < '{DateTime.Now.ToString()}'", mC);
                int visitsCount = Convert.ToInt32(cd.ExecuteScalar());
                double discountSize = 0;
                if (visitsCount > 3) discountSize = 0.05;
                if (visitsCount > 5) discountSize = 0.07;

                cmd = new OleDbCommand($"INSERT INTO Bookings (client_id, master_id, service_title, appointment_date_time, totalPrice) VALUES ({addedClientId}, {selectedMasterId}, '{selectedService}', '{selectedDate}', {servicePrice * (1-discountSize)})",mC);
                cmd.ExecuteNonQuery();

                if (textBox3.Text != String.Empty)
                {
                    SendAnEmail(textBox3.Text, selectedService, selectedMaster, selectedDate.ToString());
                }
            }
            RefreshList();
        }

        private bool emailValidated(string email)
        {
            string emailRegex = @"^[^@\s]+@[^@\s]+\.[^@\s]+$";
            return Regex.IsMatch(email, emailRegex);
        }

        private void SendAnEmail(string emailAddress, string service, string master, string date)
        {
                if (!emailValidated(emailAddress))
                {
                    MessageBox.Show("Некорректный адрес электронной почты. Пожалуйста, проверьте введённые данные.", "Ошибка");
                    return;
                }
                MailMessage mail = new MailMessage(); 
                mail.From = new MailAddress("poznahirko.cat@yandex.ru"); 
                mail.To.Add(new MailAddress(emailAddress)); 
                mail.Subject = "Напоминание о записи"; 
                mail.Body = $"Здравствуйте!\nНапоминаем, что у вас запланирована запись:\nУслуга: {service}\nМастер: {master}\nДата: {date}\nС уважением, салон красоты Центрифуга."; 
                SmtpClient client = new SmtpClient(); 
                client.Host = "smtp.yandex.ru"; 
                client.Port = 587; 
                client.EnableSsl = true; 
                client.Credentials = new NetworkCredential("poznahirko.cat@yandex.ru", "ptjkrnmoncunjltt");
        
                try
                {
                    client.Send(mail);
                }
                catch (SmtpException ex)
                {
                    MessageBox.Show(ex.Message);
                }

                //string senderEmail = "poznahirko.cat@yandex.ru";
                //string senderPassword = "ptjkrnmoncunjltt";
        }

        private void RefreshMasters() 
        {
            var selectedService = comboBox1.SelectedItem.ToString();

            OleDbCommand cmd = new OleDbCommand($"SELECT * FROM Masters LEFT JOIN Services ON Masters.spec_id = Services.category_id WHERE Services.title = '{selectedService}'", mC);


            var l = cmd.ExecuteReader();
            comboBox2.Items.Clear();
            while (l.Read())
            {
                if (FreeDayCheck(l))
                {
                    comboBox2.Items.Add(l[1].ToString());
                }
            }
            l.Close();

            
        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshMasters();
            RefreshPrice();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshPrice();
        }

        private void RefreshList()
        {
            OleDbCommand cmd = new OleDbCommand($"SELECT Bookings.*, Clients.fullName FROM Bookings LEFT JOIN Clients ON Bookings.client_id = Clients.id", mC);
            var l = cmd.ExecuteReader();
            listBox1.Items.Clear();
            while (l.Read())
            {
                listBox1.Items.Add(l[0].ToString() + ' ' + l[1].ToString() + ' ' + l[2].ToString() + ' ' + l[3].ToString() + ' ' + l[4].ToString() + ' ' + l[5].ToString() + ' ' + l[6].ToString());
            }
            l.Close();
        }

        private void RefreshListSort()
        {
            OleDbCommand cmd = new OleDbCommand($"SELECT Bookings.*, Clients.fullName FROM Bookings LEFT JOIN Clients ON Bookings.client_id = Clients.id ORDER BY appointment_date_time ASC", mC);
            var l = cmd.ExecuteReader();
            listBox1.Items.Clear();
            while (l.Read())
            {
                listBox1.Items.Add(l[0].ToString() + ' ' + l[1].ToString() + ' ' + l[2].ToString() + ' ' + l[3].ToString() + ' ' + l[4].ToString() + ' ' + l[5].ToString() + ' ' + l[6].ToString());
            }
            l.Close();
        }

        private void RefreshPrice()
        {
            int currentPrice = 0;
            if (comboBox1.SelectedItem != null)
            {
                currentPrice = getPrice(comboBox1.SelectedItem.ToString());
                if (comboBox2.SelectedItem != null)
                {
                    string category = getIdAndCat(comboBox2.SelectedItem.ToString()).Item2;
                    if (category == "II")
                    {
                        currentPrice *= 2;
                    }
                    if (category == "III")
                    {
                        currentPrice *= 3;
                    }
                    if (category == "Высшая") { currentPrice *= 4; }
                }
            }
            textBox1.Text = currentPrice.ToString();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            button1.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItems.Count > 0) {
                string selectedBooking = listBox1.SelectedItems[0].ToString();
                int _id = Convert.ToInt32(selectedBooking.Substring(0, selectedBooking.IndexOf(' ')));
                OleDbCommand cmd = new OleDbCommand($"DELETE FROM Bookings WHERE id = {_id}", mC);
                cmd.ExecuteNonQuery();

                RefreshList();
            }
            else button1.Enabled = false;
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            button1.Enabled = true;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                RefreshListSort();
            }
            else RefreshList();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (Is_it_start)
            {
                Is_it_start = false;
            }
            else
            {
                DateTime dateTimePicker1_day = DateTime.Parse(dateTimePicker1.Text.ToString());
                int between_today_and_date = (dateTimePicker1_day - DateTime.Now).Days - 1;
                if ((DateTime.Parse(dateTimePicker1.Text.ToString()).Day - (DateTime.Now.Day)) <= 7)
                {
                    RefreshMasters();
                }
                else
                {
                    comboBox2.Items.Clear();
                    MessageBox.Show("Максимальное время записи от сегодняшнего дня - неделя");
                }
            }

            
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)//фильтрация по виду услуги
        {
            OleDbCommand cmd = new OleDbCommand($"SELECT Bookings.*, Masters.fullName FROM Bookings LEFT JOIN Masters ON Bookings.master_id = Masters.id WHERE service_title='{comboBox4.SelectedItem.ToString()}'", mC);
            var l = cmd.ExecuteReader();
            listBox1.Items.Clear();
            while (l.Read())
            {
                listBox1.Items.Add(l[0].ToString() + ' ' + l[1].ToString() + ' ' + l[3].ToString() + ' ' + l[4].ToString());
            }
            l.Close();
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)//фильтрация по времени
        {
            OleDbCommand cmd = new OleDbCommand($"SELECT Bookings.*, Masters.fullName FROM Bookings LEFT JOIN Masters ON Bookings.master_id = Masters.id WHERE appointment_date_time > '{DateTime.Parse(dateTimePicker2.Text)}'", mC);
            var l = cmd.ExecuteReader();
            listBox1.Items.Clear();
            while (l.Read())
            {
                    listBox1.Items.Add(l[0].ToString() + ' ' + l[1].ToString() + ' ' + l[3].ToString() + ' ' + l[4].ToString());
            }
            l.Close();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                string clientsCount = "SELECT count(*) FROM (SELECT DISTINCT client_id FROM Bookings)";
                string query = $@"SELECT SUM(totalPrice) AS TotalRevenue, COUNT(id) AS TotalBookings FROM Bookings
WHERE appointment_date_time LIKE '???{DateTime.Now.Month.ToString()}*'";
                MessageBox.Show(DateTime.Now.Month.ToString());


                OleDbCommand command = new OleDbCommand(query, mC);
                OleDbDataReader reader = command.ExecuteReader();

                OleDbCommand cmd = new OleDbCommand(clientsCount, mC);
                int uniqueClients = Convert.ToInt32(cmd.ExecuteScalar());

                // Создание Word-документа
                var wordApp = new Word.Application();
                var document = wordApp.Documents.Add();

                // Добавление заголовка
                var titleParagraph = document.Content.Paragraphs.Add();
                titleParagraph.Range.Text = "Отчет по услугам за каждый месяц";
                titleParagraph.Range.Font.Size = 16;
                titleParagraph.Range.Font.Bold = 1;
                titleParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                titleParagraph.Range.InsertParagraphAfter();

                // Подготовка данных
                var table = document.Tables.Add(document.Bookmarks["\\endofdoc"].Range, 1, 4);
                table.Borders.Enable = 1;

                // Установка заголовков таблицы
                table.Cell(1, 1).Range.Text = "Месяц";
                table.Cell(1, 2).Range.Text = "Уникальные клиенты";
                table.Cell(1, 3).Range.Text = "Суммарная выручка";
                table.Cell(1, 4).Range.Text = "Общее количество записей";
                table.Rows[1].Range.Font.Bold = 1;
                table.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                int rowIndex = 2;

                while (reader.Read())
                {
                    table.Rows.Add();
                    table.Cell(rowIndex, 1).Range.Text = DateTime.Now.Month.ToString();
                    table.Cell(rowIndex, 2).Range.Text = uniqueClients.ToString();
                    table.Cell(rowIndex, 3).Range.Text = reader["TotalRevenue"].ToString();
                    table.Cell(rowIndex, 4).Range.Text = reader["TotalBookings"].ToString();
                    rowIndex++;
                }

                reader.Close();

                // Сохранение документа
                string filePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MonthlyReport.docx";
                document.SaveAs2(filePath);
                document.Close();
                wordApp.Quit();

                MessageBox.Show($"Отчет успешно создан: {filePath}", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании отчета: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
