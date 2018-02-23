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

namespace CDO
{
    public partial class Form1 : Form
    {
        #region глобальные переменные
        public static Form1 currentForm1;
        string активнаяТаблица = "";
        DataTable table = new DataTable();
        List<string> СписокКлиентовДляПродлеванияАбонемента = new List<string>();
        List<DateTime> СписокПоследнихДатДляПередачи = new List<DateTime>();
        #endregion

        public Form1()
        {
            InitializeComponent();
        }

        #region Model
        private void ОбновитьСписокКлиентовДляПродлеванияАбонемента()
        {
            СписокКлиентовДляПродлеванияАбонемента.Clear();
            СписокПоследнихДатДляПередачи.Clear();

            List<string> СписокАктивныхКлиентов = new List<string>();

            //получение всех активных клиентов
            string параметрыСоединения = @"
            Provider = Microsoft.ACE.OLEDB.12.0;
            Data Source=БазаДанных.accdb;
            User Id=admin; Password=;
            ";
            OleDbConnection соединение = new OleDbConnection(параметрыСоединения);
            соединение.Open();

            string текстЗапроса = @"
                 SELECT клиенты.фио
                 FROM клиенты
                 WHERE (((клиенты.активный)=True));
            ";
            OleDbCommand запрос = new OleDbCommand(текстЗапроса, соединение);

            OleDbDataReader считывательРезультатаЗапроса = запрос.ExecuteReader();
            while (считывательРезультатаЗапроса.Read())
            {
                СписокАктивныхКлиентов.Add(считывательРезультатаЗапроса[0].ToString());
            }
            //для каждого активного клиента смотрим нужно ли ему продлить абонимент
            foreach (string клиент in СписокАктивныхКлиентов)
            {
                List<DateTime> СписокДатСнятия = new List<DateTime>();
                текстЗапроса = "SELECT платежи.дата FROM платежи WHERE (((платежи.фио)='" + клиент + "') AND ((платежи.[поступление/снятие])<0));";
                запрос = new OleDbCommand(текстЗапроса, соединение);

                считывательРезультатаЗапроса = запрос.ExecuteReader();
                while (считывательРезультатаЗапроса.Read())
                {
                    СписокДатСнятия.Add(Convert.ToDateTime(считывательРезультатаЗапроса[0].ToString()));
                }

                // найти последнюю дату снятия
                if (СписокДатСнятия.Count == 0)
                {
                    СписокКлиентовДляПродлеванияАбонемента.Add(клиент);
                    СписокПоследнихДатДляПередачи.Add(DateTime.Parse("01.01.0"));
                }

                else
                {
                    DateTime последняя_дата = СписокДатСнятия[0];
                    foreach (DateTime d in СписокДатСнятия)
                    {
                        if (последняя_дата < d)
                        {
                            последняя_дата = d;
                        }
                    }

                    if (последняя_дата.AddMonths(1) < DateTime.Now)
                    {
                        СписокПоследнихДатДляПередачи.Add(последняя_дата.AddMonths(1));
                        СписокКлиентовДляПродлеванияАбонемента.Add(клиент);
                    }
                }
            }

            // обновить состояние кнопки "!"
            if (СписокКлиентовДляПродлеванияАбонемента.Count > 0)
            {
                pictureBox1.Load("on.png");
            }
            if (СписокКлиентовДляПродлеванияАбонемента.Count == 0)
            {
                pictureBox1.Load("off.png");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = table;
            вывод_списка_должников(null,e);

            ОбновитьСписокКлиентовДляПродлеванияАбонемента();
        }
        #endregion


        #region View
        private void вывод_таблицы_платежей(object sender, EventArgs e)
        {
            активнаяТаблица = "платежи";
            string connectionString = @"
              Provider=Microsoft.ACE.OLEDB.12.0;
              Data Source=БазаДанных.accdb;
              User Id=admin; Password=;";

            string selectCommand = @"
              Select * 
              From платежи;";

            OleDbDataAdapter dataAdapter = new
                OleDbDataAdapter(selectCommand, connectionString);
            table.Clear();
            table.Columns.Clear();
            dataAdapter.Fill(table);

            dataGridView1.AutoResizeColumns(
            DataGridViewAutoSizeColumnsMode.AllCells);
            label1.Text = "Платежи";
            button1.Visible=true;
        }

        private void вывод_таблицы_клиентов(object sender, EventArgs e)
        {
            активнаяТаблица = "клиенты";
            string connectionString = @"
              Provider=Microsoft.ACE.OLEDB.12.0;
              Data Source=БазаДанных.accdb;
              User Id=admin; Password=;";

            string selectCommand = @"
              Select * 
              From клиенты;";

            OleDbDataAdapter dataAdapter = new
                OleDbDataAdapter(selectCommand, connectionString);
            table.Clear();
            table.Columns.Clear();
            dataAdapter.Fill(table);

            dataGridView1.AutoResizeColumns(
            DataGridViewAutoSizeColumnsMode.AllCells);
            label1.Text = "Клиенты";
            button1.Visible = true;
        }

        private void вывод_списка_должников(object sender, EventArgs e)
        {
            активнаяТаблица = "платежи";
            string connectionString = @"
              Provider=Microsoft.ACE.OLEDB.12.0;
              Data Source=БазаДанных.accdb;
              User Id=admin; Password=;";

            string selectCommand = @"SELECT платежи.фио, Sum(платежи.[поступление/снятие]) AS [Sum-поступление/снятие]
                FROM платежи
                GROUP BY платежи.фио
                HAVING (((Sum(платежи.[поступление/снятие]))<0))
                ORDER BY Sum(платежи.[поступление/снятие]);
                ";

            OleDbDataAdapter dataAdapter = new
                OleDbDataAdapter(selectCommand, connectionString);
            table.Clear();
            table.Columns.Clear();
            dataAdapter.Fill(table);

            dataGridView1.AutoResizeColumns(
            DataGridViewAutoSizeColumnsMode.AllCells);
            label1.Text = "Список должников";
            button1.Visible = false;
        }

        private void вывод_клиентов_для_продлевания_абонемента(object sender, EventArgs e)
        {
            ОбновитьСписокКлиентовДляПродлеванияАбонемента();

            Form2.currentForm2 = new Form2();
            //newForm.MdiParent = this;
            Form2.currentForm2.Show();
            Form2.currentForm2.dataGridView1.Rows.Clear();
            if (СписокКлиентовДляПродлеванияАбонемента.Count == 0) { return; }
            Form2.currentForm2.dataGridView1.Rows.Add(СписокКлиентовДляПродлеванияАбонемента.Count);
            for (int i = 0; i < СписокКлиентовДляПродлеванияАбонемента.Count; i++)
            {
                Form2.currentForm2.dataGridView1.Rows[i].Cells[0].Value = СписокКлиентовДляПродлеванияАбонемента[i];
            }
            for (int i = 0; i < СписокПоследнихДатДляПередачи.Count; i++)
            {

                Form2.currentForm2.dataGridView1.Rows[i].Cells[1].Value = СписокПоследнихДатДляПередачи[i].ToString();
                if (СписокПоследнихДатДляПередачи[i] == (DateTime.Parse("01.01.0")))
                {
                    Form2.currentForm2.dataGridView1.Rows[i].Cells[1].Value = "абонемент раньше не приобретался";
                }
            }
        }
        #endregion

        #region Control
        public static string nullToString(object value)
        {
            if (value == null)
                return "";
            return value.ToString();
        }

        private void сохранить_таблицу(object sender, EventArgs e)//сохранение
        {
            if (активнаяТаблица == "платежи")
            {
                string параметрыСоединения = @"
                  Provider=Microsoft.ACE.OLEDB.12.0;
                  Data Source=БазаДанных.accdb;
                  User Id=admin; Password=;";
                OleDbConnection соединение = new OleDbConnection(параметрыСоединения);
                соединение.Open();

                string текстКоманды = "Delete From платежи;";
                OleDbCommand команда = new OleDbCommand(текстКоманды, соединение);
                команда.ExecuteNonQuery();

                for (int row = 0; row <= dataGridView1.Rows.Count - 2; row++)
                {
                    int id = (int)dataGridView1.Rows[row].Cells[0].Value;
                    string фио = nullToString(dataGridView1.Rows[row].Cells[1].Value);
                    string дата = nullToString(dataGridView1.Rows[row].Cells[2].Value);
                    string дата_строкой = nullToString(String.Format("{0:yyyy-MM-dd HH:mm:ss}", дата));
                    int поступление_снятие = (int)dataGridView1.Rows[row].Cells[3].Value;
                    string коментарий = nullToString(dataGridView1.Rows[row].Cells[4].Value);
                    текстКоманды = "INSERT INTO платежи VALUES (" + id + ", '" + фио + "', '"+дата_строкой+"', " + поступление_снятие + ", '" + коментарий + "');";
                    команда = new OleDbCommand(текстКоманды, соединение);
                    команда.ExecuteNonQuery();
                }
                MessageBox.Show("Сохранено!");
            }

            if (активнаяТаблица == "клиенты")
            {
                string параметрыСоединения = @"
                  Provider=Microsoft.ACE.OLEDB.12.0;
                  Data Source=БазаДанных.accdb;
                  User Id=admin; Password=;";
                OleDbConnection соединение = new OleDbConnection(параметрыСоединения);
                соединение.Open();

                string текстКоманды = "Delete From клиенты;";
                OleDbCommand команда = new OleDbCommand(текстКоманды, соединение);
                команда.ExecuteNonQuery();

                for (int row = 0; row <= dataGridView1.Rows.Count - 2; row++)
                {
                    
                    int id = (int)dataGridView1.Rows[row].Cells[0].Value;
                    string фио = nullToString(dataGridView1.Rows[row].Cells[1].Value);
                    bool активный = (bool)dataGridView1.Rows[row].Cells[2].Value;
                    string телефон = nullToString(dataGridView1.Rows[row].Cells[3].Value);
                    string дополнительная_связь = nullToString(dataGridView1.Rows[row].Cells[4].Value);
                    string день_рождения = nullToString(dataGridView1.Rows[row].Cells[5].Value);
                    string примечание_по_здоровью = nullToString(dataGridView1.Rows[row].Cells[6].Value);
                    string общее_примечание = nullToString(dataGridView1.Rows[row].Cells[7].Value);
                    string пояс = nullToString(dataGridView1.Rows[row].Cells[8].Value);
                    string дата_вступления = nullToString(dataGridView1.Rows[row].Cells[9].Value);
                    

                    текстКоманды = "INSERT INTO клиенты VALUES (" + id + ", '" + фио + "'," + активный + ", '" + телефон + "', '" + дополнительная_связь + "', '" + день_рождения+ "', '" + примечание_по_здоровью + "', '" + общее_примечание + "', '"+пояс + "', '" + дата_вступления + "');";
                    команда = new OleDbCommand(текстКоманды, соединение);
                    команда.ExecuteNonQuery();
                }
                MessageBox.Show("Сохранено!");
            }

            ОбновитьСписокКлиентовДляПродлеванияАбонемента();

        }
        #endregion
    }
}
