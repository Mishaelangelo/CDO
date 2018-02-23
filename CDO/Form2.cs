using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CDO
{
    public partial class Form2 : Form
    {
        public static Form2 currentForm2;

        public Form2()
        {
            InitializeComponent();
        }

        private void обработка_кнопок_dataGridView(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                string параметрыСоединения = @"
                    Provider = Microsoft.ACE.OLEDB.12.0;
                    Data Source=БазаДанных.accdb;
                    User Id=admin; Password=;
                ";
                OleDbConnection соединение = new OleDbConnection(параметрыСоединения);
                соединение.Open();

                string текстЗапроса = @"
                  SELECT Max(платежи.[id платежа]) AS [Max-id платежа]
                  FROM платежи;
                ";
                OleDbCommand запрос = new OleDbCommand(текстЗапроса, соединение);
                OleDbDataReader считывательРезультатаЗапроса = запрос.ExecuteReader();
                считывательРезультатаЗапроса.Read();
                int max = Convert.ToInt32(считывательРезультатаЗапроса[0].ToString());


                //баг при id=0;


                if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() == "абонемент раньше не приобретался")
                {
                    int id = max + 1;
                    max = id;
                    string фио = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                    string дата_строкой = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");
                    int поступление_снятие = Convert.ToInt32(textBox1.Text) * (-1);
                    string коментарий = "стоимость за месяц";
                    string текстКоманды = "INSERT INTO платежи VALUES (" + id + ", '" + фио + "', '" + дата_строкой + "', " + поступление_снятие + ", '" + коментарий + "');";
                    OleDbCommand команда = new OleDbCommand(текстКоманды, соединение);
                    команда.ExecuteNonQuery();
                }
                else
                {
                    DateTime дата = Convert.ToDateTime(dataGridView1.Rows[e.RowIndex].Cells[1].Value);
                    while (дата < DateTime.Now)
                    {
                        int id = max + 1;
                        max = id;
                        string фио = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                        string дата_строкой = дата.ToString("dd.MM.yyyy HH:mm:ss");
                        int поступление_снятие = Convert.ToInt32(textBox1.Text) * (-1);
                        string коментарий = "стоимость за месяц";
                        string текстКоманды = "INSERT INTO платежи VALUES (" + id + ", '" + фио + "', '" + дата_строкой + "', " + поступление_снятие + ", '" + коментарий + "');";
                        OleDbCommand команда = new OleDbCommand(текстКоманды, соединение);
                        команда.ExecuteNonQuery();

                        дата = дата.AddMonths(1);
                    }
                }
                // заполнение сделано, клиента отображать больше не надо
                dataGridView1.Rows.RemoveAt(e.RowIndex);
                if (dataGridView1.Rows.Count == 1) {Form1.currentForm1.pictureBox1.Load("off.png"); }

            }
            if (e.ColumnIndex == 3)
            {
                string параметрыСоединения = @"
                    Provider = Microsoft.ACE.OLEDB.12.0;
                    Data Source=БазаДанных.accdb;
                    User Id=admin; Password=;
                ";
                OleDbConnection соединение = new OleDbConnection(параметрыСоединения);
                соединение.Open();

                string текстКоманды = "UPDATE клиенты SET активный = false WHERE фио = '"+ dataGridView1.Rows[e.RowIndex].Cells[0].Value+ "'";
                OleDbCommand команда = new OleDbCommand(текстКоманды, соединение);
                команда.ExecuteNonQuery();
                dataGridView1.Rows.RemoveAt(e.RowIndex);
                if (dataGridView1.Rows.Count == 1) {Form1.currentForm1.pictureBox1.Load("off.png"); }
            }
        }
    }
}