using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace ActivCollab_1
{
    public partial class Form1 : Form
    {
        //--Объявление переменных--//
        String data1 = "";
        String data2 = "";
        String data3 = "";
        String data4 = "";
        String control = "";
        String project = "";
        String category = "";
        string data_sozd = "";
        string data_zakr = "";
        string zad = "";
        int strok = 0;


        string job = "";
        string text_body = "";
        string text_name = "";

        string sql1 = "select name from acx4_projects order by id desc";// запрос Проекты
        string sql2 = "select c.name from acx4_companies c where c.state = '3'";// запрос Управления
        string sql3 = "";//запрос работники


        String connString = "Server=10.129.116.5;Database=ac;User Id=ilya;password= db!!!;CharSet=utf8";// строка подключения к БД

        public Form1()
        {
            InitializeComponent();
            //--подключение к бд--//
            MySqlConnection conn = new MySqlConnection(connString);// создаём объект для подключения к БД
            conn.Open();// устанавливаем соединение с БД

            MySqlCommand command1 = new MySqlCommand(sql1, conn); // объект для выполнения SQL-запроса
            MySqlDataReader reader1 = command1.ExecuteReader(); // объект для чтения ответа сервера
            while (reader1.Read())
                this.comboBox4.Items.Add(reader1.GetString(0));
            reader1.Close(); // закрываем reader


            MySqlCommand command2 = new MySqlCommand(sql2, conn); // объект для выполнения SQL-запроса
            MySqlDataReader reader2 = command2.ExecuteReader(); // объект для чтения ответа сервера
            while (reader2.Read())
                this.comboBox3.Items.Add(reader2.GetString(0));
            reader2.Close(); // закрываем reader

            MySqlCommand command3 = new MySqlCommand(sql2, conn); // объект для выполнения SQL-запроса
            MySqlDataReader reader3 = command3.ExecuteReader(); // объект для чтения ответа сервера
            while (reader3.Read())
                this.comboBox1.Items.Add(reader3.GetString(0));
            reader3.Close(); // закрываем reader

            conn.Close(); // закрываем соединение с БД
            conn.Dispose();// Уничтожить объект, освободить ресурс.

            //настройка календаря
            dateTimePicker1.Format = DateTimePickerFormat.Short;
            dateTimePicker1.ValueChanged += dateTimePicker1_ValueChanged;
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            // data1 = DateTime.Now.ToString("yyyy-MM-dd");
            data1 = "2022-08-08";

            dateTimePicker2.Format = DateTimePickerFormat.Short;
            dateTimePicker2.ValueChanged += dateTimePicker2_ValueChanged;
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            // data2 = DateTime.Now.ToString("yyyy-MM-dd");
            data2 = "2022-08-08";

            dateTimePicker3.Format = DateTimePickerFormat.Short;
            dateTimePicker3.ValueChanged += dateTimePicker3_ValueChanged;
            dateTimePicker3.Format = DateTimePickerFormat.Custom;
            // data3 = DateTime.Now.ToString("yyyy-MM-dd");
            data3 = "2022-08-08";

            dateTimePicker4.Format = DateTimePickerFormat.Short;
            dateTimePicker4.ValueChanged += dateTimePicker4_ValueChanged;
            dateTimePicker4.Format = DateTimePickerFormat.Custom;
            // data4 = DateTime.Now.ToString("yyyy-MM-dd");
            data4 = "2022-08-08";


        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            data1 = String.Format(dateTimePicker1.Text);
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            data2 = String.Format(dateTimePicker2.Text);
        }

        private void dateTimePicker4_ValueChanged(object sender, EventArgs e)
        {
            data3 = String.Format(dateTimePicker4.Text);
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            data4 = String.Format(dateTimePicker3.Text);
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.SelectedIndex != -1)
            {
                control = "and c.name = '" + comboBox3.SelectedItem.ToString() + "'";
            }


        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox4.SelectedIndex != -1)
            {
                project = "and p.name  = '" + comboBox4.SelectedItem.ToString() + "'";

                string sql3 = "select cat.name from acx4_categories cat join acx4_projects p on cat.parent_id = p.id where p.name  = '" + comboBox4.SelectedItem.ToString() + "' and cat.type = 'TaskCategory'";
                MySqlConnection conn = new MySqlConnection(connString);// создаём объект для подключения к БД
                conn.Open();// устанавливаем соединение с БД
                MySqlCommand command = new MySqlCommand(sql3, conn); // объект для выполнения SQL-запроса
                MySqlDataReader reader = command.ExecuteReader(); // объект для чтения ответа сервера
                while (reader.Read())
                    this.comboBox5.Items.Add(reader.GetString(0));
                reader.Close(); // закрываем reader
                conn.Close(); // закрываем соединение с БД
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox5.SelectedIndex != -1)
            {
                category = "and cat.name = '" + comboBox5.SelectedItem.ToString() + "'";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            //--подключение к бд--//
            MySqlConnection conn2 = new MySqlConnection(connString);// создаём объект для подключения к БД
            conn2.Open();// устанавливаем соединение с БД

            if (checkBox1.Checked == false && data1 != data2) { data_sozd = "and po.created_on BETWEEN '" + data1 + "' and '" + data2 + "'"; } else {data_sozd = ""; }
            if (checkBox1.Checked == false && data1 == data2) {data_sozd= "and po.created_on like '%" + data1 + "%'"; } else { data_sozd = ""; }
            if (checkBox2.Checked == false) { data_zakr = "and st.completed_on BETWEEN '" + data3 + "' and '" + data4 + "'"; } else { data_zakr = ""; }
            if (richTextBox1.Text != "") { text_body = "and po.body like '%"+ richTextBox1.Text + "%'"; } else { text_body = ""; }
            if (textBox3.Text != "") { text_name = "and po.name like '%" + textBox3.Text + "%'"; } else { text_name = ""; }
            if (checkBox3.Checked == true) { zad = "and po.completed_on is NULL"; } else { zad = ""; }


            string sql3 = "SELECT po.created_by_name as 'Кем создана', po.created_by_email as 'E-mail создателя', po.name as 'Тема задачи',  po.integer_field_1 as 'Номер задачи', l.name as 'Ярлык', c.name as 'Отдел', cat.name as 'Категория', CONCAT(u.last_name, ' ', u.first_name) AS 'Исполнитель задачи', po.created_on as 'Дата создания задачи', cast(po.due_on as date) as 'Срок выполнения задачи', st.created_on 'Дата создания подзадачи', cast(st.due_on as date) as 'Срок выполнения подзадачи',  st.completed_on as 'Дата закрытия подзадачи', case when DATEDIFF(cast(st.completed_on as date), cast(st.due_on as date)) > 0 then DATEDIFF(cast(st.completed_on as date), cast(st.due_on as date)) else ' ' end as 'Просрочка подзадачи', po.completed_on as 'Дата закрытия задачи', case when DATEDIFF(cast(po.completed_on as date), cast(po.due_on as date)) > 0 then DATEDIFF(cast(po.completed_on as date), cast(po.due_on as date)) else ' ' end as 'Просрочка задачи'  FROM acx4_project_objects as po left join(select* from acx4_subtasks where state in (2,3)) st on st.parent_id = po.id   left join acx4_users u on st.assignee_id = u.id left join acx4_companies c on u.company_id = c.id left join acx4_labels l on po.label_id = l.id left join acx4_categories cat on cat.id = po.category_id left join acx4_projects p on cat.parent_id = p.id where po.state in (2, 3) " + control + " " + data_sozd + " " + data_zakr + " " + project + " " + category + " " + job + " " + text_body + " " + text_name + zad + " order by po.created_on";




            textBox2.Text = sql3;



            MySqlCommand command2 = new MySqlCommand(sql3, conn2); // объект для выполнения SQL-запроса
            MySqlDataReader reader2 = command2.ExecuteReader(); // объект для чтения ответа сервера
            if (reader2.HasRows == false)                         //  если ничего нет
            {
                MessageBox.Show("По этому запросу ничего не найдено"); //то сообщение выводит
            }
            else
            { //дальше идет магия работы с датагрид, которая мне неподвластна
                dataGridView1.ColumnCount = reader2.FieldCount;          //  set number of columns in the grid
                string[] row;
                row = new string[reader2.FieldCount];

                for (int j = 0; j < row.Length; j++)       //////  HEADER
                {
                    row[j] = reader2.GetName(j);
                }
                dataGridView1.Rows.Add(row);                // add header
                for (int i = 1; reader2.Read(); i++)      ///////  ALL ROWS
                {
                    strok++;
                    for (int j = 0; j < row.Length; j++)
                    {
                        row[j] = Convert.ToString(reader2.GetValue(j));
                    }

                    dataGridView1.Rows.Add(row);
                    dataGridView1.Rows[i].ReadOnly = true;
                }
                conn2.Close(); // закрываем соединение с БД
                dataGridView1.Show();//и хз что это) 
                textBox1.Text = strok.ToString();
            }


        }

        private void button2_Click(object sender, EventArgs e)
        {
            progressBar1.Maximum = 100; progressBar1.Value = 0; //задание параметров прогрессбара

            //-----------------------------МАГИЯ ЭКСПОРТА С ДАТАГРИД В ЭКСЕЛЬ------------------------------------
            if (this.dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("Нет данных для выгрузки в Excel!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (MessageBox.Show("Выгрузить найденные строки в Excel?", "Вопрос", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;


            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            Microsoft.Office.Interop.Excel.Range ExelWorkRange;

            ExcelWorkBook = ExcelApp.Workbooks.Add();
            ExcelWorkSheet = ExcelWorkBook.Worksheets[1]; //первый по порядку лист в книге Excel

            int RowsCount = this.dataGridView1.RowCount;
            int ColumnsCount = this.dataGridView1.ColumnCount;
            object[,] arrData = new object[RowsCount, ColumnsCount];

            for (int j = 0; j < RowsCount; j++)
                for (int i = 0; i < ColumnsCount; i++)
                    if (j != this.dataGridView1.NewRowIndex)
                    {
                        arrData[j, i] = this.dataGridView1.Rows[j].Cells[i].Value.ToString();
                        progressBar1.Value = 100 * j / strok; progressBar1.Refresh(); //работа с прогрессбаром
                    }


            //выгрузка данных на лист Excel
            ExcelWorkSheet.Range["A2"].Resize[arrData.GetUpperBound(0) + 1, arrData.GetUpperBound(1) + 1].Value = arrData;
            //переносим названия столбцов в Excel файл
            for (int j = 0; j < this.dataGridView1.Columns.Count; j++)
            ExcelWorkSheet.Cells[1, j + 1] = this.dataGridView1.Columns[j].HeaderCell.Value.ToString();

            //украшательство таблицы

            ExcelWorkSheet.Cells[1, 1] = "1";
            ExcelWorkSheet.Cells[1, 2] = "2";
            ExcelWorkSheet.Cells[1, 3] = "3";
            ExcelWorkSheet.Cells[1, 4] = "4";
            ExcelWorkSheet.Cells[1, 5] = "5";
            ExcelWorkSheet.Cells[1, 6] = "6";
            ExcelWorkSheet.Cells[1, 7] = "7";
            ExcelWorkSheet.Cells[1, 8] = "8";
            ExcelWorkSheet.Cells[1, 9] = "9";
            ExcelWorkSheet.Cells[1, 10] = "10";
            ExcelWorkSheet.Cells[1, 11] = "11";
            ExcelWorkSheet.Cells[1, 12] = "12";
            ExcelWorkSheet.Cells[1, 13] = "13";
            ExcelWorkSheet.Cells[1, 14] = "14";
            ExcelWorkSheet.Cells[1, 15] = "15";
            ExcelWorkSheet.Cells[1, 16] = "16";
            ExcelWorkSheet.Rows[1].Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;  // вертикальное выравнивание по центру
            ExcelWorkSheet.Rows[1].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // горизонтальное выравнивание по центру

            ExcelWorkSheet.Cells[1, 1].CurrentRegion.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; //границы
            ExcelWorkSheet.Rows[2].Font.Bold = true;
            ExcelWorkSheet.Columns.AutoFit();//автоподбор ширины ячейки
            ExcelWorkSheet.Columns.WrapText = true;

            //отображаем Excel
            ExcelApp.Visible = true;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            data1 = "2022-08-08"; data2 = "2022-08-08"; data3 = "2022-08-08"; data4 = "2022-08-08";
            control = ""; project = ""; category = ""; data_sozd = ""; data_zakr = ""; strok = 0; 
            job = ""; text_body = ""; text_name = "";
            sql3 = "";

            this.dataGridView1.Rows.Clear(); // очистить датугрид
            textBox1.Text = "";

            comboBox1.SelectedIndex = -1;
            comboBox1.SelectedItem = comboBox1.SelectedIndex;

            comboBox2.SelectedIndex = -1;
            comboBox2.SelectedItem = comboBox2.SelectedIndex;

            comboBox3.SelectedIndex = -1;
            comboBox3.SelectedItem = comboBox3.SelectedIndex;

            comboBox4.SelectedIndex = -1;
            comboBox4.SelectedItem = comboBox4.SelectedIndex;

            comboBox5.SelectedIndex = -1;
            comboBox5.SelectedItem = comboBox5.SelectedIndex;

            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;

            richTextBox1.Text = "";
            textBox3.Text = "";
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.SelectedIndex = -1;
            comboBox2.SelectedItem = comboBox2.SelectedIndex;
            

            if (comboBox1.SelectedIndex != -1)
            {
                sql3 = "select CONCAT(u.last_name, ' ', u.first_name) from acx4_users u join acx4_companies c on u.company_id=c.id where c.name like '%"+ comboBox1.SelectedItem.ToString() + "%' order by u.last_name, u.first_name";//запрос Работники
            textBox2.Text = sql3;

            MySqlConnection conn = new MySqlConnection(connString);// создаём объект для подключения к БД
            conn.Open();// устанавливаем соединение с БД
            MySqlCommand command1 = new MySqlCommand(sql3, conn); // объект для выполнения SQL-запроса
            MySqlDataReader reader1 = command1.ExecuteReader(); // объект для чтения ответа сервера
            while (reader1.Read())
                this.comboBox2.Items.Add(reader1.GetString(0));
            reader1.Close(); // закрываем reader

            conn.Close(); // закрываем соединение с БД
            conn.Dispose();// Уничтожить объект, освободить ресурс.
          }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex != -1)
            {
                job = "and CONCAT(u.last_name, ' ', u.first_name) like '%" + comboBox2.SelectedItem.ToString() + "%'";
            }
        }

    }
}
