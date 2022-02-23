using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace Zoo
{
    public partial class Form1 : Form
    {

        OleDbConnection cn;
        DataTable tbls;
        OleDbDataAdapter dbAdapter1;
        DataTable tb;
        OleDbCommandBuilder cb;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Directory.GetCurrentDirectory() + @"\Зоопарк.accdb");
            cn.Open();
            tbls = cn.GetSchema("Tables", new string[] { null, null, null, "TABLE" }); //список всех таблиц
            foreach (DataRow row in tbls.Rows)
            {
                string TableName = row["TABLE_NAME"].ToString();
                comboBox1.Items.Add(TableName);
            };

        }

        void button2_change(bool a)
        {
            if (a == true) { button2.Enabled = true; button8.Enabled = true; }
            else { button2.Enabled = false; button8.Enabled = false; }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button2_change(true);
            if (comboBox1.SelectedIndex != -1)
            {
                dbAdapter1 = new OleDbDataAdapter(@"SELECT * FROM " + comboBox1.SelectedItem, cn);
                tb = new DataTable();
                dbAdapter1.Fill(tb);
                dataGridView1.DataSource = tb;
            }
            else
            {
                MessageBox.Show("Не выбрана таблица для запроса");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            cb = new OleDbCommandBuilder(dbAdapter1);
            dbAdapter1.Update(tb);
            MessageBox.Show("Сохраненно");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            button2_change(false);
            dbAdapter1 = new OleDbDataAdapter(@"SELECT TOP 5 Зоопарк.СтоимостьБилета, Зоопарк.НазваниеЗоопарка FROM Зоопарк ORDER BY Зоопарк.СтоимостьБилета DESC;", cn);
            tb = new DataTable();
            dbAdapter1.Fill(tb);
            dataGridView1.DataSource = tb;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button2_change(false);
            dbAdapter1 = new OleDbDataAdapter(@"
            SELECT TOP 5 Sum(Наличие.Количество) AS Количество, A.НазваниеЗоопарка
            FROM (SELECT Зоопарк.КодЗоопарка, Зоопарк.НазваниеЗоопарка, Наличие.Количество
            FROM Зоопарк INNER JOIN (Животное INNER JOIN Наличие ON Животное.КодЖивотного = Наличие.КодЖивотного) ON Зоопарк.КодЗоопарка = Наличие.КодЗоопарка)  AS A
            GROUP BY A.НазваниеЗоопарка, A.Зоопарк.КодЗоопарка
            ORDER BY Sum(Наличие.Количество) DESC;"
            , cn);
            tb = new DataTable();
            dbAdapter1.Fill(tb);
            dataGridView1.DataSource = tb;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            button2_change(false);
            dbAdapter1 = new OleDbDataAdapter(@"
            SELECT TOP 5 COUNT(Наличие.Количество) AS Количество, [A].НазваниеЗоопарка
            FROM (SELECT Зоопарк.КодЗоопарка, Зоопарк.НазваниеЗоопарка, Наличие.Количество
            FROM Зоопарк INNER JOIN (Животное INNER JOIN Наличие ON Животное.КодЖивотного = Наличие.КодЖивотного) ON Зоопарк.КодЗоопарка = Наличие.КодЗоопарка)  AS [A]
            GROUP BY [A].Зоопарк.КодЗоопарка, [A].НазваниеЗоопарка
            ORDER BY COUNT(Наличие.Количество) DESC;"
            , cn);
            tb = new DataTable();
            dbAdapter1.Fill(tb);
            dataGridView1.DataSource = tb;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            button2_change(false);
            dbAdapter1 = new OleDbDataAdapter(@"
            SELECT Наличие.Количество, Животное.НазваниеЖивотного, Зоопарк.НазваниеЗоопарка
            FROM Животное INNER JOIN (Зоопарк INNER JOIN Наличие ON Зоопарк.КодЗоопарка = Наличие.КодЗоопарка) ON Животное.КодЖивотного = Наличие.КодЖивотного
            ORDER BY Зоопарк.НазваниеЗоопарка;"
            , cn);
            tb = new DataTable();
            dbAdapter1.Fill(tb);
            dataGridView1.DataSource = tb;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            button2_change(false);
            dbAdapter1 = new OleDbDataAdapter(@"
            SELECT COUNT(Наличие.Количество) AS Количество, [A].НазваниеЗоопарка
            FROM (SELECT Зоопарк.КодЗоопарка, Зоопарк.НазваниеЗоопарка, Наличие.Количество
            FROM Зоопарк INNER JOIN (Животное INNER JOIN Наличие ON Животное.КодЖивотного = Наличие.КодЖивотного) ON Зоопарк.КодЗоопарка = Наличие.КодЗоопарка)  AS [A]
            GROUP BY [A].Зоопарк.КодЗоопарка, [A].НазваниеЗоопарка"
            , cn);
            tb = new DataTable();
            dbAdapter1.Fill(tb);
            dataGridView1.DataSource = tb;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            int rowIndex = dataGridView1.CurrentCell.RowIndex;
            dataGridView1.Rows.RemoveAt(rowIndex);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/kqkarick");
        }

        private void pictureBox1_MouseHover(object sender, EventArgs e)
        {
            pictureBox1.Cursor = Cursors.Hand;
        }

        private void pictureBox1_MouseLeave(object sender, EventArgs e)
        {
            pictureBox1.Cursor = Cursors.Default;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/kqkarick");
        }
    }
}
