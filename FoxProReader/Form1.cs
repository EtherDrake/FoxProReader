using FoxProReader.Classes;
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace FoxProReader
{
    public partial class Form1 : Form
    {
        string f1, f2, FilePath;

        bool f1IsLoaded=false, f2IsLoaded=false;

        public delegate void Loaded(string filename, int index);
        public event Loaded onLoaded;

        public Form1()
        {
            InitializeComponent();
            onLoaded += LoadColumnNames;
        }        

        private void button1_Click(object sender, EventArgs e)
        {
            if(openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                f1 = openFileDialog1.FileName;                
                onLoaded(f1, 1);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                f2 = openFileDialog1.FileName;
                onLoaded(f2, 2);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!f1IsLoaded || !f2IsLoaded)
            {
                MessageBox.Show("Ви не завантажили обидва файла!");
                return;
            }

            DataSet dataSet = new DataSet();      

            try
            {
                string Sql = "SELECT ";
                foreach (object Item in checkedListBox1.CheckedItems)
                    Sql += "t1." + Item.ToString() + ", ";

                foreach (object Item in checkedListBox2.CheckedItems)
                    Sql += "t2." + Item.ToString() + ", ";

                Sql = Sql.Remove(Sql.Length - 2);

                Sql += " FROM " + f1 + " t1 INNER JOIN " + f2 + " t2 ON t1.KDMO=t2.KDMO";

                DBFreader.Get(Sql, ref dataSet, FilePath, "result");
                dataGridView1.DataSource = dataSet.Tables[0];

                saveFileDialog1.Filter = "Excel Files (*.xls)|*.xls";
                saveFileDialog1.DefaultExt = "xls";

                if (saveFileDialog1.ShowDialog()==DialogResult.OK)
                    MyExtensions.ExportToExcel(dataSet.Tables[0], saveFileDialog1.FileName);

            }
            catch (Exception ex)
            {
                DialogResult dialogResult = MessageBox.Show(ex.Message, "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void LoadColumnNames(string filename, int index)
        {
            if (index == 1)
            {
                setColumnList(checkedListBox1, filename);
                f1IsLoaded = true;
            }

            if (index == 2)
            {
                setColumnList(checkedListBox2, filename);
                f2IsLoaded = true;
            }
        }

        private void setColumnList(CheckedListBox list, string filename)
        {
            DataSet dataSet = new DataSet();
            FilePath = Path.GetDirectoryName(filename);
            string file = Path.GetFileName(filename);
            string Sql = "SELECT TOP 1 t1.* FROM " + file + " t1";
            DBFreader.Get(Sql, ref dataSet, FilePath, "MONOKD");
            foreach (DataColumn column in dataSet.Tables[0].Columns)
            {
                list.Items.Add(column.ColumnName);
            }
        }
    }
}
