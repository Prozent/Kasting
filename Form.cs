using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data.OleDb;

namespace WindowsFormsApp1
{
    public partial class Form : System.Windows.Forms.Form
    {
        public Form()
        {
            InitializeComponent();
        }



        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Browse Text Files";

            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;

            openFileDialog1.DefaultExt = "txt";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            openFileDialog1.ReadOnlyChecked = true;
            openFileDialog1.ShowReadOnly = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                DataTable tb = new DataTable();
                string filename = openFileDialog1.FileName;


                string ConStr = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;extended properties=\"excel 8.0;hdr=no;IMEX=1\";data source={0}",
                            filename);
                DataSet ds = new DataSet("EXCEL");
                OleDbConnection cn = new OleDbConnection(ConStr);
                cn.Open();
                DataTable schemaTable = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];
                string select = String.Format("SELECT * FROM [{0}]", sheet1);
                OleDbDataAdapter ad = new OleDbDataAdapter(select, cn);
                ad.Fill(ds);
                tb = ds.Tables[0];
                cn.Close();
                dataGridView1.DataSource = tb;
                for (int count_dg = 0; count_dg < dataGridView1.Columns.Count; count_dg++)
                    dataGridView1.Columns[count_dg].HeaderText = dataGridView1[count_dg, 0].Value.ToString();
                dataGridView1.Rows.Remove(dataGridView1.Rows[0]);
                textBox1.Text = dataGridView1.Columns.Count.ToString();//колли чество столбцов
                textBox2.Text = dataGridView1.RowCount.ToString();//колличество строк
            }

        }
        static public string[,] lines; //публичная переменная для второй формы(на будущее)
        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();

            int summ = Convert.ToInt32(textBox1.Text) * Convert.ToInt32(textBox2.Text);
            lines = new string[2, dataGridView1.RowCount];
            for (int i = 0; i < 2; i++)
                for (int j = 0; j < dataGridView1.RowCount; j++)
                {
                    if (dataGridView1[i, j].Value != null)
                        lines[i, j] = dataGridView1[i, j].Value.ToString();
                }
            
            dataGridView2.ColumnCount = Form.lines.GetUpperBound(0)+1 ;
            dataGridView2.RowCount = Form.lines.GetUpperBound(1)+ summ+1;
            bool flag_grup=true;
            for (int i = 0; flag_grup && i < dataGridView2.ColumnCount; i++)
            {
                int cmehenie = 0;
                for (int j = 0; flag_grup && j < lines.GetUpperBound(1); j++)
                {
                    if ((summ % 2) == 0)
                    {
                        if (j % ((Convert.ToInt32(Form.lines.GetUpperBound(1)) / summ) + 1) == 0)
                            
                            ++cmehenie;
                    }
                    else if (Convert.ToInt32(Form.lines.GetUpperBound(1)) / summ == 1) {
                        MessageBox.Show("вы хотите разделить на группы по " +
                                        "1му человеку, это не целесообразно, ");
                        flag_grup = false;//выход с довйного цикла
                    }
                    else
                    {
                        if (j % ((Convert.ToInt32(Form.lines.GetUpperBound(1)) / summ)) == 0)
                            ++cmehenie;
                    }
                    textBox1.Text = cmehenie.ToString();
                    textBox2.Text = j.ToString();
                    dataGridView2[i, j + cmehenie].Value = lines[i, j].ToString();
                    
                }
            
            }
            




        }
    }
}
