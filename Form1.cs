using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Data.OleDb;
namespace GHIAProj
{
    public partial class Form1 : Form
    {
        private int waterObjType1 = 1;
        private int waterObjType2 = 6;
        private bool isTransgr = false;
        private string header = "Река";
        public Form1()
        {
            InitializeComponent();
        }
        private void CleAR()
        {
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            CleAR();
            TableCreater tb = new TableCreater();
            DataTable table = new DataTable();
            if(!isTransgr)
            {
                if (radioButton1.Checked == true)//Hight level
                {
                    table = tb.GetHight(Operations.HIGHPOLLUTION, waterObjType1, waterObjType2, dateTimePicker1.Value, dateTimePicker2.Value);
                }
                else if (radioButton2.Checked == true)//Extreme level
                {
                    table = tb.GetHight(Operations.EXTREMEPOLLUTION, waterObjType1, waterObjType2, dateTimePicker1.Value, dateTimePicker2.Value);
                }
                if (table.Rows.Count != 0)
                {
                    FillGrid(dataGridView1, textBox1, table);
                    dataGridView1.Columns[0].HeaderText = header + " - Пункт";
                    dataGridView1.Columns[1].HeaderText = "Речной бассейн";
                    dataGridView1.Columns[2].HeaderText = "Ингредиент";
                    dataGridView1.Columns[3].HeaderText = "Число"+ "\n"+ "случаев"+"\n"+ "ВЗ";
                    dataGridView1.Columns[4].HeaderText = "ПДК";
                    dataGridView1.Columns[5].HeaderText = "Дата";
                    dataGridView1.Columns[6].HeaderText = "Источники загрязнения";
                    dataGridView1.Columns[7].HeaderText = "Субъект"+ "\n"+ "Российской Федерации";
                }
                else MessageBox.Show("В базе нет данных");
                GridFormator form = new GridFormator();
                form.CommonSetCellsToNull(ref dataGridView1);           
            }
            else
            {
                if (radioButton1.Checked == true)//Hight level
                {
                    table = tb.GetHight(Operations.BORDER_HIGHPOLLUTION, waterObjType1, waterObjType2, dateTimePicker1.Value, dateTimePicker2.Value);
                }
                else if (radioButton2.Checked == true)//Extreme level
                {
                    table = tb.GetHight(Operations.BORDER_EXTREMEPOLLUTION, waterObjType1, waterObjType2, dateTimePicker1.Value, dateTimePicker2.Value);
                }
                if (table.Rows.Count != 0)
                {
                    tb.CreateDataColumn(ref table);
                    FillGrid(dataGridView1, textBox1, table);

                    dataGridView1.Columns[0].HeaderText = "Сопредельное государство";
                    dataGridView1.Columns[1].HeaderText = header + " - Пункт";
                    dataGridView1.Columns[2].HeaderText = "Ингредиент";
                    dataGridView1.Columns[3].HeaderText = "Число случаев ВЗ";
                    dataGridView1.Columns[4].HeaderText = "ПДК";
                    dataGridView1.Columns[5].HeaderText = "Дата";
                }
                else MessageBox.Show("В базе нет данных");
                GridFormator form = new GridFormator();
                form.TransgrSetCellsToNull(ref dataGridView1);
            }
        }
        private void FillGrid(DataGridView dataGridView, TextBox textBox, DataTable source)
        {
            try
            {
                dataGridView.DataSource = source;
                textBox.Text = "Соединиение успешно, количество строк= " + Convert.ToString(Convert.ToInt32(dataGridView.RowCount) - 1);
            }
            catch
            {
                textBox.Text = "Не удалось подключиться к базе данных";
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "export.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Exporter exporter = new Exporter(isTransgr);
                exporter.ToWord(dataGridView1, sfd.FileName);
            }
        }
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            waterObjType1 = 1;
            waterObjType2 = 6;
            header = "Река";
        }
        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            waterObjType1 = 7;
            waterObjType2 = 8;
            header = "Водоем";
        }
        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            isTransgr = false;
        }
        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            isTransgr = true;
        }
        private void GetTableFromExel()
        {
            OpenFileDialog opf = new OpenFileDialog();
            opf.Filter = "Excel (*.XLS)|*.XLS";
            opf.ShowDialog();
            DataTable tb = new DataTable();
            string filename = opf.FileName;
            string ConStr = String.Format("Provider=Microsoft.Jet.OLEDB.4.0; Data Source={0}; Extended Properties=Excel 8.0;", filename);
            System.Data.DataSet ds = new System.Data.DataSet("EXCEL");
            OleDbConnection cn = new OleDbConnection(ConStr);
            cn.Open();
            DataTable schemaTable = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];
            string select = String.Format("SELECT * FROM [{0}]", sheet1);
            OleDbDataAdapter ad = new OleDbDataAdapter(select, cn);
            ad.Fill(ds);
            DataTable tb1 = ds.Tables[0];
            cn.Close();
            dataGridView1.DataSource = tb1;
        }
    }
}