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
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace LB_1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //private void releaseObject(Microsoft.Office.Interop.Excel.Application xlApp)
        //{
        //    throw new NotImplementedException();
        //}

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Excel 2016(*.xlsx)|*.xlsx";
            ofd.Title = "anketa.xlsx";
            if (ofd.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("Вы не открыли файл анкеты", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string xlFileName = ofd.FileName;
            Microsoft.Office.Interop.Excel.Range Rng;
            Microsoft.Office.Interop.Excel.Workbook xlWB;
            Microsoft.Office.Interop.Excel.Worksheet xlSht;
            int iLastRow, iLastCol;

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWB = xlApp.Workbooks.Open(xlFileName);
            xlSht = xlWB.Worksheets["Лист1"];

            iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Microsoft.Office.Interop.Excel.XlDirection.xlUp].Row;
            iLastCol = xlSht.Cells[1, xlSht.Columns.Count].End[Microsoft.Office.Interop.Excel.XlDirection.xlToLeft].Column;

            Rng = (Microsoft.Office.Interop.Excel.Range)xlSht.Range["A1", xlSht.Cells[iLastRow, iLastCol]];
            var dataArr = (object[,])Rng.Value;

            xlWB.Close(true);
            xlApp.Quit();
            releaseObject(xlSht);
            releaseObject(xlWB);
            releaseObject(xlApp);

            DataTable dt = new DataTable();
            DataRow dtRow;

            for (int i = 1; i <= dataArr.GetUpperBound(1); i++)
                dt.Columns.Add((string)dataArr[1, i]);

            for (int i = 2; i <= dataArr.GetUpperBound(0); i++)
            {
                dtRow = dt.NewRow();
                for (int n = 1; n <= dataArr.GetUpperBound(1); n++)
                {
                    dtRow[n - 1] = dataArr[i, n];
                }
                dt.Rows.Add(dtRow);
            }
            this.dataGridView1.DataSource = dt;
            MessageBox.Show("End", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
