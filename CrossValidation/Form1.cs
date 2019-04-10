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
using Microsoft.Office.Interop.Excel;

namespace CrossValidation
{
    public partial class Form1 : Form
    {
        //Public Variables
        string fileName;
        int columnLenght = 0;
        int rowLenght = 0;
        int counterTest = 0;
        int counterEducation = -1;
        int rNumber;
        List<int> usedNumbers = new List<int>();
        HashSet<int> exclude = new HashSet<int>() { };
        bool runOnceTest = false;
        bool runOnceEducation = false;
        List<int> numberPool = new List<int>();
        int myNumber;
        int missed = 0;
        bool testDataClicked = false;
        int testPercent;
        bool error = true;
        int k;

        public Form1()
        {
            InitializeComponent();
        }

        private int GiveMeANumber()
        {
            var range = Enumerable.Range(0, 100).Where(i => !exclude.Contains(i));
            var rand = new System.Random();
            int index = rand.Next(0, rowLenght - exclude.Count);
            return range.ElementAt(index);
        }

        private void open()
        {
            using (OpenFileDialog openFileDialog1 = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel WorkBook 97-2003|*.xls" })
            {
                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    fileName = openFileDialog1.FileName;
                    this.toolStripStatusLabel1.Text = fileName;
                    toolStripStatusLabel1.Text = "Excel file loaded.";
                }
            }

            //string PathConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";"; //Only Support 97-2003 Excel File
            string PathConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties = \"Excel 12.0 Xml;HDR=YES\"; ";
            OleDbConnection conn = new OleDbConnection(PathConn);
            try
            {
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter("Select * from [" + "Sayfa1" + "$]", conn);
                System.Data.DataTable dt = new System.Data.DataTable();
                myDataAdapter.Fill(dt);
                dataGridView1.DataSource = dt;

                columnLenght = dataGridView1.ColumnCount;
                rowLenght = dataGridView1.RowCount - 1;

                toolStripStatusLabel1.Text = "[TR] File is on view. Ready to proceed.";
                error = false;
            }
            catch (Exception)
            {
                try
                {
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter("Select * from [" + "Sheet1" + "$]", conn);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    myDataAdapter.Fill(dt);
                    dataGridView1.DataSource = dt;

                    columnLenght = dataGridView1.ColumnCount;
                    rowLenght = dataGridView1.RowCount - 1;

                    toolStripStatusLabel1.Text = "[EN] File is on view. Ready to proceed.";
                    error = false;
                }
                catch (Exception)
                {
                    toolStripStatusLabel1.Text = "Wrong language or file format!";
                    error = true;
                }
            }
            if (error == false)
            {
                int labelDataNumber = dataGridView1.RowCount - 1;
                labelData.Text = "Rows:" + " " + labelDataNumber;
                testPercent = Convert.ToInt32(Math.Floor(numericUpDown1.Value * Convert.ToDecimal(dataGridView1.RowCount - 1) / 100));
                if (testPercent < 1)
                {
                    testPercent = 1;
                }
            }
        }

        private void testData()
        {
            if (error == false)
            {
                int labelDataNumber = dataGridView1.RowCount - 1;
                labelData.Text = "Rows:" + " " + labelDataNumber;
                testPercent = Convert.ToInt32(Math.Floor(numericUpDown1.Value * Convert.ToDecimal(dataGridView1.RowCount - 1) / 100));
                if (testPercent < 1)
                {
                    testPercent = 1;
                }
            }

            testDataClicked = true;
            for (int o = 0; o < testPercent; o++)
            {
                counterTest++;
                rNumber = GiveMeANumber();
                exclude.Add(rNumber);
                usedNumbers.Add(rNumber);

                if (runOnceTest == false)
                {
                    int columnCounter = 1;
                    for (int cT = 0; cT < columnLenght; cT++)
                    {
                        dataGridTest.Columns.Add("ColunmName", "Column" + columnCounter);
                        columnCounter++;
                    }
                    runOnceTest = true;
                }
                DataGridViewRow row = (DataGridViewRow)dataGridTest.Rows[counterTest - 1].Clone();
                for (int i = 0; i < columnLenght; i++)
                {
                    var item = dataGridView1.Rows[rNumber].Cells[i].Value;
                    row.Cells[i].Value = item;
                }
                dataGridTest.Rows.Add(row);
                toolStripStatusLabel1.Text = "Test data selection process is done.";
            }

            //This part gets rid of used lines in gridViewTest and gives revised numberPool list.

            for (int i = 0; i <= rowLenght; i++)
            {
                numberPool.Add(i);
            }
            for (int a = 0; a < usedNumbers.Count(); a++)
            {
                for (int b = 0; b < numberPool.Count(); b++)
                {
                    if (numberPool[b] == usedNumbers[a])
                    {
                        numberPool[b] = -1;

                    }
                }
            }
            int labelTestNumber = dataGridTest.RowCount - 1;
            labelTest.Text = "Rows:" + " " + labelTestNumber;
        }

        private void educationData()
        {
            if (testDataClicked == true)
            {
                for (int p = 0; p < rowLenght; p++)
                {
                    counterEducation++;
                    if (runOnceEducation == false)
                    {
                        int columnCounter = 1;
                        for (int cT = 0; cT < columnLenght; cT++)
                        {
                            dataGridEducation.Columns.Add("ColunmName", "Column" + columnCounter);
                            columnCounter++;
                        }
                        runOnceEducation = true;
                    }

                    myNumber = numberPool[counterEducation];
                    if (myNumber == -1)
                    {
                        missed = missed + 1;
                    }
                    else
                    {
                        DataGridViewRow row = (DataGridViewRow)dataGridEducation.Rows[counterEducation - missed].Clone();
                        for (int i = 0; i < columnLenght; i++)
                        {
                            var item = dataGridView1.Rows[myNumber].Cells[i].Value;
                            row.Cells[i].Value = item;
                        }
                        dataGridEducation.Rows.Add(row);
                    }
                }
                toolStripStatusLabel1.Text = "Education data selection process is done.";
            }
            else
            {
                toolStripStatusLabel1.Text = "Please click \"Test Data\" first!";
            }
            int labelEducationNumber = dataGridEducation.RowCount - 1;
            labelEducation.Text = "Rows:" + " " + labelEducationNumber;
        }

        private void export()
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook 
                
                wb1 = excelApp.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet ws1 = (Worksheet)excelApp.ActiveSheet;
            excelApp.Visible = true;

            //Test Data
            for (int i = 1; i < dataGridTest.Columns.Count + 1; i++)
            {
                ws1.Cells[1, i] = dataGridTest.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < dataGridTest.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridTest.Columns.Count; j++)
                {
                    if (dataGridTest.Rows[i].Cells[j].Value == null || dataGridTest.Rows[i].Cells[j].Value == DBNull.Value || String.IsNullOrWhiteSpace(dataGridTest.Rows[i].Cells[j].Value.ToString()))
                    {
                        toolStripStatusLabel1.Text = "There are empty cells in table.";
                    }
                    else
                    {
                        ws1.Cells[i + 2, j + 1] = dataGridTest.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }

            //Education Data
            Workbook wb2 = excelApp.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet ws2 = (Worksheet)excelApp.ActiveSheet;
            for (int i = 1; i < dataGridEducation.Columns.Count + 1; i++)
            {
                ws2.Cells[1, i] = dataGridEducation.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < dataGridEducation.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridEducation.Columns.Count; j++)
                {
                    if (dataGridEducation.Rows[i].Cells[j].Value == null || dataGridEducation.Rows[i].Cells[j].Value == DBNull.Value || String.IsNullOrWhiteSpace(dataGridEducation.Rows[i].Cells[j].Value.ToString()))
                    {
                        toolStripStatusLabel1.Text = "There are empty cells in table.";
                    }
                    else
                    {
                        ws2.Cells[i + 2, j + 1] = dataGridEducation.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            open();
        }

        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            export();
        }

        private void automaticToolStripMenuItem_Click(object sender, EventArgs e)
        {
            k = Convert.ToInt32(numericUpDownK.Value);
            if (error == false)
            {
                for (int i = 0; i < k; i++)
                {
                    dataGridTest.Rows.Clear();
                    dataGridTest.Refresh();
                    dataGridEducation.Rows.Clear();
                    dataGridEducation.Refresh();
                    counterTest = 0;
                    counterEducation = -1;
                    missed = 0;

                    exclude.Clear();
                    usedNumbers.Clear();
                    numberPool.Clear();

                    testData();
                    educationData();
                    export();
                    
                }
            }
        }

        private void testDataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridTest.Rows.Clear();
            dataGridTest.Refresh();
            dataGridEducation.Rows.Clear();
            dataGridEducation.Refresh();
            counterTest = 0;
            counterEducation = -1;
            missed = 0;

            exclude.Clear();
            usedNumbers.Clear();
            numberPool.Clear();

            testData();
        }

        private void educationDataToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            educationData();
        }
    }
}

