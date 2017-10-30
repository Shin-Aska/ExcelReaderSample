using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Project1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            FileDialog fd = new OpenFileDialog();
            fd.Title = "Select a Worksheet";
            fd.Filter = "Excel Files(*.xls;*.xlsx)|*.xls;*.xlsx";
            fd.FilterIndex = 2;
            fd.RestoreDirectory = true;
            DialogResult res = fd.ShowDialog();
            if (res == DialogResult.OK)
            {
                
                ExcelLibrary.EPPLibrary.setFile(fd.FileName);
                
                ExcelLibrary.EPPLibrary.setWorkSheet(1);
                for (int i = 1; true ; i++)
                {
                    
                    try
                    {
                        List<String> row = new List<String>();
                        row.Add((String)ExcelLibrary.EPPLibrary.getValueFrom(i, 1));
                        row.Add((String)ExcelLibrary.EPPLibrary.getValueFrom(i, 2));
                        row.Add((String)ExcelLibrary.EPPLibrary.getValueFrom(i, 3));
                        dataGridView1.Rows.Add(row[0], row[1], row[2]);
                    }
                    catch (Exception ex)
                    {
                        
                        break;
                    }
                }
            }
        }
    }
}
