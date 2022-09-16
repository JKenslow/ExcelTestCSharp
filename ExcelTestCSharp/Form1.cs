using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelTestCSharp
{
    public partial class Form1 : Form
    {

        private XL sheet;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fileName = openFileDialog1.FileName;
                sheet = new XL(fileName);

                dataGridView1.ColumnCount = 10;
                for (int i = 0; i < 10; i++)
                {
                    string[] row = new string[10];
                    for (int j = 0; j < 10; j++)
                    {
                        row[j] = sheet.GetCell(i + 1, j + 1);
                    }
                    dataGridView1.Columns[i].Name = ((char)(i+65)).ToString();
                    dataGridView1.Rows.Add(row);
                    dataGridView1.Columns[i].DisplayIndex = i;
                }

                for(int k = 0; k < 10; k++)
                {
                    dataGridView1.Rows[k].HeaderCell.Value = (k + 1).ToString();
                }
                dataGridView1.AutoResizeColumns();
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fileName = saveFileDialog1.FileName;
                
            }
        }
    }
}
