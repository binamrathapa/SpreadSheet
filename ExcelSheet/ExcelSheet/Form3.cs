using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SpreadSheet
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            dataGridView1.RowPostPaint += new DataGridViewRowPostPaintEventHandler(this.dgvUserDetails_RowPostPaint);
            //dataGridView1.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(this.dataGridView1_EditingControlShowing);
            // dataGridView1.ce

            for (char c = 'A'; c <= 'Z'; c++)
            {
                DataGridViewTextBoxColumn cell = new DataGridViewTextBoxColumn();
                cell.Name = c.ToString();
                cell.HeaderText = c.ToString();
                dataGridView1.Columns.Add(cell);
            }
            for (int row = 0; row < 25; row++)
            {
                DataGridViewRow dr = new DataGridViewRow();

                dataGridView1.Rows.Add(dr);

            }
        }

        private void dgvUserDetails_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(dataGridView1.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }
        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //DataGridViewTextBoxEditingControl tb = (DataGridViewTextBoxEditingControl)e.Control;
            // tb.KeyPress += new KeyPressEventHandler(dataGridViewTextBox_KeyPress);
            //tb.Enter += Tb_Enter;
            //e.Control.KeyPress += new KeyPressEventHandler(dataGridViewTextBox_KeyPress);
            //TextBox txtbox = e.Control as TextBox;
            //if (txtbox != null)
            //{
            //    txtbox.KeyPress += new KeyPressEventHandler(dataGridViewTextBox_KeyPress);
            //}
        }

        private void Tb_Enter(object sender, EventArgs e)
        {
            //if(e.KeyChar == (char)Keys.Enter)
            //{
            //    MessageBox.Show("You press Enter");
            //}
        }

        private void dataGridViewTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            var v = sender;

            //when i press enter,bellow code never run?
            if (e.KeyChar == (char)Keys.Enter)
            {
                MessageBox.Show("You press Enter");
            }
        }

        private void dataGridView1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {

            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dv = (DataGridView)sender;
            string text = Convert.ToString(dv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
            if (text.StartsWith("="))
            {
                if (text.Contains(":"))
                {
                    string operationName = text.Substring(1, (text.IndexOf(' ') - text.IndexOf('='))-1);
                    string lowerValue = text.Substring(text.IndexOf(' ') + 1, (text.IndexOf(':') - text.IndexOf(' ')) - 1);
                    string higherValue = text.Substring(text.IndexOf(':') + 1, (text.Length - 1) - text.IndexOf(':'));

                    char lowerLetter = Convert.ToChar(lowerValue.Substring(0, 1));
                    char higherLetter = Convert.ToChar(higherValue.Substring(0, 1));

                    string lowerNumber = lowerValue.Length < 3 ? lowerValue.Substring(1, 1) : lowerValue.Substring(1, 2);
                    string higherNumber = higherValue.Length < 3 ? higherValue.Substring(1, 1) : higherValue.Substring(1, 2);
                    List<double> values = new List<double>();
                    if (lowerLetter == higherLetter)
                    {
                        //int sum = 0;
                        
                        for (int row = Convert.ToInt32(lowerNumber); row <= Convert.ToInt32(higherNumber); row++)
                        {
                            values.Add(Convert.ToDouble( dv.Rows[row - 1].Cells[e.ColumnIndex].Value));
                            //sum += Convert.ToInt32(dv.Rows[row - 1].Cells[e.ColumnIndex].Value);
                        }

                        dv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = new ExcellFormulaCalculation().GetCalculatedValue(operationName, values);//sum.ToString();
                    }
                    else if (lowerNumber == higherNumber)
                    {
                        int sum = 0;
                        
                        for (char c = lowerLetter; c <= higherLetter; c++)
                        {
                            values.Add(Convert.ToDouble(dv.Rows[e.RowIndex].Cells[c.ToString()].Value));
                            //sum += Convert.ToInt32(dv.Rows[e.RowIndex].Cells[c.ToString()].Value);
                        }
                        dv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = new ExcellFormulaCalculation().GetCalculatedValue(operationName, values);//sum.ToString();
                    }

                    else
                    {
                        //MessageBox.Show("format is invalid");
                        //different calculation example =MEAN A1:B4
                    }
                }
                else
                {
                    var calculationType = new System.Text.RegularExpressions.Regex(@"([\*]{1}|[\+]{1}|[\-]{1})|[\/]{1}");
                    var matchPattern = calculationType.Match(text);

                    if (matchPattern.Index > 0)
                    {
                        string actualCalculation = text.Substring(1);
                        string[] numbers = actualCalculation.Split(Convert.ToChar(matchPattern.Value));

                        int lowerValue =Convert.ToInt32( numbers[0].Substring(1, numbers[0].Length-1));
                        int higherValue =Convert.ToInt32( numbers[1].Substring(1, numbers[1].Length-1));

                        char lowerLetter = Convert.ToChar(numbers[0].Substring(0, 1));
                        char higherLetter = Convert.ToChar(numbers[1].Substring(0, 1));

                        //string lowerNumber = lowerValue.Length < 3 ? lowerValue.Substring(1, 1) : lowerValue.Substring(1, 2);
                        //string higherNumber = higherValue.Length < 3 ? higherValue.Substring(1, 1) : higherValue.Substring(1, 2);
                       // dv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = dv.Rows[lowerValue - 1].Cells[lowerLetter].Value;

                    }
                }

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
                //DataGridViewSelectedRowCollection rows = dataGridView1.SelectedRows;
                //string val = (string)rows[2].Cells["Late_Time"].Value;
                //DataGridViewCell cell = dataGridView1.SelectedCells[0] as DataGridViewCell;
                //string value = cell.Value.ToString();
                List<int> values = new List<int>();
                List<int> rowIndexes = new List<int>();
                List<int> columnIndexes = new List<int>();
                foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
                {
                    var value = dataGridView1.Rows[cell.RowIndex].Cells[cell.ColumnIndex].Value;
                    values.Add(Convert.ToInt32(value));
                    rowIndexes.Add(cell.RowIndex);
                    columnIndexes.Add(cell.ColumnIndex);
                }

                ArrayList chars = new ArrayList();
                List<int> uniqueValues = new List<int>();

                int valueIndex = 0;
                foreach (int s in columnIndexes)
                {
                    int charIndex = 0;
                    for (char c = 'A'; c <= 'Z'; c++)
                    {
                        if (charIndex == s)
                        {
                            if (!chars.Contains(c))
                            {
                                uniqueValues.Add(values[valueIndex]);
                                chars.Add(c);
                            }
                            else
                            {
                                uniqueValues[chars.IndexOf(c)] += values[valueIndex];
                            }
                        }
                        charIndex++;

                    }
                    valueIndex++;

                }
                string[] letters = new string[chars.Count];
                int index = 0;
                foreach (char c in chars)
                {
                    letters[index] = c.ToString();
                    index++;
                }

                int[] dataValues = new int[uniqueValues.Count];
                int arrIndex = 0;
                foreach (int val in uniqueValues)
                {
                    dataValues[arrIndex] = val;
                    arrIndex++;
                }

            frmBarChart f = new frmBarChart(letters, dataValues);
                f.Show();

            }
        }
    }

