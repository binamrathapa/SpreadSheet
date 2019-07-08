using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using System.Windows.Forms.DataVisualization.Charting;

namespace SpreadSheet
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            for (int row = 0; row < 26; row++)
            {
                int column = 0;
                for (char c = 'A'; c <= 'Z'; c++)
                {
                    if (row == 0)
                    {
                        Label lc = new Label();
                        lc.Text = c.ToString();
                        lc.Location = new Point(50 + column * 100, 10);
                        panel1.Controls.Add(lc);
                    }
                    if (column == 0)
                    {
                        Label ln = new Label();
                        ln.Width = 20;
                        ln.Text = (row + 1).ToString();
                        ln.Location = new Point(20, 45 + row * 20);
                        panel1.Controls.Add(ln);
                    }

                    TextBox tb = new TextBox();
                    tb.Name = c.ToString() + (row + 1).ToString();
                    tb.Location = new Point(40 + column * 100, 40 + row * 20);
                    tb.Width = 100;
                    tb.Height = 20;
                    tb.KeyPress += Tb_KeyPress;
                    panel1.Controls.Add(tb);
                    column++;
                }
            }
        }

        private void Tb_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (e.KeyChar == (char)13)
                {
                    TextBox tb = (TextBox)sender;
                    string text = tb.Text;
                    if (text.StartsWith("="))
                    {
                        string lowerValue = text.Substring(text.IndexOf(' ') + 1, (text.IndexOf(':') - text.IndexOf(' ')) - 1);
                        string higherValue = text.Substring(text.IndexOf(':') + 1, (text.Length - 1) - text.IndexOf(':'));

                        char lowerLetter = Convert.ToChar(lowerValue.Substring(0, 1));
                        char higherLetter = Convert.ToChar(higherValue.Substring(0, 1));

                        string lowerNumber = lowerValue.Length < 3 ? lowerValue.Substring(1, 1) : lowerValue.Substring(1, 2);
                        string higherNumber = higherValue.Length < 3 ? higherValue.Substring(1, 1) : higherValue.Substring(1, 2);
                        if (lowerLetter == higherLetter)
                        {
                            int sum = 0;
                            for (int row = Convert.ToInt32(lowerNumber); row <= Convert.ToInt32(higherNumber); row++)
                            {
                                TextBox v = (TextBox)panel1.Controls.Find("A" + row, false)[0];
                                sum += Convert.ToInt32(v.Text);
                            }

                            tb.Text = sum.ToString();
                        }
                        else if (lowerNumber == higherNumber)
                        {
                            int sum = 0;

                            for (char c = lowerLetter; c <= higherLetter; c++)
                            {
                                TextBox tbTarget = (TextBox)panel1.Controls.Find(c.ToString() + lowerNumber, false)[0];
                                sum += Convert.ToInt32(tbTarget.Text);
                            }
                            tb.Text = sum.ToString();
                        }
                        else
                        {
                           // MessageBox.Show("format is invalid");
                           //different calculation
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void GenerateChart()
        {
           /* int[] arr = { 7,12,18,25,23};
            Chart chart = new Chart();
            chart.Location = new Point(10, 650);
            chart.Series.Add("duck");
            chart.Series[0].LegendText = "Brazil Order Statistics";
            chart.Series[0].ChartType = SeriesChartType.Bar;
            chart.Series[0].IsValueShownAsLabel = true;
            foreach (int i in arr)
            {
                chart.Series[0].Points.AddY(i);
            }
            panel1.Controls.Add(chart);*/
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GenerateChart();
        }
    }
}
