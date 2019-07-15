using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace SpreadSheet
{
    public partial class frmBarChart : Form
    {
        string[] chars;
        int[] values;
        public frmBarChart(string[] chars, int[] values)
        {
            this.chars = chars;
            this.values = values;
            InitializeComponent();
        }

        

        public void BarExample()
        {
            chart1.Series.Clear();

            // Data arrays
            string[] seriesArray = chars; // { "A", "B", "C", "D" };
            int[] pointsArray = values;//{ 2, 1, 7, 5 };

            // Set palette
            chart1.Palette = ChartColorPalette.EarthTones;

            // Set title
            chart1.Titles.Add("Spread Sheet");

            // Add series.
            for (int i = 0; i < seriesArray.Length; i++)
            {
                Series series = chart1.Series.Add(seriesArray[i]);
                series.Points.Add(pointsArray[i]);
            }
        }

        private void frmBarChart_Load(object sender, EventArgs e)
        {
            BarExample();
        }
    }
}
