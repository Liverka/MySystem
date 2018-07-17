using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Система
{
    public partial class Form2 : Form
    {
        public Form2(DataGridView data)
        {
            InitializeComponent();
            this.data = data;
        }

        DataGridView data;

        private void Form2_Load(object sender, EventArgs e)
        {
            double Ymin = Convert.ToDouble(data[3, 0].Value);
            double Ymax = Convert.ToDouble(data[4, data.Rows.Count-1].Value);

            int count = data.Rows.Count;
            for (int i = 0; i < data.Rows.Count - 1; ++i)
            {
                if (Ymin > Convert.ToDouble(data[3, i].Value))
                    Ymin = Convert.ToDouble(data[3, i].Value);
                if (Ymax < Convert.ToDouble(data[4, i].Value))
                    Ymax = Convert.ToDouble(data[4, i].Value);
            }

            int massSize = 1;
            for (int i = 0; i < data.Rows.Count - 1; ++i)
            {
                double pointY = Convert.ToDouble(data[3, i].Value);
                double steps = (Convert.ToDouble(data[4, i].Value) - Convert.ToDouble(data[3, i].Value)) / Convert.ToDouble(data[2, i].Value);
                int step = 0;
                while (step < steps)
                {
                    massSize++;
                    step++;
                }
            }

            double[] x = new double[massSize];
            double[] y = new double[massSize];
            
            int c = 1;
            int xc = 1;
            y[0] = Convert.ToDouble(data[3, 0].Value);

            for (int i = 0; i < data.Rows.Count - 1; ++i)
            {
                double pointY = Convert.ToDouble(data[3, i].Value);
                double steps = (Convert.ToDouble(data[4, i].Value) - Convert.ToDouble(data[3, i].Value)) / Convert.ToDouble(data[2, i].Value);
                int step = 0;                
                while (step < steps)
                {                    
                    x[c] = xc;
                    y[c] = y[c-1] + Convert.ToDouble(data[2, i].Value);
                    c++;
                    step++;
                    xc++;
                }                  
                                
            }

            chart1.ChartAreas[0].AxisX.Title = "время, мин.";
            chart1.ChartAreas[0].AxisY.Title = "температура, °С";
            chart1.ChartAreas[0].AxisX.TitleFont = new Font("Times New Roman", 14, FontStyle.Bold);
            chart1.ChartAreas[0].AxisY.TitleFont = new Font("Times New Roman", 14, FontStyle.Bold);
            chart1.ChartAreas[0].AxisY.Minimum = Ymin;
            chart1.ChartAreas[0].AxisY.Maximum = Ymax;                     
            chart1.ChartAreas[0].AxisX.Minimum = 0;            
            chart1.ChartAreas[0].AxisX.MajorGrid.Interval = 1;
            chart1.ChartAreas[0].AxisX.Interval = 1;
            chart1.ChartAreas[0].AxisY.Interval = 10;
            chart1.Series[0].BorderWidth = 3;
            chart1.Series[0].Points.DataBindXY(x, y);
        }
    }
}
