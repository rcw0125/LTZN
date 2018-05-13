using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace LTZN.查询
{
    public partial class 基准风温 : Form
    {
        public double 基准风温1;
        public double 基准风温2;
        public double 基准风温3;
        public double 基准风温4;
        public double 基准风温5;
        public double 基准风温6;

        public 基准风温()
        {
            InitializeComponent();
        }

        private void 基准风温_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            基准风温1 = Convert.ToDouble(textBox1.Text);
            基准风温2 = Convert.ToDouble(textBox2.Text);
            基准风温3 = Convert.ToDouble(textBox3.Text);
            基准风温4 = Convert.ToDouble(textBox4.Text);
            基准风温5 = Convert.ToDouble(textBox5.Text);
            基准风温6 = Convert.ToDouble(textBox6.Text);
        }
    }
}