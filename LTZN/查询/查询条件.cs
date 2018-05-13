using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace LTZN.查询
{
    public partial class 查询条件 : Form
    {
        public string tiaojian = "";
        public 查询条件()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string tiaojian1 = "";
            string tiaojian2 = "";
            string tiaojian3 = "";
            string tiaojian4 = "";
            string tiaojian5 = "";
            string tiaojian6 = "";

            if (this.c1NumericEdit1.Value != DBNull.Value)
            {
                switch (this.guanxi1.Text)
                {
                    case "大于":
                        tiaojian1 = " FESI>" + c1NumericEdit1.Value;
                        break;
                    case "等于":
                        tiaojian1 = " FESI=" + c1NumericEdit1.Value;
                        break;
                    case "小于":
                        tiaojian1 = " FESI<" + c1NumericEdit1.Value;
                        break;
                }
            }

            if (this.c1NumericEdit2.Value != DBNull.Value)
            {
                switch (this.guanxi2.Text)
                {
                    case "大于":
                        tiaojian2 = " FESI>" + c1NumericEdit2.Value;
                        break;
                    case "等于":
                        tiaojian2 = " FESI=" + c1NumericEdit2.Value;
                        break;
                    case "小于":
                        tiaojian2 = " FESI<" + c1NumericEdit2.Value;
                        break;
                }
            }

            if (this.c1NumericEdit3.Value != DBNull.Value)
            {
                switch (this.guanxi3.Text)
                {
                    case "大于":
                        tiaojian3 = " FES>" + c1NumericEdit3.Value;
                        break;
                    case "等于":
                        tiaojian3 = " FES=" + c1NumericEdit3.Value;
                        break;
                    case "小于":
                        tiaojian3 = " FES<" + c1NumericEdit3.Value;
                        break;
                }
            }

            if (this.c1NumericEdit4.Value != DBNull.Value)
            {
                switch (this.guanxi4.Text)
                {
                    case "大于":
                        tiaojian4 = " FES>" + c1NumericEdit4.Value;
                        break;
                    case "等于":
                        tiaojian4 = " FES=" + c1NumericEdit4.Value;
                        break;
                    case "小于":
                        tiaojian4 = " FES<" + c1NumericEdit4.Value;
                        break;
                }
            }

            string guanxi = "";
            switch (this.guanxi5.Text)
            {
                case "并且":
                    guanxi = " and ";
                    break;
                case "或者":
                    guanxi = " or ";
                    break;
            }

            if (!String.IsNullOrEmpty(tiaojian1) && !String.IsNullOrEmpty(tiaojian2))
                tiaojian5 = "(" + tiaojian1 + guanxi + tiaojian2 + ")";
            else
            {
                if (!String.IsNullOrEmpty(tiaojian1)) tiaojian5 = tiaojian1;
                if (!String.IsNullOrEmpty(tiaojian2)) tiaojian5 = tiaojian2;
            }


            switch (this.guanxi7.Text)
            {
                case "并且":
                    guanxi = " and ";
                    break;
                case "或者":
                    guanxi = " or ";
                    break;
            }
            if (!String.IsNullOrEmpty(tiaojian3) && !String.IsNullOrEmpty(tiaojian4))
                tiaojian6 = "(" + tiaojian3 + guanxi + tiaojian4 + ")";
            else
            {
                if (!String.IsNullOrEmpty(tiaojian3)) tiaojian6 = tiaojian3;
                if (!String.IsNullOrEmpty(tiaojian4)) tiaojian6 = tiaojian4;
            }

            switch (this.guanxi6.Text)
            {
                case "并且":
                    guanxi = " and ";
                    break;
                case "或者":
                    guanxi = " or ";
                    break;
            }
           
            if (!String.IsNullOrEmpty(tiaojian5) && !String.IsNullOrEmpty(tiaojian6))
                tiaojian = "(" + tiaojian5 + guanxi + tiaojian6 + ")";
            else
            {
                if (!String.IsNullOrEmpty(tiaojian5)) tiaojian = tiaojian5;
                if (!String.IsNullOrEmpty(tiaojian6)) tiaojian = tiaojian6;
            }
           
        }
    }
}
