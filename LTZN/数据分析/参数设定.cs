using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;
using System.Security.Principal;

namespace LTZN.数据分析
{
    public partial class 参数设定 : Form
    {
        public 参数设定()
        {
            InitializeComponent();
        }
        private void cs11_TextChanged(object sender, EventArgs e)
        {
            button1.Enabled = true;
            button3.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            save();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            save();
        }

        private void save()
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "update LTZN_Para set Val=:vl where Name=:name";
            OracleCommand saveCmd = new OracleCommand(sql, conn);
            OracleParameter  valPara= saveCmd.Parameters.Add(":vl", OracleType.Double);
            OracleParameter namePara = saveCmd.Parameters.Add(":name", OracleType.VarChar);
            
            namePara.Value = "1高炉S最大";
            valPara.Value = c1NumericEdit36.Value;
            if (Convert.ToDouble(c1NumericEdit36.Value) > 0.07)
            {
                MessageBox.Show("上限不能大于0.07");
                valPara.Value = 0.07;
            }

            saveCmd.ExecuteNonQuery();

            namePara.Value = "1高炉S次大";
            valPara.Value = c1NumericEdit35.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "1高炉S次小";
            valPara.Value = c1NumericEdit34.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "1高炉S最小";
            valPara.Value = c1NumericEdit33.Value;
            saveCmd.ExecuteNonQuery();

        

            namePara.Value = "3高炉S最大";
            valPara.Value = c1NumericEdit44.Value;
            if (Convert.ToDouble(c1NumericEdit44.Value) > 0.07)
            {
                MessageBox.Show("上限不能大于0.07");
                valPara.Value = 0.07;
            }
            saveCmd.ExecuteNonQuery();

            namePara.Value = "3高炉S次大";
            valPara.Value = c1NumericEdit43.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "3高炉S次小";
            valPara.Value = c1NumericEdit42.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "3高炉S最小";
            valPara.Value = c1NumericEdit41.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "5高炉S最大";
            valPara.Value = c1NumericEdit52.Value;
            if (Convert.ToDouble(c1NumericEdit52.Value) > 0.07)
            {
                MessageBox.Show("上限不能大于0.07");
                valPara.Value = 0.07;
            }

            saveCmd.ExecuteNonQuery();

            namePara.Value = "5高炉S次大";
            valPara.Value = c1NumericEdit51.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "5高炉S次小";
            valPara.Value = c1NumericEdit50.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "5高炉S最小";
            valPara.Value = c1NumericEdit49.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "6高炉S最大";
            valPara.Value = c1NumericEdit56.Value;
            if (Convert.ToDouble(c1NumericEdit56.Value) > 0.07)
            {
                MessageBox.Show("上限不能大于0.07");
                valPara.Value = 0.07;
            }
            saveCmd.ExecuteNonQuery();

            namePara.Value = "6高炉S次大";
            valPara.Value = c1NumericEdit55.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "6高炉S次小";
            valPara.Value = c1NumericEdit54.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "6高炉S最小";
            valPara.Value = c1NumericEdit53.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "S显示最大值";
            valPara.Value = c1NumericEdit32.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "S显示最小值";
            valPara.Value = c1NumericEdit31.Value;
            saveCmd.ExecuteNonQuery();



            namePara.Value = "1高炉R最大";
            valPara.Value = c1NumericEdit62.Value;
            if (Convert.ToDouble(c1NumericEdit62.Value) > 1.25)
                MessageBox.Show("上限不能大于1.25");
            valPara.Value = 1.25;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "1高炉R次大";
            valPara.Value = c1NumericEdit61.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "1高炉R次小";
            valPara.Value = c1NumericEdit60.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "1高炉R最小";
            valPara.Value = c1NumericEdit59.Value;
            if (Convert.ToDouble(c1NumericEdit59.Value) < 1.07)
                MessageBox.Show("下限不能小于1.07");
            valPara.Value = 1.07;
            saveCmd.ExecuteNonQuery();



            namePara.Value = "3高炉R最大";
            valPara.Value = c1NumericEdit70.Value;
            if (Convert.ToDouble(c1NumericEdit70.Value) > 1.25)
                MessageBox.Show("上限不能大于1.25");
            valPara.Value = 1.25;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "3高炉R次大";
            valPara.Value = c1NumericEdit69.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "3高炉R次小";
            valPara.Value = c1NumericEdit68.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "3高炉R最小";
            valPara.Value = c1NumericEdit67.Value;
            if (Convert.ToDouble(c1NumericEdit67.Value) < 1.07)
                MessageBox.Show("下限不能小于1.07");
            valPara.Value = 1.07;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "5高炉R最大";
            valPara.Value = c1NumericEdit78.Value;
            if (Convert.ToDouble(c1NumericEdit78.Value) > 1.25)
                MessageBox.Show("上限不能大于1.25");
            valPara.Value = 1.25;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "5高炉R次大";
            valPara.Value = c1NumericEdit77.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "5高炉R次小";
            valPara.Value = c1NumericEdit76.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "5高炉R最小";
            valPara.Value = c1NumericEdit75.Value;
            if (Convert.ToDouble(c1NumericEdit75.Value) < 1.07)
                MessageBox.Show("下限不能小于1.07");
            valPara.Value = 1.07;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "6高炉R最大";
            valPara.Value = c1NumericEdit82.Value;
            if (Convert.ToDouble(c1NumericEdit82.Value) > 1.25)
                MessageBox.Show("上限不能大于1.25");
            valPara.Value = 1.25;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "6高炉R次大";
            valPara.Value = c1NumericEdit81.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "6高炉R次小";
            valPara.Value = c1NumericEdit80.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "6高炉R最小";
            valPara.Value = c1NumericEdit79.Value;
            if (Convert.ToDouble(c1NumericEdit79.Value) < 1.07)
                MessageBox.Show("下限不能小于1.07");
            valPara.Value = 1.07;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "R显示最大";
            valPara.Value = c1NumericEdit58.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "R显示最小";
            valPara.Value = c1NumericEdit57.Value;
        
            saveCmd.ExecuteNonQuery();

            
            namePara.Value = "1高炉Si最大";
            valPara.Value = cs111.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "1高炉Si次大";
            valPara.Value = cs112.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "1高炉Si次小";
            valPara.Value = cs113.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "1高炉Si最小";
            valPara.Value = cs114.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "2高炉Si最大";
            valPara.Value = cs121.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "2高炉Si次大";
            valPara.Value = cs122.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "2高炉Si次小";
            valPara.Value = cs123.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "2高炉Si最小";
            valPara.Value = cs124.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "3高炉Si最大";
            valPara.Value = cs131.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "3高炉Si次大";
            valPara.Value = cs132.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "3高炉Si次小";
            valPara.Value = cs133.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "3高炉Si最小";
            valPara.Value = cs134.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "4高炉Si最大";
            valPara.Value = cs141.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "4高炉Si次大";
            valPara.Value = cs142.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "4高炉Si次小";
            valPara.Value = cs143.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "4高炉Si最小";
            valPara.Value = cs144.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "5高炉Si最大";
            valPara.Value = cs151.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "5高炉Si次大";
            valPara.Value = cs152.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "5高炉Si次小";
            valPara.Value = cs153.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "5高炉Si最小";
            valPara.Value = cs154.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "6高炉Si最大";
            valPara.Value = cs161.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "6高炉Si次大";
            valPara.Value = cs162.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "6高炉Si次小";
            valPara.Value = cs163.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "6高炉Si最小";
            valPara.Value = cs164.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "1高炉物理热最大";
            valPara.Value = c1NumericEdit10.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "1高炉物理热次大";
            valPara.Value = c1NumericEdit9.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "1高炉物理热次小";
            valPara.Value = c1NumericEdit8.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "1高炉物理热最小";
            valPara.Value = c1NumericEdit7.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "3高炉物理热最大";
            valPara.Value = c1NumericEdit18.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "3高炉物理热次大";
            valPara.Value = c1NumericEdit17.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "3高炉物理热次小";
            valPara.Value = c1NumericEdit16.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "3高炉物理热最小";
            valPara.Value = c1NumericEdit15.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "5高炉物理热最大";
            valPara.Value = c1NumericEdit26.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "5高炉物理热次大";
            valPara.Value = c1NumericEdit25.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "5高炉物理热次小";
            valPara.Value = c1NumericEdit24.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "5高炉物理热最小";
            valPara.Value = c1NumericEdit23.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "6高炉物理热最大";
            valPara.Value = c1NumericEdit30.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "6高炉物理热次大";
            valPara.Value = c1NumericEdit29.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "6高炉物理热次小";
            valPara.Value = c1NumericEdit28.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "6高炉物理热最小";
            valPara.Value = c1NumericEdit27.Value;
            saveCmd.ExecuteNonQuery();
               
            namePara.Value = "物理热显示最大值";
            valPara.Value = c1NumericEdit6.Value;
            saveCmd.ExecuteNonQuery();
          
            namePara.Value = "物理热显示最小值";
            valPara.Value = c1NumericEdit5.Value;
            saveCmd.ExecuteNonQuery(); 
  
            namePara.Value = "S单值上限";
            valPara.Value = cs21.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "S单值均值";
            valPara.Value = cs22.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "S单值下限";
            valPara.Value = cs23.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "S极差上限";
            valPara.Value = cs24.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "S极差均值";
            valPara.Value = cs25.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "S极差下限";
            valPara.Value = cs26.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "R最大";
            valPara.Value = cs31.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "R次大";
            valPara.Value = cs32.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "R次小";
            valPara.Value = cs33.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "R最小";
            valPara.Value = cs34.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "Si显示最大值";
            valPara.Value = cs15.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "Si显示最小值";
            valPara.Value = cs16.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "S单值图显示最大值";
            valPara.Value = cs27.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "S单值图显示最小值";
            valPara.Value = cs28.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "S极差图显示最大值";
            valPara.Value = cs29.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "S极差图显示最小值";
            valPara.Value = cs210.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "R显示最大值";
            valPara.Value = cs35.Value;
            saveCmd.ExecuteNonQuery();

            namePara.Value = "R显示最小值";
            valPara.Value = cs36.Value;
            saveCmd.ExecuteNonQuery();

            conn.Close();
            button1.Enabled = false;
            button3.Enabled = false;
        }

        private void loadPara()
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "select name,val from LTZN_Para ";
            OracleCommand cmd = new OracleCommand(sql, conn);
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                string name = dr.GetString(0);
                switch (name)
                {
                    case "1高炉R最大":
                        c1NumericEdit62.Value = dr.GetValue(1);
                        break;
                    case "1高炉R次大":
                        c1NumericEdit61.Value = dr.GetValue(1);
                        break;
                    case "1高炉R次小":
                        c1NumericEdit60.Value = dr.GetValue(1);
                        break;
                    case "1高炉R最小":
                        c1NumericEdit59.Value = dr.GetValue(1);
                        break;
                    case "3高炉R最大":
                        c1NumericEdit70.Value = dr.GetValue(1);
                        break;
                    case "3高炉R次大":
                        c1NumericEdit69.Value = dr.GetValue(1);
                        break;
                    case "3高炉R次小":
                        c1NumericEdit68.Value = dr.GetValue(1);
                        break;
                    case "3高炉R最小":
                        c1NumericEdit67.Value = dr.GetValue(1);
                        break;
                    case "5高炉R最大":
                        c1NumericEdit78.Value = dr.GetValue(1);
                        break;
                    case "5高炉R次大":
                        c1NumericEdit77.Value = dr.GetValue(1);
                        break;
                    case "5高炉R次小":
                        c1NumericEdit76.Value = dr.GetValue(1);
                        break;
                    case "5高炉R最小":
                        c1NumericEdit75.Value = dr.GetValue(1);
                        break;
                    case "6高炉R最大":
                        c1NumericEdit82.Value = dr.GetValue(1);
                        break;
                    case "6高炉R次大":
                        c1NumericEdit81.Value = dr.GetValue(1);
                        break;
                    case "6高炉R次小":
                        c1NumericEdit80.Value = dr.GetValue(1);
                        break;
                    case "6高炉R最小":
                        c1NumericEdit79.Value = dr.GetValue(1);
                        break;
                    case "R显示最大":
                        c1NumericEdit58.Value = dr.GetValue(1);
                        break;
                    case "R显示最小":
                        c1NumericEdit57.Value = dr.GetValue(1);
                        break;
                    case "1高炉S最大":
                        c1NumericEdit36.Value = dr.GetValue(1);
                        break;
                    case "1高炉S次大":
                        c1NumericEdit35.Value = dr.GetValue(1);
                        break;
                    case "1高炉S次小":
                        c1NumericEdit34.Value = dr.GetValue(1);
                        break;
                    case "1高炉S最小":
                        c1NumericEdit33.Value = dr.GetValue(1);
                        break;
                    case "3高炉S最大":
                        c1NumericEdit44.Value = dr.GetValue(1);
                        break;
                    case "3高炉S次大":
                        c1NumericEdit43.Value = dr.GetValue(1);
                        break;
                    case "3高炉S次小":
                        c1NumericEdit42.Value = dr.GetValue(1);
                        break;
                    case "3高炉S最小":
                        c1NumericEdit41.Value = dr.GetValue(1);
                        break;
                    case "5高炉S最大":
                        c1NumericEdit52.Value = dr.GetValue(1);
                        break;
                    case "5高炉S次大":
                        c1NumericEdit51.Value = dr.GetValue(1);
                        break;
                    case "5高炉S次小":
                        c1NumericEdit50.Value = dr.GetValue(1);
                        break;
                    case "5高炉S最小":
                        c1NumericEdit49.Value = dr.GetValue(1);
                        break;
                    case "6高炉S最大":
                        c1NumericEdit56.Value = dr.GetValue(1);
                        break;
                    case "6高炉S次大":
                        c1NumericEdit55.Value = dr.GetValue(1);
                        break;
                    case "6高炉S次小":
                        c1NumericEdit54.Value = dr.GetValue(1);
                        break;
                    case "6高炉S最小":
                        c1NumericEdit53.Value = dr.GetValue(1);
                        break;
                    case "S显示最大值":
                        c1NumericEdit32.Value = dr.GetValue(1);
                        break;
                    case "S显示最小值":
                        c1NumericEdit31.Value = dr.GetValue(1);
                        break;
                    case "1高炉Si最大":
                        cs111.Value = dr.GetValue(1);
                        break;
                    case "1高炉Si次大":
                        cs112.Value = dr.GetValue(1);
                        break;
                    case "1高炉Si次小":
                        cs113.Value = dr.GetValue(1);
                        break;
                    case "1高炉Si最小":
                        cs114.Value = dr.GetValue(1);
                        break;
                    case "2高炉Si最大":
                        cs121.Value = dr.GetValue(1);
                        break;
                    case "2高炉Si次大":
                        cs122.Value = dr.GetValue(1);
                        break;
                    case "2高炉Si次小":
                        cs123.Value = dr.GetValue(1);
                        break;
                    case "2高炉Si最小":
                        cs124.Value = dr.GetValue(1);
                        break;
                    case "3高炉Si最大":
                        cs131.Value = dr.GetValue(1);
                        break;
                    case "3高炉Si次大":
                        cs132.Value = dr.GetValue(1);
                        break;
                    case "3高炉Si次小":
                        cs133.Value = dr.GetValue(1);
                        break;
                    case "3高炉Si最小":
                        cs134.Value = dr.GetValue(1);
                        break;
                    case "4高炉Si最大":
                        cs141.Value = dr.GetValue(1);
                        break;
                    case "4高炉Si次大":
                        cs142.Value = dr.GetValue(1);
                        break;
                    case "4高炉Si次小":
                        cs143.Value = dr.GetValue(1);
                        break;
                    case "4高炉Si最小":
                        cs144.Value = dr.GetValue(1);
                        break;
                    case "5高炉Si最大":
                        cs151.Value = dr.GetValue(1);
                        break;
                    case "5高炉Si次大":
                        cs152.Value = dr.GetValue(1);
                        break;
                    case "5高炉Si次小":
                        cs153.Value = dr.GetValue(1);
                        break;
                    case "5高炉Si最小":
                        cs154.Value = dr.GetValue(1);
                        break;
                    case "6高炉Si最大":
                        cs161.Value = dr.GetValue(1);
                        break;
                    case "6高炉Si次大":
                        cs162.Value = dr.GetValue(1);
                        break;
                    case "6高炉Si次小":
                        cs163.Value = dr.GetValue(1);
                        break;
                    case "6高炉Si最小":
                        cs164.Value = dr.GetValue(1);
                        break;
                    case "S单值上限":
                        cs21.Value = dr.GetValue(1);
                        break;
                    case "S单值均值":
                        cs22.Value = dr.GetValue(1);
                        break;
                    case "S单值下限":
                        cs23.Value = dr.GetValue(1);
                        break;
                    case "S极差上限":
                        cs24.Value = dr.GetValue(1);
                        break;
                    case "S极差均值":
                        cs25.Value = dr.GetValue(1);
                        break;
                    case "S极差下限":
                        cs26.Value = dr.GetValue(1);
                        break;
                    case "R最大":
                        cs31.Value = dr.GetValue(1);
                        break;
                    case "R次大":
                        cs32.Value = dr.GetValue(1);
                        break;
                    case "R次小":
                        cs33.Value = dr.GetValue(1);
                        break;
                    case "R最小":
                        cs34.Value = dr.GetValue(1);
                        break;
                    case "Si显示最大值":
                        cs15.Value = dr.GetValue(1);
                        break;
                    case "Si显示最小值":
                        cs16.Value = dr.GetValue(1);
                        break;

                    case "S单值图显示最大值":
                        cs27.Value = dr.GetValue(1);
                        break;
                    case "S单值图显示最小值":
                        cs28.Value = dr.GetValue(1);
                        break;
                    case "S极差图显示最大值":
                        cs29.Value = dr.GetValue(1);
                        break;
                    case "S极差图显示最小值":
                        cs210.Value = dr.GetValue(1);
                        break;
                    case "R显示最大值":
                        cs35.Value = dr.GetValue(1);
                        break;
                    case "R显示最小值":
                        cs36.Value = dr.GetValue(1);
                        break;
                    case "1高炉物理热最大":
                        c1NumericEdit10.Value = dr.GetValue(1);
                        break;

                    case "1高炉物理热次大":
                        c1NumericEdit9.Value = dr.GetValue(1);
                        break;

                    case "1高炉物理热次小":
                        c1NumericEdit8.Value = dr.GetValue(1);
                        break;
                    case "1高炉物理热最小":
                        c1NumericEdit7.Value = dr.GetValue(1);
                        break;

                    case "3高炉物理热最大":
                        c1NumericEdit18.Value = dr.GetValue(1);
                        break;

                    case "3高炉物理热次大":
                        c1NumericEdit17.Value = dr.GetValue(1);
                        break;


                    case "3高炉物理热次小":
                        c1NumericEdit16.Value = dr.GetValue(1);
                        break;

                    case "3高炉物理热最小":
                        c1NumericEdit15.Value = dr.GetValue(1);
                        break;
                    case "5高炉物理热最大":
                        c1NumericEdit26.Value = dr.GetValue(1);
                        break;
                    case "5高炉物理热次大":
                        c1NumericEdit25.Value = dr.GetValue(1);
                        break;

                    case "5高炉物理热次小":
                        c1NumericEdit24.Value = dr.GetValue(1);
                        break;

                    case "5高炉物理热最小":
                        c1NumericEdit23.Value = dr.GetValue(1);
                        break;

                    case "6高炉物理热最大":
                        c1NumericEdit30.Value = dr.GetValue(1);
                        break;


                    case "6高炉物理热次大":
                        c1NumericEdit29.Value = dr.GetValue(1);
                        break;

                    case "6高炉物理热次小":
                        c1NumericEdit28.Value = dr.GetValue(1);
                        break;
                    case "6高炉物理热最小":
                        c1NumericEdit27.Value = dr.GetValue(1);
                        break;
                    case "物理热显示最大值":
                        c1NumericEdit6.Value = dr.GetValue(1);
                        break;
                    case "物理热显示最小值":
                        c1NumericEdit5.Value = dr.GetValue(1);
                        break;  
                }
            }
            dr.Close();
            conn.Close(); 
            button1.Enabled = false;
            button3.Enabled = false;
        }

        private void 参数设定_Load(object sender, EventArgs e)
        {
  
            IPrincipal p = LtznUserManager.instance.CurrentUser;
      
            //this.cs11.TabPages.Remove(cs11.TabPages[1]);
          //  this.cs11.TabPages.Remove(cs11.TabPages[2]);
            loadPara();

            if (!p.IsInRole("管理员"))
            {
                if (p.IsInRole("1高炉"))
                {
                    tabPage13.Parent = null;
                    tabPage15.Parent = null;
                    tabPage16.Parent = null;
                    tabPage20.Parent = null;
                    tabPage22.Parent = null;
                    tabPage23.Parent = null;
                    tabPage26.Parent = null;
                    tabPage28.Parent = null;
                    tabPage29.Parent = null;
                }
                if (p.IsInRole("3高炉"))
                {
                    tabPage11.Parent = null;
                    tabPage15.Parent = null;
                    tabPage16.Parent = null;
                    tabPage18.Parent = null;
                    tabPage22.Parent = null;
                    tabPage23.Parent = null;
                    tabPage24.Parent = null;
                    tabPage28.Parent = null;
                    tabPage29.Parent = null;
                }
                if (p.IsInRole("5高炉"))
                {

                    tabPage11.Parent = null;
                    tabPage13.Parent = null;
                    tabPage16.Parent = null;
                    tabPage18.Parent = null;
                    tabPage20.Parent = null;
                    tabPage23.Parent = null;
                    tabPage24.Parent = null;
                    tabPage26.Parent = null;
                    tabPage29.Parent = null;


                }
                if (p.IsInRole("6高炉"))
                {
                    tabPage11.Parent = null;
                    tabPage13.Parent = null;
                    tabPage15.Parent = null;
                    tabPage18.Parent = null;
                    tabPage20.Parent = null;
                    tabPage22.Parent = null;
                    tabPage24.Parent = null;
                    tabPage26.Parent = null;
                    tabPage28.Parent = null;
                }
            }
            loadPara();
        }

        private void cs11_SelectedIndexChanged(object sender, EventArgs e)
        {
            button1.Enabled = true;
            button3.Enabled = true;

        }

      
    }
}