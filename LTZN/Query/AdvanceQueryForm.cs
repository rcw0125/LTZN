using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using C1.Win.C1FlexGrid;
using System.Xml;
using ZHCDB;
namespace LTZN.Query
{
    public partial class AdvanceQueryForm : Form
    {
       private  QuerytemplateTable queryTaemplate = new QuerytemplateTable();
       private QueryEntity queryEntity = null;
        //private Querytemplate curTemplate = null;

       public AdvanceQueryForm(QueryEntity queryEntity)
        {
            InitializeComponent();
            this.queryEntity = queryEntity;
        }

        private void AdvanceQueryForm_Load(object sender, EventArgs e)
        {
            string comboList = " ";
            foreach (var item in queryEntity)
            {
                comboList += "|" + item.FieldDispaly;

            }

            this.c1FlexGrid1.Cols[2].ComboList = comboList;

            // this.c1FlexGrid1.ec

            CellStyle dateCs = c1FlexGrid1.Styles.Add("DateTime");
            dateCs.DataType = typeof(DateTime);
            dateCs.Format = "yyyy-MM-dd HH:mm:ss";
            //  dateCs.ForeColor = Color.DarkGoldenrod;

            CellStyle strCs = c1FlexGrid1.Styles.Add("string");
            strCs.DataType = typeof(string);
            // strCs.ForeColor = Color.DarkGoldenrod;

            CellStyle doubleCs = c1FlexGrid1.Styles.Add("double");
            doubleCs.DataType = typeof(double);

            CellStyle intCs = c1FlexGrid1.Styles.Add("int");
            intCs.DataType = typeof(int);
            // doubleCs.ForeColor = Color.DarkGoldenrod;


            CellStyle boolCs = c1FlexGrid1.Styles.Add("bool");
            boolCs.DataType = typeof(bool);

            CellStyle dateCs_op = c1FlexGrid1.Styles.Add("DateTime_RelationOp");
            dateCs_op.DataType = typeof(string);
            dateCs_op.ComboList = "等于|不等于|大于|大于等于|小于|小于等于|是空|非空";

            CellStyle strCs_op = c1FlexGrid1.Styles.Add("string_RelationOp");
            strCs_op.DataType = typeof(string);
            strCs_op.ComboList = "等于|不等于|大于|大于等于|小于|小于等于|包含|以其开始|以其结束|是空|非空";

            CellStyle doubleCs_op = c1FlexGrid1.Styles.Add("double_RelationOp");
            doubleCs_op.DataType = typeof(string);
            doubleCs_op.ComboList = "等于|不等于|大于|大于等于|小于|小于等于|是空|非空";

            CellStyle intoubleCs_op = c1FlexGrid1.Styles.Add("int_RelationOp");
            intoubleCs_op.DataType = typeof(string);
            intoubleCs_op.ComboList = "等于|不等于|大于|大于等于|小于|小于等于|是空|非空";

            CellStyle boolCs_op = c1FlexGrid1.Styles.Add("bool_RelationOp");
            boolCs_op.DataType = typeof(string);
            boolCs_op.ComboList = "等于|不等于|是空|非空";


            queryTaemplate.LoadByEntityName(queryEntity.QueryEntityName);
            listBox1.DataSource = queryTaemplate;
            listBox1.DisplayMember = "TEMPLATENAME";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<OpItem> opList = new List<OpItem>();
            OpItem ksOpt = new OpItem(OpState.开始);
            opList.Add(ksOpt);

            for (int row = 1; row < this.c1FlexGrid1.Rows.Count; row++)
            {
                string lgStr = getCellString(row, 0);
                if (row > 1 && lgStr != "")
                {
                    OpItem lgOpt = new OpItem(OpState.逻辑关系, (object)Enum.Parse(typeof(LogicOpType), lgStr));
                    opList.Add(lgOpt);
                }

                string zkhStr = getCellString(row, 1);
                if (zkhStr != "")
                {
                    for (int i = 0; i < zkhStr.Length; i++)
                    {
                        OpItem lgOpt = new OpItem(OpState.左括号);
                        opList.Add(lgOpt);
                    }
                }

                string fieldStr = getCellString(row, 2);
                string relationStr = getCellString(row, 3);
                object fieldVal = getCellVal(row, 4);
                if (fieldStr != "" && relationStr != "")
                {
                    EntityField field = queryEntity.GetByDisplay(fieldStr);
                    if (field != null)
                    {
                        SealedCondition condition = field.Create(relationStr, fieldVal);
                        if (condition != null)
                        {
                            OpItem cdOpt = new OpItem(OpState.条件, condition);
                            opList.Add(cdOpt);
                        }
                    }

                }

                string ykhStr = getCellString(row, 5);
                if (ykhStr != "")
                {
                    for (int i = 0; i < ykhStr.Length; i++)
                    {
                        OpItem lgOpt = new OpItem(OpState.右括号);
                        opList.Add(lgOpt);
                    }
                }
            }

            OpItem esOpt = new OpItem(OpState.结束);
            opList.Add(esOpt);

            if (ConditionManager.CheckOpList(opList))
            {
                Condition rCd = ConditionManager.ConvertOpList(opList);
                if (rCd != null)
                {
                    this.WhereSql = rCd.ToSql(); ;
                    DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                      MessageBox.Show("查询条件不正确");
                }
            }
            else
            {
                MessageBox.Show("查询条件不正确");
            }

        }

        private string getCellString(int row, int col)
        {
            string result = "";
            if (this.c1FlexGrid1[row, col] != null && this.c1FlexGrid1[row, col] != DBNull.Value)
            {
                result= this.c1FlexGrid1[row, col].ToString().Trim();
            }
            return result;
        }

        private object getCellVal(int row, int col)
        {
            object result = null;
            if (this.c1FlexGrid1[row, col] != null && this.c1FlexGrid1[row, col] != DBNull.Value)
            {
                result = this.c1FlexGrid1[row, col];
            }
            return result;
        }

        private void c1FlexGrid1_AfterEdit(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            if (e.Col == 2)
            {
                if (this.c1FlexGrid1[e.Row, e.Col] != null && this.c1FlexGrid1[e.Row, e.Col] != DBNull.Value && this.c1FlexGrid1[e.Row, e.Col].ToString().Trim() != "")
                {
                    string fieldName = this.c1FlexGrid1[e.Row, e.Col].ToString();
                    foreach (var item in queryEntity)
                    {
                        if (item.FieldDispaly == fieldName)
                        {
                            CellStyle style = this.c1FlexGrid1.Styles[item.FieldType];
                            this.c1FlexGrid1.SetCellStyle(e.Row, 4, style);
                            CellStyle style2 = this.c1FlexGrid1.Styles[item.FieldType + "_RelationOp"];
                            this.c1FlexGrid1.SetCellStyle(e.Row, 3, style2);
                            break;
                        }
                    }
                }
                else
                {
                    this.c1FlexGrid1[e.Row, 3] = null;
                    this.c1FlexGrid1[e.Row, 4] = null;
                }
            }
        }

        private void saveTemplate(object sender, EventArgs e)
        {
            NameDialog dialog = new NameDialog();
            if (dialog.ShowDialog() == DialogResult.OK && dialog.Result != "")
            {
                Querytemplate tem = new Querytemplate();
                tem.ENTITYNAME = queryEntity.QueryEntityName;
                tem.TEMPLATENAME = dialog.Result;
                tem.XMLTEXT = GetQueryXml();
                queryTaemplate.Add(tem);
                queryTaemplate.Save();
            }
        }
     
        private string GetQueryXml()
        {
            XmlDocument doc = new XmlDocument();
            XmlElement rows = doc.CreateElement("Rows");
            doc.AppendChild(rows);

            for (int i = 1; i < this.c1FlexGrid1.Rows.Count; i++)
            {
                bool hasContent = false;
                XmlElement row = doc.CreateElement("Row");
                for (int col = 0; col < 6; col++)
                {
                    string col1Str = getCellString(i, col);
                    XmlElement col1 = doc.CreateElement("Col" + col.ToString());
                    col1.InnerText = col1Str;
                    if (col1Str != "")
                    {
                        hasContent = true;
                    }
                    row.AppendChild(col1);
                }
                if (hasContent)
                    rows.AppendChild(row);
            }
            return doc.InnerXml;
        }

        private void applyTemplate(object sender, EventArgs e)
        {
            Querytemplate curTemplate = this.listBox1.SelectedItem as Querytemplate;
            if (curTemplate != null)
            {
                C1FlexGrid flex = this.c1FlexGrid1;
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(curTemplate.XMLTEXT);

                foreach (Row r in flex.Rows)
                {
                    if (r.Index > 0)
                        r.Clear(ClearFlags.Content);
                }

                int row = 1;
                foreach (XmlNode item in doc.ChildNodes[0])
                {
                    int col = 0;
                    foreach (XmlNode subitem in item.ChildNodes)
                    {
                        flex[row, col] = subitem.InnerText;
                        col++;
                    }
                    row++;
                }

                for (int iRow = 1; iRow < this.c1FlexGrid1.Rows.Count; iRow++)
                {

                    string fieldName = getCellString(iRow, 2);
                    if (fieldName != "")
                    {
                        foreach (var item in queryEntity)
                        {
                            if (item.FieldDispaly == fieldName)
                            {
                                CellStyle style = this.c1FlexGrid1.Styles[item.FieldType];
                                this.c1FlexGrid1.SetCellStyle(iRow, 4, style);
                                CellStyle style2 = this.c1FlexGrid1.Styles[item.FieldType + "_RelationOp"];
                                this.c1FlexGrid1.SetCellStyle(iRow, 3, style2);
                                break;
                            }
                        }
                    }
                    else
                    {
                        this.c1FlexGrid1[iRow, 3] = null;
                        this.c1FlexGrid1[iRow, 4] = null;
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
             Querytemplate curTemplate = this.listBox1.SelectedItem as Querytemplate;
             if (curTemplate != null)
             {
                 curTemplate.XMLTEXT = GetQueryXml();
                 curTemplate.Save();
                 MessageBox.Show("保存完成");
             }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Querytemplate curTemplate = this.listBox1.SelectedItem as Querytemplate;
            if (curTemplate != null)
            {
                queryTaemplate.Remove(curTemplate);
                queryTaemplate.Save();
            }
        }

        public string WhereSql
        {
            get;
            set;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
