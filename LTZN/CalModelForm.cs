using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using LTZN.CalFramework;
using System.Reflection;
using LTZN.Repository;
using C1.Win.C1FlexGrid;

namespace LTZN
{
    public partial class CalModelForm : Form
    {
        public CalModelForm()
        {
            InitializeComponent();
        }

        public CalModelForm(CalModel currentModel)
        {
            InitializeComponent();
            this.CurrentModel = currentModel;
        }

        private CalModel CurrentModel = null;

        private CalModelRepository repository = new CalModelRepository();

        private void CalModelForm_Load(object sender, EventArgs e)
        {
            if (CurrentModel != null)
            {
                btnAddModel.Enabled = true;
                btnAddOrg.Enabled = true;
                btnAddTag.Enabled = true;
                btnAddTag2.Enabled = true;
                btnDeleteModel.Enabled = true;
                btnDeleteOrg.Enabled = true;
                btnDeleteTag.Enabled = true;
                btnCopy.Enabled = true;
                btnPaste.Enabled = true;
                btnSave.Enabled = true;
                buttonCal.Visible = true;
                buttonOk.Visible = true;
                buttonCancel.Visible = true;

                TreeNode node = treeView1.Nodes.Add(CurrentModel.ModelName);
                node.Tag = CurrentModel;
                foreach (var org in CurrentModel.RootOrgs)
                {
                    AddOrgNode(node, org);
                }
            }
            else
            {
                buttonCal.Visible = false;
                buttonOk.Visible = false;
                buttonCancel.Visible = false;
                List<CalModel> models = repository.GetAllModels();
                foreach (CalModel model in models)
                {
                    TreeNode node = treeView1.Nodes.Add(model.ModelName);
                    node.Tag = model;
                    foreach (var org in model.RootOrgs)
                    {
                        AddOrgNode(node, org);
                    }
                }
            }
        }
      

        private void AddOrgNode(TreeNode node, CalOrg org)
        {
            TreeNode orgNode = node.Nodes.Add(org.OrgName);
            orgNode.Tag = org;
            foreach (var subOrg in org.ChildCalOrgs)
            {
                AddOrgNode(orgNode, subOrg);
            }
        }

        private void treeView1_AfterLabelEdit(object sender, NodeLabelEditEventArgs e)
        {
            if (e.CancelEdit) return;

            if (e.Node.Tag.GetType() == typeof(CalModel))
            {
                CalModel model = e.Node.Tag as CalModel;
                if (model != null)
                {
                    model.ModelName = e.Label;
                }
            }

            if (e.Node.Tag.GetType() == typeof(CalOrg))
            {
                CalOrg org = e.Node.Tag as CalOrg;
                if (org != null)
                {
                    org.OrgName = e.Label;
                }
            }
        }

        private void 添加模型ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NameDialog dialog = new NameDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                CalModel model = new CalModel();
                model.EID = Guid.NewGuid().ToString();
                model.ModelName = dialog.Result;
                TreeNode node = treeView1.Nodes.Add(model.ModelName);
                node.Tag = model;
            }
        }

        private void 添加组织结构ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode node = this.treeView1.SelectedNode;
            if (node!= null)
            {
                 NameDialog dialog = new NameDialog();
                 if (dialog.ShowDialog() == DialogResult.OK)
                 {
                     if (node.Tag.GetType() == typeof(CalModel))
                     {
                         CalModel model = node.Tag as CalModel;
                         if (model != null)
                         {
                             CalOrg nOrg = model.AddOrg(dialog.Result);
                             TreeNode nNode = node.Nodes.Add(nOrg.OrgName);
                             nNode.Tag = nOrg;
                         }
                     }
                     else if (node.Tag.GetType() == typeof(CalOrg))
                     {
                         CalOrg org = node.Tag as CalOrg;
                         if (org != null)
                         {
                             CalOrg nOrg = org.AddOrg(dialog.Result);
                             TreeNode nNode = node.Nodes.Add(nOrg.OrgName);
                             nNode.Tag = nOrg;
                         }
                     }
                 }
            }
        }

        private void 添加标签ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode node = this.treeView1.SelectedNode;
            if (node != null)
            {
                NameDialog dialog = new NameDialog();
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    //string tagName = dialog.Result;
                    //string formula = "";
                    //int splitIdx = dialog.Result.IndexOf('=');
                    //if (splitIdx > 0)
                    //{
                    //    tagName = dialog.Result.Substring(0, splitIdx);
                    //    formula = dialog.Result.Substring(splitIdx, dialog.Result.Length - splitIdx - 1);
                    //}
                    
                    if (node.Tag.GetType() == typeof(CalModel))
                    {
                        CalModel model = node.Tag as CalModel;
                        if (model != null)
                        {
                            model.AddTagGroup(dialog.Result);
                        }
                    }
                    else if (node.Tag.GetType() == typeof(CalOrg))
                    {
                        CalOrg org = node.Tag as CalOrg;
                        if (org != null)
                        {
                           org.AddTagGroup(dialog.Result);
                        }
                    }
                    bindingSource1.ResetBindings(false);
                }
            }
        }

        private void 添加标签_Click(object sender, EventArgs e)
        {
            TreeNode node = this.treeView1.SelectedNode;
            if (node != null)
            {
                NameDialog dialog = new NameDialog();
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    if (node.Tag.GetType() == typeof(CalModel))
                    {
                        CalModel model = node.Tag as CalModel;
                        if (model != null)
                        {
                            model.AddTag(dialog.Result);
                        }
                    }
                    else if (node.Tag.GetType() == typeof(CalOrg))
                    {
                        CalOrg org = node.Tag as CalOrg;
                        if (org != null)
                        {
                            org.AddTag(dialog.Result);
                        }
                    }
                    bindingSource1.ResetBindings(false);
                }
            }
        }

        private void 删除模型ToolStripMenuItem_Click(object sender, EventArgs e)
        {
             TreeNode node = this.treeView1.SelectedNode;
             if (node != null)
             {
                 if (node.Tag.GetType() == typeof(CalModel))
                 {
                     if (MessageBox.Show("确实要删除吗?", "提示", MessageBoxButtons.OKCancel) == DialogResult.OK)
                     {
                         CalModel model = node.Tag as CalModel;
                         repository.Delete(model);
                         treeView1.SelectedNode.Remove();
                     }
                 }
             }
        }

        private void 保存ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode node = this.treeView1.SelectedNode;
            if (node != null)
            {
                if (node.Tag.GetType() == typeof(CalModel))
                {
                    CalModel model = node.Tag as CalModel;
                    if (model != null)
                    {
                        repository.Save(model);
                        MessageBox.Show("保存完成!");
                    }
                }
                else
                {
                    while (node.Parent != null && node.Tag.GetType() != typeof(CalModel))
                    {
                        node = node.Parent;
                        if (node.Tag.GetType() == typeof(CalModel))
                        {
                            CalModel model = node.Tag as CalModel;
                            if (model != null)
                            {
                                repository.Save(model);
                                MessageBox.Show("保存完成!");
                            }
                            break;
                        }
                    }
                }
            }
        }

        private void 删除组织ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode node = this.treeView1.SelectedNode;
            if (node != null)
            {
                if (node.Tag.GetType() == typeof(CalOrg))
                {
                    if (node.Parent != null)
                    {
                        if (node.Parent.Tag.GetType() == typeof(CalModel))
                        {
                            if (MessageBox.Show("确实要删除吗?", "提示", MessageBoxButtons.OKCancel) == DialogResult.OK)
                            {
                                CalModel model = node.Parent.Tag as CalModel;
                                model.RemoveOrg(node.Tag as CalOrg);
                                treeView1.SelectedNode.Remove();
                            }
                        }
                        else if (node.Parent.Tag.GetType() == typeof(CalOrg))
                        {
                            if (MessageBox.Show("确实要删除吗?", "提示", MessageBoxButtons.OKCancel) == DialogResult.OK)
                            {
                                CalOrg parentOrg = node.Parent.Tag as CalOrg;
                                parentOrg.RemoveOrg(node.Tag as CalOrg);
                                treeView1.SelectedNode.Remove();
                            }
                        }
                    }
                }
            }
        }

        private List<CalTag> oriData = new List<CalTag>();
        private List<CalTag> filterData = new List<CalTag>();

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            TreeNode node = this.treeView1.SelectedNode;
            if (node != null)
            {
                if (node.Tag.GetType() == typeof(CalModel))
                {
                    CalModel model = node.Tag as CalModel;
                    if (ckbViewType.Checked)
                    {
                        oriData = model.AllCalTags;
                    }
                    else
                    {
                        oriData = model.RootTags;
                    }

                }
                else if (node.Tag.GetType() == typeof(CalOrg))
                {
                    CalOrg org = node.Tag as CalOrg;
                    oriData = org.CalTags;
                }

                Filter();
            }
        }
       

        private void 删除标签ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            for (int rowIdx = c1FlexGrid1.Selection.TopRow; rowIdx <= c1FlexGrid1.Selection.BottomRow; rowIdx++)
            {
                Row row = c1FlexGrid1.Rows[rowIdx];

                CalTag tag = row.DataSource as CalTag;
                if (tag != null)
                {
                    if (tag.MyCalOrg != null)
                    {
                        tag.MyCalOrg.RemoveTag(tag);
                    }
                    else
                    {
                        tag.MyModel.RemoveTag(tag);
                    }
                }
            }
            bindingSource1.ResetBindings(false);
        }

        private void 复制ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            for (int rowIdx = c1FlexGrid1.Selection.TopRow; rowIdx <= c1FlexGrid1.Selection.BottomRow; rowIdx++)
            {
                Row row = c1FlexGrid1.Rows[rowIdx];
                CalTag tag = row.DataSource as CalTag;
                if (tag != null)
                {
                    sb.AppendLine(tag.TagName);
                }
            }

            Clipboard.SetData(DataFormats.UnicodeText, sb.ToString());
        }

        private void 粘贴ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode node = this.treeView1.SelectedNode;
            if (node != null)
            {
                string clipTags = Clipboard.GetData(DataFormats.UnicodeText).ToString();
                System.IO.StringReader strReader = new System.IO.StringReader(clipTags);
                string tagName = null;
                while ((tagName = strReader.ReadLine()) != null)
                {
                    if (node.Tag.GetType() == typeof(CalModel))
                    {
                        CalModel model = node.Tag as CalModel;
                        if (model != null)
                        {
                            CalTag nTag = model.AddTag(tagName);
                        }
                    }
                    else if (node.Tag.GetType() == typeof(CalOrg))
                    {
                        CalOrg org = node.Tag as CalOrg;
                        if (org != null)
                        {
                            CalTag nTag = org.AddTag(tagName);
                        }
                    }
                }
                strReader.Close();
                bindingSource1.ResetBindings(false);
            }
        }

        private void 设定求和公式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (Row row in c1FlexGrid1.Rows)
            {
                if (c1FlexGrid1.Selection.ContainsRow(row.SafeIndex))
                {
                    CalTag tag = row.DataSource as CalTag;
                    List<CalTag> subTags = tag.GetChildTags();
                    if (subTags.Count > 0)
                    {
                        string forma = "";
                        foreach (var item in subTags)
                        {
                            if (forma == "")
                                forma = item.TagFullName;
                            else
                                forma += " + " + item.TagFullName;
                        }
                        tag.Forma = forma;
                    }
                }
            }
        }

        private void 设定分摊公式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
             NameDialog dialog = new NameDialog();
             if (dialog.ShowDialog() == DialogResult.OK)
             {
                 foreach (Row row in c1FlexGrid1.Rows)
                 {
                     if (c1FlexGrid1.Selection.ContainsRow(row.SafeIndex))
                     {
                         
                         CalTag tag = row.DataSource as CalTag;
                         
                         CalTag fenTanTag = tag.GetBrotherTag(dialog.Result);
                         if (fenTanTag != null)
                         {
                             List<CalTag> subTags = tag.GetChildTags();
                             if (subTags.Count > 1)
                             {
                                 string lastFormua = "";
                                 for (int i = 0; i < subTags.Count-1; i++)
                                 {
                                     CalTag item = subTags[i];
                                     lastFormua += "-" + item.TagFullName;
                                     CalTag subFenTag = item.GetBrotherTag(dialog.Result);
                                     if (subFenTag != null)
                                     {
                                         item.Forma = string.Format("{0}*{1}/{2}", tag.TagFullName, subFenTag.TagFullName, fenTanTag.TagFullName);
                                     }
                                 }
                                 subTags[subTags.Count - 1].Forma = tag.TagFullName + lastFormua;
                             }
                             else if (subTags.Count == 1)
                             {
                                 subTags[0].Forma = tag.TagFullName;
                             }
                         }
                     }
                 }
             }
        }

        private void 计算ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CalModel model = getSelectedCalModel();
            if (model != null)
            {
                model.Calc();
            }
        }

        private CalModel getSelectedCalModel()
        {
            CalModel result = null;
            TreeNode node = this.treeView1.SelectedNode;
            if (node != null)
            {
                if (node.Tag.GetType() == typeof(CalModel))
                {
                    result = node.Tag as CalModel;
                }
            }
            return result;
        }

        private CalOrg getSelectedCalOrg()
        {
            CalOrg result = null;
            TreeNode node = this.treeView1.SelectedNode;
            if (node != null)
            {
                if (node.Tag.GetType() == typeof(CalOrg))
                {
                    result = node.Tag as CalOrg;
                }
            }
            return result;
        }

        private void ckbViewType_CheckedChanged(object sender, EventArgs e)
        {
            treeView1_AfterSelect(null, null);
        }

        
        private void txtFilter_TextChanged(object sender, EventArgs e)
        {
            Filter();
        }

        private void Filter()
        {
            if (string.IsNullOrEmpty(txtFilter.Text.Trim()))
            {
                //oriData.Sort(new Comparison<CalTag>(SortCalTag));
                bindingSource1.DataSource = oriData;
                bindingSource1.ResetBindings(false);
            }
            else
            {
                filterData.Clear();
                foreach (var item in oriData)
                {
                    string[] keys = txtFilter.Text.Split(' ');
                    bool find = true;
                    foreach (var key in keys)
                    {
                        switch (cmbFilterTyp.Text)
                        {
                            case "标签名":
                                if (key.Trim() != "" && !item.TagName.Contains(key.Trim()))
                                    find = false;
                                break;
                            case "公式":
                                if (key.Trim() != "" && !item.Forma.Contains(key.Trim()))
                                    find = false;
                                break;
                            default:
                                if (key.Trim() != "" && !item.TagFullName.Contains(key.Trim()))
                                    find = false;
                                break;
                        }

                    }
                    if (find)
                        filterData.Add(item);

                }
                filterData.Sort(new Comparison<CalTag>(SortCalTag));
                bindingSource1.DataSource = filterData;
                bindingSource1.ResetBindings(false);
            }

        }

        private void 设定累计公式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (Row row in c1FlexGrid1.Rows)
            {
                if (c1FlexGrid1.Selection.ContainsRow(row.SafeIndex))
                {
                    CalTag tag = row.DataSource as CalTag;
                    if (tag.TagFullName.StartsWith("累计"))
                    {
                        tag.Forma = tag.TagFullName.Replace("累计", "昨日累计") + "+" + tag.TagFullName.Replace("累计", "本日");
                    }
                }
            }
        }


        private int SortCalTag(CalTag a1, CalTag a2)
        {
            return a1.TagFullName.CompareTo(a2.TagFullName);
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void 上移ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (c1FlexGrid1.RowSel >= 0)
            {
                Row row = c1FlexGrid1.Rows[c1FlexGrid1.RowSel];
                CalTag tag = row.DataSource as CalTag;
                if (tag != null)
                {
                    CalOrg parent = tag.MyCalOrg;
                    if (parent != null)
                    {
                        parent.UpIdx(tag);
                    }
                }
                c1FlexGrid1.Refresh();
            }
        }

        private void 下移ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (c1FlexGrid1.RowSel >= 0)
            {
                Row row = c1FlexGrid1.Rows[c1FlexGrid1.RowSel];
                CalTag tag = row.DataSource as CalTag;
                if (tag != null)
                {
                    CalOrg parent = tag.MyCalOrg;
                    if (parent != null)
                    {
                        parent.DownIdx(tag);
                    }
                }
                c1FlexGrid1.Refresh();
            }
        }
    }
}
