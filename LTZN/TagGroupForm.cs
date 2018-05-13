using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using LTZN.CalFramework;

namespace LTZN
{
    public partial class TagGroupForm : Form
    {
        public TagGroupForm()
        {
            InitializeComponent();
        }

        private Calf_TaggroupTable taggroupTable = new Calf_TaggroupTable();

        private void TagGroupForm_Load(object sender, EventArgs e)
        {
            taggroupTable.Load();
            foreach (var item in taggroupTable)
            {
                lstTagGroups.Items.Add(item);
            }

        }

        private Calf_Taggroup _selectTagGroup = null;
        public Calf_Taggroup SelectTagGroup
        {
            get {
                return _selectTagGroup;
            }
            set {
                if (_selectTagGroup != value)
                {
                    _selectTagGroup = value;
                    lstTags.Items.Clear();
                    foreach (var item in _selectTagGroup.Calf_Taggroup_ItemTable)
                    {
                        lstTags.Items.Add(item);
                    }
                }
            }
        }

        private void lstTagGroups_SelectedIndexChanged(object sender, EventArgs e)
        {
            SelectTagGroup = lstTagGroups.SelectedItem as Calf_Taggroup;
        }

        private void 添加标签组ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NameDialog dialog=new NameDialog();
            if(dialog.ShowDialog()==DialogResult.OK)
            {
                Calf_Taggroup group=new Calf_Taggroup();
                group.EID=Guid.NewGuid().ToString();
                group.ENAME = dialog.Result;
                taggroupTable.Add(group);
                lstTagGroups.Items.Add(group);
            }
        }

        private void 添加标签ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (SelectTagGroup != null)
            {
                NameDialog dialog = new NameDialog();
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    Calf_Taggroup_Item groupItem = new Calf_Taggroup_Item();
                    groupItem.EID = Guid.NewGuid().ToString();
                    groupItem.TAGGROUPID = SelectTagGroup.EID;
                    groupItem.TAGNAME = dialog.Result;
                    SelectTagGroup.Calf_Taggroup_ItemTable.Add(groupItem);
                    lstTags.Items.Add(groupItem);
                }

            }
        }

        private void TagGroupForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            taggroupTable.Save();
        }

        private void 删除标签ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (SelectTagGroup != null && lstTags.SelectedIndex >= 0)
            {
                SelectTagGroup.Calf_Taggroup_ItemTable.Remove(lstTags.SelectedItem as Calf_Taggroup_Item);
                lstTags.Items.RemoveAt(lstTags.SelectedIndex);
            }
        }

        private void 删除标签组ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (lstTagGroups.SelectedIndex >= 0)
            {
    
                taggroupTable.Remove(lstTagGroups.SelectedItem as Calf_Taggroup);
                lstTagGroups.Items.RemoveAt(lstTagGroups.SelectedIndex);
            }
        }

    }
}
