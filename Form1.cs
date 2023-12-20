using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.ListBox;
using TreeView = System.Windows.Forms.TreeView;
using Microsoft.Office.Interop.Excel;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ExcelInstanceLoader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            m_ghostProcesses = new List<int>();
            m_excelProcesses = new Dictionary<int, bool>();
            m_processId2WbNames = new Dictionary<int, List<string>>();
            m_checkedWbs = new Dictionary<int, List<Tuple<int, string>>>();

            Reload();
        }

        private List<int> m_ghostProcesses;
        private Dictionary<int, bool> m_excelProcesses;
        private Dictionary<int, List<string>> m_processId2WbNames;
        private Dictionary<int, List<Tuple<int,string>>> m_checkedWbs;

        private void Reload()
        {
            labelStatus.Text = "Refreshing...";

            m_excelProcesses = Utils.GetExcelProcessIds();
            m_ghostProcesses = m_excelProcesses.Where(item => item.Value).ToDictionary(dict => dict.Key, dict => dict.Value).Keys.ToList();

            m_processId2WbNames.Clear();
            foreach (var processId in m_excelProcesses)
            {
                if (!processId.Value)
                    m_processId2WbNames[processId.Key] = Utils.GetWorkbookNamesByProcessId(processId.Key);
            }

            Utils.GenerateTreeView(treeView1, m_processId2WbNames);

            UpdateStatus();
        }

        private void UpdateStatus()
        {
            var nProcesses = m_excelProcesses.Count;
            var nGhost = m_ghostProcesses.Count;

            m_checkedWbs = Utils.GetCheckedNodes(treeView1);
            int nSelected = 0;
            foreach (var item in m_checkedWbs) { nSelected += item.Value.Count; }
            int nWbs = 0;
            foreach (var item in m_processId2WbNames)
            {
                nWbs += item.Value.Count;
            }

            string status = "Found " + nWbs + " workbook";
            status += nWbs == 1 ? "  " : "s ";
            status += "(" + nSelected + " selected), " + nProcesses + " Excel process";
            status += nProcesses == 1 ? "   " : "es ";
            status += "(" + nGhost + " invisible).";
            labelStatus.Text = status;
        }

        private void buttonRefresh_Click(object sender, EventArgs e)
        {
            Reload();
        }

        private void KillGhost()
        {
            foreach (int processId in m_ghostProcesses)
            {
                try
                {
                    Process process = Process.GetProcessById(processId);
                    process.Kill();
                }
                catch (Exception)
                {
                    continue;
                }

                m_excelProcesses.Remove(processId);
            }

            m_ghostProcesses.Clear();
        }

        private void KillSelected()
        {
            Dictionary<int, List<string>> processId2WbNames = new Dictionary<int, List<string>>();

            foreach (var item in m_checkedWbs)
            {
                List<string> value = item.Value.Select(tuple => tuple.Item2).ToList();
                processId2WbNames[item.Key] = value;
            }

            if (Utils.CloseWorkbookByIdAndName(processId2WbNames))
            {
                foreach (var item in processId2WbNames)
                {
                    m_processId2WbNames[item.Key].RemoveAll(listItem => item.Value.Contains(listItem));
                    if (m_processId2WbNames[item.Key].Count == 0)
                    {
                        m_processId2WbNames.Remove(item.Key);
                        m_excelProcesses.Remove(item.Key);
                    }
                }

                m_checkedWbs.Clear();

                Utils.GenerateTreeView(treeView1, m_processId2WbNames);
            }
        }

        private void buttonKillGhost_Click(object sender, EventArgs e)
        {
            KillGhost();
            UpdateStatus();
        }

        private void buttonKillSelected_Click(object sender, EventArgs e)
        {
            KillSelected();
            UpdateStatus();
        }

        private void buttonKillAll_Click(object sender, EventArgs e)
        {
            if (Utils.CloseWorkbookByIdAndName(m_processId2WbNames))
            {
                m_processId2WbNames.Clear();
                m_excelProcesses.Clear();
                m_checkedWbs.Clear();
                treeView1.Nodes[0].Nodes.Clear();

                Utils.GenerateTreeView(treeView1, m_processId2WbNames);
            }

            buttonKillGhost_Click(sender, e);
        }

        private void ListView_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            // Synchronize the checked state with the selection
            if (sender is System.Windows.Forms.ListView listView)
            {
                if (e.Item.Checked)
                {
                    if (listView.SelectedIndices.Contains(e.Item.Index) == false)
                    {
                        listView.SelectedIndices.Add(e.Item.Index);
                    }
                }
                else
                {
                    listView.SelectedIndices.Remove(e.Item.Index);
                }
            }
        }

        private void ListView_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Synchronize the selection with the checked state
            if (sender is System.Windows.Forms.ListView listView)
            {
                foreach (ListViewItem item in listView.SelectedItems)
                {
                    item.Checked = true;
                }

                foreach (ListViewItem item in listView.Items)
                {
                    if (item.Checked && listView.SelectedItems.Contains(item) == false)
                    {
                        item.Checked = false;
                    }
                }
            }
        }

        private void TreeView_NodeAfterCheck(object sender, TreeViewEventArgs e)
        {
            // The code only executes if the user caused the checked state to change.
            if (e.Action != TreeViewAction.Unknown)
            {
                if (sender is TreeView tv)
                {
                    tv.BeginUpdate();
                    Utils.UpdateCheckStatus(e.Node, e.Node.Checked);
                    tv.EndUpdate();
                }

            }

            UpdateStatus();
        }

        private void TreeView_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (sender is TreeView treeView)
            {
                var localPosition = treeView.PointToClient(Cursor.Position);
                var hitTestInfo = treeView.HitTest(localPosition);
                if (hitTestInfo.Location == TreeViewHitTestLocations.StateImage)
                    return;
            }
        }

        private void TreeView_NodeAfterSelect(object sender, TreeViewEventArgs e)
        {
            // The code only executes if the user caused the checked state to change.
            if (e.Action != TreeViewAction.Unknown)
            {
                e.Node.ExpandAll();
            }
        }
    }
}