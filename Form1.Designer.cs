using System.Windows.Forms;

namespace ExcelInstanceLoader
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("All");
            this.buttonRefresh = new System.Windows.Forms.Button();
            this.buttonKillGhost = new System.Windows.Forms.Button();
            this.buttonKillSelected = new System.Windows.Forms.Button();
            this.buttonKillAll = new System.Windows.Forms.Button();
            this.labelStatus = new System.Windows.Forms.Label();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.checkBoxSelectAll = new System.Windows.Forms.CheckBox();
            this.listView1 = new System.Windows.Forms.ListView();
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.SuspendLayout();
            // 
            // buttonRefresh
            // 
            this.buttonRefresh.Location = new System.Drawing.Point(53, 37);
            this.buttonRefresh.Name = "buttonRefresh";
            this.buttonRefresh.Size = new System.Drawing.Size(111, 59);
            this.buttonRefresh.TabIndex = 0;
            this.buttonRefresh.Text = "Refresh";
            this.buttonRefresh.UseVisualStyleBackColor = true;
            this.buttonRefresh.Click += new System.EventHandler(this.buttonRefresh_Click);
            // 
            // buttonKillGhost
            // 
            this.buttonKillGhost.Location = new System.Drawing.Point(189, 37);
            this.buttonKillGhost.Name = "buttonKillGhost";
            this.buttonKillGhost.Size = new System.Drawing.Size(111, 59);
            this.buttonKillGhost.TabIndex = 1;
            this.buttonKillGhost.Text = "Kill Ghost";
            this.buttonKillGhost.UseVisualStyleBackColor = true;
            this.buttonKillGhost.Click += new System.EventHandler(this.buttonKillGhost_Click);
            // 
            // buttonKillSelected
            // 
            this.buttonKillSelected.Location = new System.Drawing.Point(326, 37);
            this.buttonKillSelected.Name = "buttonKillSelected";
            this.buttonKillSelected.Size = new System.Drawing.Size(111, 59);
            this.buttonKillSelected.TabIndex = 3;
            this.buttonKillSelected.Text = "Kill Selected";
            this.buttonKillSelected.UseVisualStyleBackColor = true;
            this.buttonKillSelected.Click += new System.EventHandler(this.buttonKillSelected_Click);
            // 
            // buttonKillAll
            // 
            this.buttonKillAll.Location = new System.Drawing.Point(463, 37);
            this.buttonKillAll.Name = "buttonKillAll";
            this.buttonKillAll.Size = new System.Drawing.Size(111, 59);
            this.buttonKillAll.TabIndex = 9;
            this.buttonKillAll.Text = "Kill All";
            this.buttonKillAll.UseVisualStyleBackColor = true;
            this.buttonKillAll.Click += new System.EventHandler(this.buttonKillAll_Click);
            // 
            // labelStatus
            // 
            this.labelStatus.AutoSize = true;
            this.labelStatus.Location = new System.Drawing.Point(50, 114);
            this.labelStatus.Name = "labelStatus";
            this.labelStatus.Size = new System.Drawing.Size(55, 13);
            this.labelStatus.TabIndex = 2;
            this.labelStatus.Text = "Welcome!";
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.CheckOnClick = true;
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(590, 144);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(171, 424);
            this.checkedListBox1.TabIndex = 7;
            this.checkedListBox1.Visible = false;
            // 
            // checkBoxSelectAll
            // 
            this.checkBoxSelectAll.AutoSize = true;
            this.checkBoxSelectAll.Location = new System.Drawing.Point(590, 114);
            this.checkBoxSelectAll.Name = "checkBoxSelectAll";
            this.checkBoxSelectAll.Size = new System.Drawing.Size(70, 17);
            this.checkBoxSelectAll.TabIndex = 8;
            this.checkBoxSelectAll.Text = "Select All";
            this.checkBoxSelectAll.UseVisualStyleBackColor = true;
            this.checkBoxSelectAll.CheckedChanged += new System.EventHandler(this.checkBoxSelectAll_CheckedChanged);
            this.checkBoxSelectAll.Visible = false;
            // 
            // listView1
            // 
            this.listView1.CheckBoxes = true;
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(769, 145);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(176, 423);
            this.listView1.TabIndex = 10;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.List;
            this.listView1.SelectedIndexChanged += new System.EventHandler(this.ListView_SelectedIndexChanged);
            this.listView1.Visible = false;
            // 
            // treeView1
            // 
            this.treeView1.CheckBoxes = true;
            this.treeView1.FullRowSelect = true;
            this.treeView1.Location = new System.Drawing.Point(53, 144);
            this.treeView1.Name = "treeView1";
            this.treeView1.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode1});
            this.treeView1.Size = new System.Drawing.Size(521, 423);
            this.treeView1.TabIndex = 11;
            this.treeView1.AfterCheck += this.TreeView_NodeAfterCheck;
            this.treeView1.AfterSelect += this.TreeView_NodeAfterSelect;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(957, 651);
            this.Controls.Add(this.treeView1);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.buttonKillAll);
            this.Controls.Add(this.checkBoxSelectAll);
            this.Controls.Add(this.checkedListBox1);
            this.Controls.Add(this.buttonKillSelected);
            this.Controls.Add(this.labelStatus);
            this.Controls.Add(this.buttonKillGhost);
            this.Controls.Add(this.buttonRefresh);
            this.Name = "Form1";
            this.Text = "ExcelViewer";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonRefresh;
        private System.Windows.Forms.Button buttonKillGhost;
        private System.Windows.Forms.Label labelStatus;
        private System.Windows.Forms.Button buttonKillSelected;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.CheckBox checkBoxSelectAll;
        private System.Windows.Forms.Button buttonKillAll;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.TreeView treeView1;
    }
}

