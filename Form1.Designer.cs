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
            // treeView1
            // 
            this.treeView1.CheckBoxes = true;
            this.treeView1.FullRowSelect = true;
            this.treeView1.Location = new System.Drawing.Point(53, 144);
            this.treeView1.Name = "treeView1";
            treeNode1.Name = "";
            treeNode1.Text = "All";
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
            this.ClientSize = new System.Drawing.Size(643, 651);
            this.Controls.Add(this.treeView1);
            this.Controls.Add(this.buttonKillAll);
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
        private System.Windows.Forms.Button buttonKillAll;
        private System.Windows.Forms.TreeView treeView1;
    }
}

