namespace ManagerCustomers
{
    partial class SortForm
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
            this.btnRemove = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.lblExpressionCurrent = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.lvExpressions = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.cklbColumns = new System.Windows.Forms.CheckedListBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.grbSortType = new System.Windows.Forms.GroupBox();
            this.radDESC = new System.Windows.Forms.RadioButton();
            this.radASC = new System.Windows.Forms.RadioButton();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.grbSortType.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnRemove
            // 
            this.btnRemove.Enabled = false;
            this.btnRemove.Location = new System.Drawing.Point(539, 329);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(69, 23);
            this.btnRemove.TabIndex = 19;
            this.btnRemove.Text = "Remove";
            this.btnRemove.UseVisualStyleBackColor = true;
            this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.lblExpressionCurrent);
            this.groupBox3.Location = new System.Drawing.Point(166, 245);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(608, 61);
            this.groupBox3.TabIndex = 18;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "View";
            // 
            // lblExpressionCurrent
            // 
            this.lblExpressionCurrent.AutoSize = true;
            this.lblExpressionCurrent.Location = new System.Drawing.Point(6, 16);
            this.lblExpressionCurrent.Name = "lblExpressionCurrent";
            this.lblExpressionCurrent.Size = new System.Drawing.Size(0, 13);
            this.lblExpressionCurrent.TabIndex = 11;
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.lvExpressions);
            this.groupBox2.Location = new System.Drawing.Point(163, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(611, 230);
            this.groupBox2.TabIndex = 17;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "View Expression";
            // 
            // lvExpressions
            // 
            this.lvExpressions.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1});
            this.lvExpressions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lvExpressions.GridLines = true;
            this.lvExpressions.Location = new System.Drawing.Point(3, 16);
            this.lvExpressions.Name = "lvExpressions";
            this.lvExpressions.Size = new System.Drawing.Size(605, 211);
            this.lvExpressions.TabIndex = 0;
            this.lvExpressions.UseCompatibleStateImageBehavior = false;
            this.lvExpressions.View = System.Windows.Forms.View.Details;
            this.lvExpressions.SelectedIndexChanged += new System.EventHandler(this.lvExpressions_SelectedIndexChanged);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Expression";
            this.columnHeader1.Width = 600;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.cklbColumns);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(142, 230);
            this.groupBox1.TabIndex = 16;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Columns";
            // 
            // cklbColumns
            // 
            this.cklbColumns.CheckOnClick = true;
            this.cklbColumns.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cklbColumns.FormattingEnabled = true;
            this.cklbColumns.Location = new System.Drawing.Point(3, 16);
            this.cklbColumns.Name = "cklbColumns";
            this.cklbColumns.Size = new System.Drawing.Size(136, 211);
            this.cklbColumns.TabIndex = 0;
            this.cklbColumns.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.cklbColumns_ItemCheck);
            this.cklbColumns.SelectedIndexChanged += new System.EventHandler(this.cklbColumns_SelectedIndexChanged);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(709, 329);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(67, 23);
            this.btnClose.TabIndex = 15;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Enabled = false;
            this.btnAdd.Location = new System.Drawing.Point(623, 329);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(66, 23);
            this.btnAdd.TabIndex = 14;
            this.btnAdd.Text = "Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // grbSortType
            // 
            this.grbSortType.Controls.Add(this.radDESC);
            this.grbSortType.Controls.Add(this.radASC);
            this.grbSortType.Location = new System.Drawing.Point(15, 245);
            this.grbSortType.Name = "grbSortType";
            this.grbSortType.Size = new System.Drawing.Size(136, 54);
            this.grbSortType.TabIndex = 13;
            this.grbSortType.TabStop = false;
            this.grbSortType.Text = "Sort Type";
            // 
            // radDESC
            // 
            this.radDESC.AutoSize = true;
            this.radDESC.Location = new System.Drawing.Point(76, 19);
            this.radDESC.Name = "radDESC";
            this.radDESC.Size = new System.Drawing.Size(54, 17);
            this.radDESC.TabIndex = 0;
            this.radDESC.Text = "DESC";
            this.radDESC.UseVisualStyleBackColor = true;
            // 
            // radASC
            // 
            this.radASC.AutoSize = true;
            this.radASC.Checked = true;
            this.radASC.Location = new System.Drawing.Point(6, 19);
            this.radASC.Name = "radASC";
            this.radASC.Size = new System.Drawing.Size(46, 17);
            this.radASC.TabIndex = 0;
            this.radASC.TabStop = true;
            this.radASC.Text = "ASC";
            this.radASC.UseVisualStyleBackColor = true;
            // 
            // SortForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(780, 360);
            this.Controls.Add(this.btnRemove);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.grbSortType);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SortForm";
            this.Text = "SortForm2";
            this.Load += new System.EventHandler(this.SortForm_Load);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.grbSortType.ResumeLayout(false);
            this.grbSortType.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnRemove;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label lblExpressionCurrent;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ListView lvExpressions;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckedListBox cklbColumns;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.GroupBox grbSortType;
        private System.Windows.Forms.RadioButton radDESC;
        private System.Windows.Forms.RadioButton radASC;
    }
}