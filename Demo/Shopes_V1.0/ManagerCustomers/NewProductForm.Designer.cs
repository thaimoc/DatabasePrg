namespace ManagerCustomers
{
    partial class NewProductForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(NewProductForm));
            this.btnUpdate = new System.Windows.Forms.Button();
            this.btnCancell = new System.Windows.Forms.Button();
            this.btnNewSupplier = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.ckbDiscontinued = new System.Windows.Forms.CheckBox();
            this.cbbSupplier = new System.Windows.Forms.ComboBox();
            this.txtQuantityPerUnit = new System.Windows.Forms.TextBox();
            this.txtProductName = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txtUnitPrice = new System.Windows.Forms.TextBox();
            this.txtUnitInStoct = new System.Windows.Forms.TextBox();
            this.txtUnitOnOrder = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtProductID = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnUpdate
            // 
            this.btnUpdate.Location = new System.Drawing.Point(565, 204);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(75, 23);
            this.btnUpdate.TabIndex = 40;
            this.btnUpdate.Text = "Update";
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // btnCancell
            // 
            this.btnCancell.Location = new System.Drawing.Point(778, 204);
            this.btnCancell.Name = "btnCancell";
            this.btnCancell.Size = new System.Drawing.Size(75, 23);
            this.btnCancell.TabIndex = 39;
            this.btnCancell.Text = "Cancell";
            this.btnCancell.UseVisualStyleBackColor = true;
            this.btnCancell.Click += new System.EventHandler(this.btnCancell_Click);
            // 
            // btnNewSupplier
            // 
            this.btnNewSupplier.Location = new System.Drawing.Point(436, 123);
            this.btnNewSupplier.Name = "btnNewSupplier";
            this.btnNewSupplier.Size = new System.Drawing.Size(100, 23);
            this.btnNewSupplier.TabIndex = 38;
            this.btnNewSupplier.Text = "New Supplier";
            this.btnNewSupplier.UseVisualStyleBackColor = true;
            this.btnNewSupplier.Click += new System.EventHandler(this.btnNewSupplier_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Enabled = false;
            this.btnAdd.Location = new System.Drawing.Point(671, 204);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 36;
            this.btnAdd.Text = "Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // ckbDiscontinued
            // 
            this.ckbDiscontinued.AutoSize = true;
            this.ckbDiscontinued.Location = new System.Drawing.Point(569, 181);
            this.ckbDiscontinued.Name = "ckbDiscontinued";
            this.ckbDiscontinued.Size = new System.Drawing.Size(88, 17);
            this.ckbDiscontinued.TabIndex = 35;
            this.ckbDiscontinued.Text = "Discontinued";
            this.ckbDiscontinued.UseVisualStyleBackColor = true;
            // 
            // cbbSupplier
            // 
            this.cbbSupplier.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbbSupplier.FormattingEnabled = true;
            this.cbbSupplier.Location = new System.Drawing.Point(207, 126);
            this.cbbSupplier.Name = "cbbSupplier";
            this.cbbSupplier.Size = new System.Drawing.Size(223, 21);
            this.cbbSupplier.TabIndex = 31;
            this.cbbSupplier.SelectedIndexChanged += new System.EventHandler(this.cbbSupplier_SelectedIndexChanged);
            // 
            // txtQuantityPerUnit
            // 
            this.txtQuantityPerUnit.Location = new System.Drawing.Point(207, 156);
            this.txtQuantityPerUnit.Name = "txtQuantityPerUnit";
            this.txtQuantityPerUnit.Size = new System.Drawing.Size(329, 20);
            this.txtQuantityPerUnit.TabIndex = 29;
            // 
            // txtProductName
            // 
            this.txtProductName.Location = new System.Drawing.Point(207, 97);
            this.txtProductName.Name = "txtProductName";
            this.txtProductName.Size = new System.Drawing.Size(329, 20);
            this.txtProductName.TabIndex = 30;
            this.txtProductName.TextChanged += new System.EventHandler(this.txtProductName_TextChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Location = new System.Drawing.Point(114, 156);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(87, 13);
            this.label4.TabIndex = 24;
            this.label4.Text = "Quantity Per Unit";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(566, 131);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(74, 13);
            this.label7.TabIndex = 23;
            this.label7.Text = "Units In Stock";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Location = new System.Drawing.Point(114, 100);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(75, 13);
            this.label2.TabIndex = 22;
            this.label2.Text = "Product Name";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(566, 160);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(77, 13);
            this.label6.TabIndex = 25;
            this.label6.Text = "Units On Order";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Location = new System.Drawing.Point(114, 129);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(45, 13);
            this.label3.TabIndex = 28;
            this.label3.Text = "Supplier";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(566, 104);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(53, 13);
            this.label5.TabIndex = 27;
            this.label5.Text = "Unit Price";
            // 
            // txtUnitPrice
            // 
            this.txtUnitPrice.Location = new System.Drawing.Point(666, 101);
            this.txtUnitPrice.Name = "txtUnitPrice";
            this.txtUnitPrice.Size = new System.Drawing.Size(182, 20);
            this.txtUnitPrice.TabIndex = 41;
            this.txtUnitPrice.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtUnitPrice_KeyPress);
            // 
            // txtUnitInStoct
            // 
            this.txtUnitInStoct.Location = new System.Drawing.Point(666, 124);
            this.txtUnitInStoct.Name = "txtUnitInStoct";
            this.txtUnitInStoct.Size = new System.Drawing.Size(182, 20);
            this.txtUnitInStoct.TabIndex = 41;
            this.txtUnitInStoct.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtUnitInStoct_KeyPress);
            // 
            // txtUnitOnOrder
            // 
            this.txtUnitOnOrder.Location = new System.Drawing.Point(666, 150);
            this.txtUnitOnOrder.Name = "txtUnitOnOrder";
            this.txtUnitOnOrder.Size = new System.Drawing.Size(182, 20);
            this.txtUnitOnOrder.TabIndex = 41;
            this.txtUnitOnOrder.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtUnitInStoct_KeyPress);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Location = new System.Drawing.Point(114, 72);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 13);
            this.label1.TabIndex = 42;
            this.label1.Text = "Product ID";
            // 
            // txtProductID
            // 
            this.txtProductID.Location = new System.Drawing.Point(207, 72);
            this.txtProductID.Name = "txtProductID";
            this.txtProductID.ReadOnly = true;
            this.txtProductID.Size = new System.Drawing.Size(160, 20);
            this.txtProductID.TabIndex = 43;
            // 
            // NewProductForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ClientSize = new System.Drawing.Size(866, 257);
            this.Controls.Add(this.txtProductID);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtUnitOnOrder);
            this.Controls.Add(this.txtUnitInStoct);
            this.Controls.Add(this.txtUnitPrice);
            this.Controls.Add(this.btnUpdate);
            this.Controls.Add(this.btnCancell);
            this.Controls.Add(this.btnNewSupplier);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.ckbDiscontinued);
            this.Controls.Add(this.cbbSupplier);
            this.Controls.Add(this.txtQuantityPerUnit);
            this.Controls.Add(this.txtProductName);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label5);
            this.DoubleBuffered = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "NewProductForm";
            this.Text = "NewProductForm";
            this.Load += new System.EventHandler(this.NewProductForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.Button btnCancell;
        private System.Windows.Forms.Button btnNewSupplier;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.CheckBox ckbDiscontinued;
        private System.Windows.Forms.ComboBox cbbSupplier;
        private System.Windows.Forms.TextBox txtQuantityPerUnit;
        private System.Windows.Forms.TextBox txtProductName;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtUnitPrice;
        private System.Windows.Forms.TextBox txtUnitInStoct;
        private System.Windows.Forms.TextBox txtUnitOnOrder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtProductID;
    }
}