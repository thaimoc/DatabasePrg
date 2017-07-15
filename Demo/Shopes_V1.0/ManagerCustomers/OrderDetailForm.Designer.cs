namespace ManagerCustomers
{
    partial class OrderDetailForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OrderDetailForm));
            this.grbadd = new System.Windows.Forms.GroupBox();
            this.txtOrderID = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.cbbEmployee = new System.Windows.Forms.ComboBox();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.cbbCustomer = new System.Windows.Forms.ComboBox();
            this.dtpRequiredDate = new System.Windows.Forms.DateTimePicker();
            this.btnNewEmployee = new System.Windows.Forms.Button();
            this.dtpShippedDate = new System.Windows.Forms.DateTimePicker();
            this.btnNewCustomer = new System.Windows.Forms.Button();
            this.dtpOrderDate = new System.Windows.Forms.DateTimePicker();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.btnCancell = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.txtStatus = new System.Windows.Forms.TextBox();
            this.txtShipPostalCode = new System.Windows.Forms.TextBox();
            this.txtShipAddress = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtFreight = new System.Windows.Forms.TextBox();
            this.Status = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.gvOrderDetails = new System.Windows.Forms.DataGridView();
            this.UnitPrice = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Quantity = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Discount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.grbadd.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvOrderDetails)).BeginInit();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // grbadd
            // 
            this.grbadd.BackColor = System.Drawing.Color.White;
            this.grbadd.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("grbadd.BackgroundImage")));
            this.grbadd.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.grbadd.Controls.Add(this.txtOrderID);
            this.grbadd.Controls.Add(this.label8);
            this.grbadd.Controls.Add(this.cbbEmployee);
            this.grbadd.Controls.Add(this.btnUpdate);
            this.grbadd.Controls.Add(this.cbbCustomer);
            this.grbadd.Controls.Add(this.dtpRequiredDate);
            this.grbadd.Controls.Add(this.btnNewEmployee);
            this.grbadd.Controls.Add(this.dtpShippedDate);
            this.grbadd.Controls.Add(this.btnNewCustomer);
            this.grbadd.Controls.Add(this.dtpOrderDate);
            this.grbadd.Controls.Add(this.label11);
            this.grbadd.Controls.Add(this.label12);
            this.grbadd.Controls.Add(this.btnCancell);
            this.grbadd.Controls.Add(this.btnAdd);
            this.grbadd.Controls.Add(this.txtStatus);
            this.grbadd.Controls.Add(this.txtShipPostalCode);
            this.grbadd.Controls.Add(this.txtShipAddress);
            this.grbadd.Controls.Add(this.label5);
            this.grbadd.Controls.Add(this.txtFreight);
            this.grbadd.Controls.Add(this.Status);
            this.grbadd.Controls.Add(this.label7);
            this.grbadd.Controls.Add(this.label6);
            this.grbadd.Controls.Add(this.label15);
            this.grbadd.Controls.Add(this.label13);
            this.grbadd.Controls.Add(this.label14);
            this.grbadd.Location = new System.Drawing.Point(3, 5);
            this.grbadd.Name = "grbadd";
            this.grbadd.Size = new System.Drawing.Size(359, 374);
            this.grbadd.TabIndex = 43;
            this.grbadd.TabStop = false;
            // 
            // txtOrderID
            // 
            this.txtOrderID.Location = new System.Drawing.Point(80, 16);
            this.txtOrderID.Name = "txtOrderID";
            this.txtOrderID.ReadOnly = true;
            this.txtOrderID.Size = new System.Drawing.Size(163, 20);
            this.txtOrderID.TabIndex = 62;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.BackColor = System.Drawing.Color.Transparent;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label8.Location = new System.Drawing.Point(9, 16);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(53, 15);
            this.label8.TabIndex = 61;
            this.label8.Text = "Order ID";
            // 
            // cbbEmployee
            // 
            this.cbbEmployee.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbbEmployee.FormattingEnabled = true;
            this.cbbEmployee.Location = new System.Drawing.Point(80, 76);
            this.cbbEmployee.Name = "cbbEmployee";
            this.cbbEmployee.Size = new System.Drawing.Size(163, 21);
            this.cbbEmployee.TabIndex = 48;
            // 
            // btnUpdate
            // 
            this.btnUpdate.Location = new System.Drawing.Point(83, 344);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(81, 23);
            this.btnUpdate.TabIndex = 60;
            this.btnUpdate.Text = "Update";
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // cbbCustomer
            // 
            this.cbbCustomer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbbCustomer.FormattingEnabled = true;
            this.cbbCustomer.Location = new System.Drawing.Point(80, 49);
            this.cbbCustomer.Name = "cbbCustomer";
            this.cbbCustomer.Size = new System.Drawing.Size(163, 21);
            this.cbbCustomer.TabIndex = 49;
            // 
            // dtpRequiredDate
            // 
            this.dtpRequiredDate.CustomFormat = "dd/MM/yyyy";
            this.dtpRequiredDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpRequiredDate.Location = new System.Drawing.Point(262, 111);
            this.dtpRequiredDate.Name = "dtpRequiredDate";
            this.dtpRequiredDate.Size = new System.Drawing.Size(91, 20);
            this.dtpRequiredDate.TabIndex = 57;
            this.dtpRequiredDate.ValueChanged += new System.EventHandler(this.dtpRequiredDate_ValueChanged);
            // 
            // btnNewEmployee
            // 
            this.btnNewEmployee.Location = new System.Drawing.Point(261, 74);
            this.btnNewEmployee.Name = "btnNewEmployee";
            this.btnNewEmployee.Size = new System.Drawing.Size(91, 23);
            this.btnNewEmployee.TabIndex = 47;
            this.btnNewEmployee.Text = "New Employee";
            this.btnNewEmployee.UseVisualStyleBackColor = true;
            this.btnNewEmployee.Click += new System.EventHandler(this.btnNewEmployee_Click);
            // 
            // dtpShippedDate
            // 
            this.dtpShippedDate.CustomFormat = "dd/MM/yyyy";
            this.dtpShippedDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpShippedDate.Location = new System.Drawing.Point(83, 149);
            this.dtpShippedDate.Name = "dtpShippedDate";
            this.dtpShippedDate.Size = new System.Drawing.Size(81, 20);
            this.dtpShippedDate.TabIndex = 58;
            this.dtpShippedDate.ValueChanged += new System.EventHandler(this.dtpShippedDate_ValueChanged);
            // 
            // btnNewCustomer
            // 
            this.btnNewCustomer.Location = new System.Drawing.Point(261, 47);
            this.btnNewCustomer.Name = "btnNewCustomer";
            this.btnNewCustomer.Size = new System.Drawing.Size(91, 23);
            this.btnNewCustomer.TabIndex = 46;
            this.btnNewCustomer.Text = "New Customer";
            this.btnNewCustomer.UseVisualStyleBackColor = true;
            this.btnNewCustomer.Click += new System.EventHandler(this.btnNewCustomer_Click);
            // 
            // dtpOrderDate
            // 
            this.dtpOrderDate.CustomFormat = "dd/MM/yyyy";
            this.dtpOrderDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpOrderDate.Location = new System.Drawing.Point(83, 111);
            this.dtpOrderDate.Name = "dtpOrderDate";
            this.dtpOrderDate.Size = new System.Drawing.Size(81, 20);
            this.dtpOrderDate.TabIndex = 59;
            this.dtpOrderDate.ValueChanged += new System.EventHandler(this.dtpOrderDate_ValueChanged);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.BackColor = System.Drawing.Color.Transparent;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label11.Location = new System.Drawing.Point(5, 47);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(60, 15);
            this.label11.TabIndex = 45;
            this.label11.Text = "Customer";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.BackColor = System.Drawing.Color.Transparent;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label12.Location = new System.Drawing.Point(5, 79);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(62, 15);
            this.label12.TabIndex = 44;
            this.label12.Text = "Employee";
            // 
            // btnCancell
            // 
            this.btnCancell.Location = new System.Drawing.Point(278, 344);
            this.btnCancell.Name = "btnCancell";
            this.btnCancell.Size = new System.Drawing.Size(75, 23);
            this.btnCancell.TabIndex = 55;
            this.btnCancell.Text = "Cancell";
            this.btnCancell.UseVisualStyleBackColor = true;
            this.btnCancell.Click += new System.EventHandler(this.btnCancell_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(170, 344);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(102, 23);
            this.btnAdd.TabIndex = 56;
            this.btnAdd.Text = "Add New Order";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // txtStatus
            // 
            this.txtStatus.Location = new System.Drawing.Point(219, 142);
            this.txtStatus.Name = "txtStatus";
            this.txtStatus.Size = new System.Drawing.Size(134, 20);
            this.txtStatus.TabIndex = 52;
            // 
            // txtShipPostalCode
            // 
            this.txtShipPostalCode.Location = new System.Drawing.Point(180, 183);
            this.txtShipPostalCode.Name = "txtShipPostalCode";
            this.txtShipPostalCode.Size = new System.Drawing.Size(175, 20);
            this.txtShipPostalCode.TabIndex = 51;
            // 
            // txtShipAddress
            // 
            this.txtShipAddress.Location = new System.Drawing.Point(83, 209);
            this.txtShipAddress.Multiline = true;
            this.txtShipAddress.Name = "txtShipAddress";
            this.txtShipAddress.Size = new System.Drawing.Size(272, 129);
            this.txtShipAddress.TabIndex = 54;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.BackColor = System.Drawing.Color.Transparent;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label5.Location = new System.Drawing.Point(9, 183);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(45, 15);
            this.label5.TabIndex = 44;
            this.label5.Text = "Freight";
            // 
            // txtFreight
            // 
            this.txtFreight.Location = new System.Drawing.Point(83, 183);
            this.txtFreight.Name = "txtFreight";
            this.txtFreight.Size = new System.Drawing.Size(81, 20);
            this.txtFreight.TabIndex = 53;
            this.txtFreight.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFreight_KeyPress);
            // 
            // Status
            // 
            this.Status.AutoSize = true;
            this.Status.BackColor = System.Drawing.Color.Transparent;
            this.Status.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Status.ForeColor = System.Drawing.Color.MidnightBlue;
            this.Status.Location = new System.Drawing.Point(182, 145);
            this.Status.Name = "Status";
            this.Status.Size = new System.Drawing.Size(41, 15);
            this.Status.TabIndex = 46;
            this.Status.Text = "Status";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.BackColor = System.Drawing.Color.Transparent;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label7.Location = new System.Drawing.Point(182, 165);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(101, 15);
            this.label7.TabIndex = 45;
            this.label7.Text = "Ship Postal Code";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.BackColor = System.Drawing.Color.Transparent;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label6.Location = new System.Drawing.Point(5, 207);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(79, 15);
            this.label6.TabIndex = 50;
            this.label6.Text = "Ship Address";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.BackColor = System.Drawing.Color.Transparent;
            this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label15.Location = new System.Drawing.Point(170, 111);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(87, 15);
            this.label15.TabIndex = 49;
            this.label15.Text = "Required Date";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.BackColor = System.Drawing.Color.Transparent;
            this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label13.Location = new System.Drawing.Point(5, 111);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(67, 15);
            this.label13.TabIndex = 48;
            this.label13.Text = "Order Date";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.BackColor = System.Drawing.Color.Transparent;
            this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label14.Location = new System.Drawing.Point(5, 149);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(82, 15);
            this.label14.TabIndex = 47;
            this.label14.Text = "Shipped Date";
            // 
            // gvOrderDetails
            // 
            this.gvOrderDetails.BackgroundColor = System.Drawing.Color.White;
            this.gvOrderDetails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.gvOrderDetails.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvOrderDetails.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.UnitPrice,
            this.Quantity,
            this.Discount});
            this.gvOrderDetails.Location = new System.Drawing.Point(368, 12);
            this.gvOrderDetails.Name = "gvOrderDetails";
            this.gvOrderDetails.Size = new System.Drawing.Size(570, 367);
            this.gvOrderDetails.TabIndex = 44;
            this.gvOrderDetails.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.gvOrderDetails_DataError);
            // 
            // UnitPrice
            // 
            this.UnitPrice.DataPropertyName = "UnitPrice";
            this.UnitPrice.HeaderText = "Unit Price";
            this.UnitPrice.Name = "UnitPrice";
            this.UnitPrice.Width = 120;
            // 
            // Quantity
            // 
            this.Quantity.DataPropertyName = "Quantity";
            this.Quantity.HeaderText = "Quantity";
            this.Quantity.Name = "Quantity";
            // 
            // Discount
            // 
            this.Discount.DataPropertyName = "Discount";
            this.Discount.HeaderText = "Discount";
            this.Discount.Name = "Discount";
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 385);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(942, 25);
            this.statusStrip1.TabIndex = 45;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.toolStripStatusLabel1.ForeColor = System.Drawing.SystemColors.Highlight;
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(99, 20);
            this.toolStripStatusLabel1.Text = "Total money: ";
            // 
            // OrderDetailForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(942, 410);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.gvOrderDetails);
            this.Controls.Add(this.grbadd);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "OrderDetailForm";
            this.Text = "OrderDetailForm";
            this.Load += new System.EventHandler(this.OrderDetailForm_Load);
            this.grbadd.ResumeLayout(false);
            this.grbadd.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvOrderDetails)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox grbadd;
        private System.Windows.Forms.TextBox txtOrderID;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox cbbEmployee;
        private System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.ComboBox cbbCustomer;
        private System.Windows.Forms.DateTimePicker dtpRequiredDate;
        private System.Windows.Forms.Button btnNewEmployee;
        private System.Windows.Forms.DateTimePicker dtpShippedDate;
        private System.Windows.Forms.Button btnNewCustomer;
        private System.Windows.Forms.DateTimePicker dtpOrderDate;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Button btnCancell;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.TextBox txtStatus;
        private System.Windows.Forms.TextBox txtShipPostalCode;
        private System.Windows.Forms.TextBox txtShipAddress;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtFreight;
        private System.Windows.Forms.Label Status;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.DataGridView gvOrderDetails;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.DataGridViewTextBoxColumn UnitPrice;
        private System.Windows.Forms.DataGridViewTextBoxColumn Quantity;
        private System.Windows.Forms.DataGridViewTextBoxColumn Discount;
    }
}