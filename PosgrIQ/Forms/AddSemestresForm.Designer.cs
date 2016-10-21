namespace PosgrIQ
{
    partial class AddSemestresForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddSemestresForm));
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.numCod = new System.Windows.Forms.NumericUpDown();
            this.label1 = new System.Windows.Forms.Label();
            this.numAno = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbPeriodo = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.dateTomarQualify = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.datePropuestaDoct = new System.Windows.Forms.DateTimePicker();
            this.label6 = new System.Windows.Forms.Label();
            this.datePropuestaMaes = new System.Windows.Forms.DateTimePicker();
            this.label7 = new System.Windows.Forms.Label();
            this.datePedirQualify = new System.Windows.Forms.DateTimePicker();
            this.label8 = new System.Windows.Forms.Label();
            this.dateTema = new System.Windows.Forms.DateTimePicker();
            ((System.ComponentModel.ISupportInitialize)(this.numCod)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numAno)).BeginInit();
            this.SuspendLayout();
            // 
            // btnAdd
            // 
            this.btnAdd.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnAdd.FlatAppearance.BorderColor = System.Drawing.Color.DarkRed;
            this.btnAdd.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAdd.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAdd.Location = new System.Drawing.Point(227, 265);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 10;
            this.btnAdd.Text = "Agregar";
            this.btnAdd.UseVisualStyleBackColor = false;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.FlatAppearance.BorderColor = System.Drawing.Color.DarkRed;
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancel.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnCancel.Location = new System.Drawing.Point(146, 265);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 9;
            this.btnCancel.Text = "Cancelar";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // numCod
            // 
            this.numCod.Font = new System.Drawing.Font("Calibri", 9F);
            this.numCod.Location = new System.Drawing.Point(60, 9);
            this.numCod.Maximum = new decimal(new int[] {
            999,
            0,
            0,
            0});
            this.numCod.Name = "numCod";
            this.numCod.Size = new System.Drawing.Size(45, 22);
            this.numCod.TabIndex = 1;
            this.numCod.Value = new decimal(new int[] {
            999,
            0,
            0,
            0});
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 9F);
            this.label1.Location = new System.Drawing.Point(12, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 14);
            this.label1.TabIndex = 15;
            this.label1.Text = "Codigo:";
            // 
            // numAno
            // 
            this.numAno.Font = new System.Drawing.Font("Calibri", 9F);
            this.numAno.Location = new System.Drawing.Point(147, 9);
            this.numAno.Maximum = new decimal(new int[] {
            3000,
            0,
            0,
            0});
            this.numAno.Minimum = new decimal(new int[] {
            2000,
            0,
            0,
            0});
            this.numAno.Name = "numAno";
            this.numAno.Size = new System.Drawing.Size(45, 22);
            this.numAno.TabIndex = 2;
            this.numAno.Value = new decimal(new int[] {
            2999,
            0,
            0,
            0});
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Calibri", 9F);
            this.label2.Location = new System.Drawing.Point(115, 13);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(31, 14);
            this.label2.TabIndex = 15;
            this.label2.Text = "Año:";
            // 
            // cmbPeriodo
            // 
            this.cmbPeriodo.FormattingEnabled = true;
            this.cmbPeriodo.Items.AddRange(new object[] {
            "1",
            "2"});
            this.cmbPeriodo.Location = new System.Drawing.Point(257, 9);
            this.cmbPeriodo.Name = "cmbPeriodo";
            this.cmbPeriodo.Size = new System.Drawing.Size(45, 22);
            this.cmbPeriodo.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Calibri", 9F);
            this.label3.Location = new System.Drawing.Point(204, 13);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(52, 14);
            this.label3.TabIndex = 15;
            this.label3.Text = "Periodo:";
            // 
            // dateTomarQualify
            // 
            this.dateTomarQualify.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTomarQualify.Location = new System.Drawing.Point(206, 53);
            this.dateTomarQualify.Name = "dateTomarQualify";
            this.dateTomarQualify.Size = new System.Drawing.Size(96, 22);
            this.dateTomarQualify.TabIndex = 4;
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(12, 48);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(164, 33);
            this.label4.TabIndex = 20;
            this.label4.Text = "Fecha límite para presentar el Qualify (3er Doctorado):";
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(12, 89);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(164, 33);
            this.label5.TabIndex = 22;
            this.label5.Text = "Fecha límite para presentar la propuesta(4to Doctorado):";
            // 
            // datePropuestaDoct
            // 
            this.datePropuestaDoct.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.datePropuestaDoct.Location = new System.Drawing.Point(206, 94);
            this.datePropuestaDoct.Name = "datePropuestaDoct";
            this.datePropuestaDoct.Size = new System.Drawing.Size(96, 22);
            this.datePropuestaDoct.TabIndex = 5;
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(12, 130);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(164, 33);
            this.label6.TabIndex = 24;
            this.label6.Text = "Fecha límite para presentar la propuesta (2do Maestria):";
            // 
            // datePropuestaMaes
            // 
            this.datePropuestaMaes.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.datePropuestaMaes.Location = new System.Drawing.Point(206, 135);
            this.datePropuestaMaes.Name = "datePropuestaMaes";
            this.datePropuestaMaes.Size = new System.Drawing.Size(96, 22);
            this.datePropuestaMaes.TabIndex = 6;
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(12, 171);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(164, 33);
            this.label7.TabIndex = 26;
            this.label7.Text = "Fecha límite para solicitar el Qualify (2do Doctorado):";
            // 
            // datePedirQualify
            // 
            this.datePedirQualify.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.datePedirQualify.Location = new System.Drawing.Point(206, 176);
            this.datePedirQualify.Name = "datePedirQualify";
            this.datePedirQualify.Size = new System.Drawing.Size(96, 22);
            this.datePedirQualify.TabIndex = 7;
            // 
            // label8
            // 
            this.label8.Location = new System.Drawing.Point(12, 212);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(164, 48);
            this.label8.TabIndex = 28;
            this.label8.Text = "Fecha límite para presentar el Tema (1er Maestria y Doctorado):";
            // 
            // dateTema
            // 
            this.dateTema.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTema.Location = new System.Drawing.Point(206, 225);
            this.dateTema.Name = "dateTema";
            this.dateTema.Size = new System.Drawing.Size(96, 22);
            this.dateTema.TabIndex = 8;
            // 
            // AddSemestresForm
            // 
            this.AcceptButton = this.btnAdd;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(310, 297);
            this.ControlBox = false;
            this.Controls.Add(this.label8);
            this.Controls.Add(this.dateTema);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.datePedirQualify);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.datePropuestaMaes);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.datePropuestaDoct);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.dateTomarQualify);
            this.Controls.Add(this.cmbPeriodo);
            this.Controls.Add(this.numAno);
            this.Controls.Add(this.numCod);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnCancel);
            this.Font = new System.Drawing.Font("Calibri", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "AddSemestresForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "AddSemestresForm";
            this.Load += new System.EventHandler(this.AddSemestresForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numCod)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numAno)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.NumericUpDown numCod;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown numAno;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbPeriodo;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker dateTomarQualify;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.DateTimePicker datePropuestaDoct;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.DateTimePicker datePropuestaMaes;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.DateTimePicker datePedirQualify;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.DateTimePicker dateTema;
    }
}