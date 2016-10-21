namespace PosgrIQ
{
    partial class AddProfesoresForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddProfesoresForm));
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.numCod = new System.Windows.Forms.NumericUpDown();
            this.txtNombre = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.cmbColegiatura = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbEscuela = new System.Windows.Forms.ComboBox();
            this.txtCorreo = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtTelefono = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtUniversidad = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.btnAddEscuela = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.numCod)).BeginInit();
            this.SuspendLayout();
            // 
            // btnAdd
            // 
            this.btnAdd.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnAdd.FlatAppearance.BorderColor = System.Drawing.Color.DarkRed;
            this.btnAdd.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAdd.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAdd.Location = new System.Drawing.Point(222, 234);
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
            this.btnCancel.Location = new System.Drawing.Point(141, 234);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 9;
            this.btnCancel.Text = "Cancelar";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Calibri", 9F);
            this.label2.Location = new System.Drawing.Point(3, 37);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 14);
            this.label2.TabIndex = 10;
            this.label2.Text = "Nombre:";
            // 
            // numCod
            // 
            this.numCod.Font = new System.Drawing.Font("Calibri", 9F);
            this.numCod.Location = new System.Drawing.Point(80, 5);
            this.numCod.Maximum = new decimal(new int[] {
            10000000,
            0,
            0,
            0});
            this.numCod.Name = "numCod";
            this.numCod.Size = new System.Drawing.Size(70, 22);
            this.numCod.TabIndex = 1;
            this.numCod.Value = new decimal(new int[] {
            9999999,
            0,
            0,
            0});
            // 
            // txtNombre
            // 
            this.txtNombre.Font = new System.Drawing.Font("Calibri", 9F);
            this.txtNombre.Location = new System.Drawing.Point(80, 33);
            this.txtNombre.Name = "txtNombre";
            this.txtNombre.Size = new System.Drawing.Size(217, 22);
            this.txtNombre.TabIndex = 2;
            this.txtNombre.Text = "Ingenieria de Sistemas e Informatica";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 9F);
            this.label1.Location = new System.Drawing.Point(3, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 14);
            this.label1.TabIndex = 7;
            this.label1.Text = "Codigo:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Calibri", 9F);
            this.label3.Location = new System.Drawing.Point(3, 64);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(73, 14);
            this.label3.TabIndex = 10;
            this.label3.Text = "Colegiatura:";
            // 
            // cmbColegiatura
            // 
            this.cmbColegiatura.FormattingEnabled = true;
            this.cmbColegiatura.Location = new System.Drawing.Point(80, 61);
            this.cmbColegiatura.Name = "cmbColegiatura";
            this.cmbColegiatura.Size = new System.Drawing.Size(121, 22);
            this.cmbColegiatura.TabIndex = 3;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Calibri", 9F);
            this.label4.Location = new System.Drawing.Point(3, 92);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(52, 14);
            this.label4.TabIndex = 10;
            this.label4.Text = "Escuela:";
            // 
            // cmbEscuela
            // 
            this.cmbEscuela.FormattingEnabled = true;
            this.cmbEscuela.Location = new System.Drawing.Point(80, 89);
            this.cmbEscuela.Name = "cmbEscuela";
            this.cmbEscuela.Size = new System.Drawing.Size(217, 22);
            this.cmbEscuela.TabIndex = 4;
            // 
            // txtCorreo
            // 
            this.txtCorreo.Font = new System.Drawing.Font("Calibri", 9F);
            this.txtCorreo.Location = new System.Drawing.Point(80, 146);
            this.txtCorreo.Name = "txtCorreo";
            this.txtCorreo.Size = new System.Drawing.Size(217, 22);
            this.txtCorreo.TabIndex = 6;
            this.txtCorreo.Text = "Ingenieria de Sistemas e Informatica";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Calibri", 9F);
            this.label5.Location = new System.Drawing.Point(3, 150);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(45, 14);
            this.label5.TabIndex = 10;
            this.label5.Text = "Correo:";
            // 
            // txtTelefono
            // 
            this.txtTelefono.Font = new System.Drawing.Font("Calibri", 9F);
            this.txtTelefono.Location = new System.Drawing.Point(80, 174);
            this.txtTelefono.Name = "txtTelefono";
            this.txtTelefono.Size = new System.Drawing.Size(217, 22);
            this.txtTelefono.TabIndex = 7;
            this.txtTelefono.Text = "Ingenieria de Sistemas e Informatica";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Calibri", 9F);
            this.label6.Location = new System.Drawing.Point(3, 178);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(58, 14);
            this.label6.TabIndex = 10;
            this.label6.Text = "Telefono:";
            // 
            // txtUniversidad
            // 
            this.txtUniversidad.Font = new System.Drawing.Font("Calibri", 9F);
            this.txtUniversidad.Location = new System.Drawing.Point(80, 202);
            this.txtUniversidad.Name = "txtUniversidad";
            this.txtUniversidad.Size = new System.Drawing.Size(217, 22);
            this.txtUniversidad.TabIndex = 8;
            this.txtUniversidad.Text = "Ingenieria de Sistemas e Informatica";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Calibri", 9F);
            this.label7.Location = new System.Drawing.Point(3, 206);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(76, 14);
            this.label7.TabIndex = 10;
            this.label7.Text = "Universidad:";
            // 
            // btnAddEscuela
            // 
            this.btnAddEscuela.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnAddEscuela.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAddEscuela.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAddEscuela.Location = new System.Drawing.Point(192, 117);
            this.btnAddEscuela.Name = "btnAddEscuela";
            this.btnAddEscuela.Size = new System.Drawing.Size(105, 23);
            this.btnAddEscuela.TabIndex = 5;
            this.btnAddEscuela.Text = "Agregar Escuela";
            this.btnAddEscuela.UseVisualStyleBackColor = false;
            this.btnAddEscuela.Click += new System.EventHandler(this.btnAddEscuela_Click);
            // 
            // AddProfesoresForm
            // 
            this.AcceptButton = this.btnAdd;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(305, 264);
            this.ControlBox = false;
            this.Controls.Add(this.cmbEscuela);
            this.Controls.Add(this.cmbColegiatura);
            this.Controls.Add(this.btnAddEscuela);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.numCod);
            this.Controls.Add(this.txtUniversidad);
            this.Controls.Add(this.txtTelefono);
            this.Controls.Add(this.txtCorreo);
            this.Controls.Add(this.txtNombre);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Calibri", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "AddProfesoresForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "AddProfesoresForm";
            this.Load += new System.EventHandler(this.AddProfesoresForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numCod)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown numCod;
        private System.Windows.Forms.TextBox txtNombre;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmbColegiatura;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cmbEscuela;
        private System.Windows.Forms.TextBox txtCorreo;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtTelefono;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtUniversidad;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button btnAddEscuela;
    }
}