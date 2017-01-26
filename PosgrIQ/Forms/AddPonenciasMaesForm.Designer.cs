namespace PosgrIQ
{
    partial class AddPonenciasMaesForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddPonenciasMaesForm));
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.cmbAlcance = new System.Windows.Forms.ComboBox();
            this.txtRevista = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.dateAceptado = new System.Windows.Forms.DateTimePicker();
            this.txtTitulo = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.cmbEstudiante = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnAdd
            // 
            this.btnAdd.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnAdd.FlatAppearance.BorderColor = System.Drawing.Color.DarkRed;
            this.btnAdd.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAdd.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAdd.Location = new System.Drawing.Point(547, 172);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 7;
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
            this.btnCancel.Location = new System.Drawing.Point(466, 172);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 8;
            this.btnCancel.Text = "Cancelar";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // cmbAlcance
            // 
            this.cmbAlcance.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbAlcance.FormattingEnabled = true;
            this.cmbAlcance.Items.AddRange(new object[] {
            "Nacional",
            "Internacional"});
            this.cmbAlcance.Location = new System.Drawing.Point(89, 148);
            this.cmbAlcance.Name = "cmbAlcance";
            this.cmbAlcance.Size = new System.Drawing.Size(96, 22);
            this.cmbAlcance.TabIndex = 6;
            // 
            // txtRevista
            // 
            this.txtRevista.Location = new System.Drawing.Point(89, 120);
            this.txtRevista.Name = "txtRevista";
            this.txtRevista.Size = new System.Drawing.Size(533, 22);
            this.txtRevista.TabIndex = 5;
            // 
            // label14
            // 
            this.label14.Font = new System.Drawing.Font("Calibri", 9F);
            this.label14.Location = new System.Drawing.Point(6, 91);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(69, 35);
            this.label14.TabIndex = 59;
            this.label14.Text = "Fecha Aceptado:";
            // 
            // dateAceptado
            // 
            this.dateAceptado.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateAceptado.Location = new System.Drawing.Point(89, 92);
            this.dateAceptado.Name = "dateAceptado";
            this.dateAceptado.Size = new System.Drawing.Size(96, 22);
            this.dateAceptado.TabIndex = 4;
            // 
            // txtTitulo
            // 
            this.txtTitulo.Font = new System.Drawing.Font("Calibri", 9F);
            this.txtTitulo.Location = new System.Drawing.Point(89, 34);
            this.txtTitulo.Multiline = true;
            this.txtTitulo.Name = "txtTitulo";
            this.txtTitulo.Size = new System.Drawing.Size(533, 51);
            this.txtTitulo.TabIndex = 3;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Calibri", 9F);
            this.label13.Location = new System.Drawing.Point(6, 39);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(42, 14);
            this.label13.TabIndex = 57;
            this.label13.Text = "Titulo:";
            // 
            // cmbEstudiante
            // 
            this.cmbEstudiante.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbEstudiante.FormattingEnabled = true;
            this.cmbEstudiante.Location = new System.Drawing.Point(89, 6);
            this.cmbEstudiante.Name = "cmbEstudiante";
            this.cmbEstudiante.Size = new System.Drawing.Size(223, 22);
            this.cmbEstudiante.TabIndex = 2;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Calibri", 9F);
            this.label3.Location = new System.Drawing.Point(6, 152);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(52, 14);
            this.label3.TabIndex = 54;
            this.label3.Text = "Alcance:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Calibri", 9F);
            this.label2.Location = new System.Drawing.Point(6, 126);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 14);
            this.label2.TabIndex = 55;
            this.label2.Text = "Evento:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Calibri", 9F);
            this.label9.Location = new System.Drawing.Point(6, 10);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(69, 14);
            this.label9.TabIndex = 56;
            this.label9.Text = "Estudiante:";
            // 
            // AddPonenciasMaesForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(628, 203);
            this.ControlBox = false;
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.cmbAlcance);
            this.Controls.Add(this.txtRevista);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.dateAceptado);
            this.Controls.Add(this.txtTitulo);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.cmbEstudiante);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label9);
            this.Font = new System.Drawing.Font("Calibri", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(644, 242);
            this.MinimumSize = new System.Drawing.Size(644, 242);
            this.Name = "AddPonenciasMaesForm";
            this.Text = "AGREGAR PONENCIAS DE MAESTRIA";
            this.Load += new System.EventHandler(this.AddPonenciasMaesForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.ComboBox cmbAlcance;
        private System.Windows.Forms.TextBox txtRevista;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.DateTimePicker dateAceptado;
        private System.Windows.Forms.TextBox txtTitulo;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.ComboBox cmbEstudiante;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label9;
    }
}