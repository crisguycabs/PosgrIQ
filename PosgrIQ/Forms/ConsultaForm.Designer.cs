namespace PosgrIQ
{
    partial class ConsultaForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConsultaForm));
            this.dataGridConsulta = new System.Windows.Forms.DataGridView();
            this.btnCerrar = new System.Windows.Forms.Button();
            this.dateIni = new System.Windows.Forms.DateTimePicker();
            this.btnConsultar = new System.Windows.Forms.Button();
            this.dateFin = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.comboConsulta = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridConsulta)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridConsulta
            // 
            this.dataGridConsulta.AllowUserToAddRows = false;
            this.dataGridConsulta.AllowUserToDeleteRows = false;
            this.dataGridConsulta.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridConsulta.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dataGridConsulta.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dataGridConsulta.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridConsulta.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dataGridConsulta.Location = new System.Drawing.Point(8, 84);
            this.dataGridConsulta.Name = "dataGridConsulta";
            this.dataGridConsulta.Size = new System.Drawing.Size(757, 293);
            this.dataGridConsulta.TabIndex = 1;
            // 
            // btnCerrar
            // 
            this.btnCerrar.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnCerrar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCerrar.FlatAppearance.BorderColor = System.Drawing.Color.DarkRed;
            this.btnCerrar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCerrar.Location = new System.Drawing.Point(690, 384);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(75, 23);
            this.btnCerrar.TabIndex = 5;
            this.btnCerrar.Text = "Cerrar";
            this.btnCerrar.UseVisualStyleBackColor = false;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
            // 
            // dateIni
            // 
            this.dateIni.Location = new System.Drawing.Point(64, 44);
            this.dateIni.Margin = new System.Windows.Forms.Padding(4);
            this.dateIni.Name = "dateIni";
            this.dateIni.Size = new System.Drawing.Size(250, 22);
            this.dateIni.TabIndex = 1;
            // 
            // btnConsultar
            // 
            this.btnConsultar.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnConsultar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnConsultar.FlatAppearance.BorderColor = System.Drawing.Color.DarkRed;
            this.btnConsultar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnConsultar.Location = new System.Drawing.Point(671, 44);
            this.btnConsultar.Name = "btnConsultar";
            this.btnConsultar.Size = new System.Drawing.Size(75, 23);
            this.btnConsultar.TabIndex = 6;
            this.btnConsultar.Text = "Consultar";
            this.btnConsultar.UseVisualStyleBackColor = false;
            this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
            // 
            // dateFin
            // 
            this.dateFin.Location = new System.Drawing.Point(380, 44);
            this.dateFin.Margin = new System.Windows.Forms.Padding(4);
            this.dateFin.Name = "dateFin";
            this.dateFin.Size = new System.Drawing.Size(250, 22);
            this.dateFin.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 48);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(48, 14);
            this.label1.TabIndex = 8;
            this.label1.Text = "Desde: ";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(325, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(45, 14);
            this.label2.TabIndex = 9;
            this.label2.Text = "Hasta: ";
            // 
            // comboConsulta
            // 
            this.comboConsulta.FormattingEnabled = true;
            this.comboConsulta.Items.AddRange(new object[] {
            "Fecha de entrega del Tema de Maestria",
            "Fecha de entrega del Tema de Doctorado",
            "Fecha de entrega de la Propuesta de Maestria",
            "Fecha de entrega de las correcciones de la Propuesta de Maestria",
            "Fecha de sustentacion de la Propuesta de Maestria",
            "Fecha de entrega de la Propuesta de Doctorado",
            "Fecha de entrega de las correcciones de la Propuesta de Doctorado",
            "Fecha de sustentacion de la Propuesta de Doctorado",
            "Fecha de entrega del Trabajo de Maestria",
            "Fecha de entrega de las correcciones del Trabajo de Maestria",
            "Fecha de sustentacion del Trabajo de Maestria",
            "Fecha de entrega de la Tesis de Doctorado",
            "Fecha de entrega de las correcciones de la Tesis de Doctorado",
            "Fecha de sustentacion de la Tesis de Doctorado"});
            this.comboConsulta.Location = new System.Drawing.Point(12, 12);
            this.comboConsulta.Name = "comboConsulta";
            this.comboConsulta.Size = new System.Drawing.Size(302, 22);
            this.comboConsulta.TabIndex = 10;
            // 
            // ConsultaForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(771, 415);
            this.ControlBox = false;
            this.Controls.Add(this.comboConsulta);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnCerrar);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dataGridConsulta);
            this.Controls.Add(this.dateFin);
            this.Controls.Add(this.btnConsultar);
            this.Controls.Add(this.dateIni);
            this.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ConsultaForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CONSULTAS";
            this.Load += new System.EventHandler(this.ConsultaForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridConsulta)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridConsulta;
        private System.Windows.Forms.Button btnCerrar;
        private System.Windows.Forms.DateTimePicker dateIni;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dateFin;
        private System.Windows.Forms.Button btnConsultar;
        private System.Windows.Forms.ComboBox comboConsulta;
    }
}