namespace PosgrIQ
{
    partial class ProfesoresForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ProfesoresForm));
            this.dataGridProfesores = new System.Windows.Forms.DataGridView();
            this.btnCerrar = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridProfesores)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridProfesores
            // 
            this.dataGridProfesores.AllowUserToAddRows = false;
            this.dataGridProfesores.AllowUserToDeleteRows = false;
            this.dataGridProfesores.AllowUserToResizeRows = false;
            this.dataGridProfesores.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridProfesores.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dataGridProfesores.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dataGridProfesores.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridProfesores.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dataGridProfesores.GridColor = System.Drawing.SystemColors.ControlLightLight;
            this.dataGridProfesores.Location = new System.Drawing.Point(10, 10);
            this.dataGridProfesores.MultiSelect = false;
            this.dataGridProfesores.Name = "dataGridProfesores";
            this.dataGridProfesores.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridProfesores.Size = new System.Drawing.Size(450, 200);
            this.dataGridProfesores.TabIndex = 0;
            // 
            // btnCerrar
            // 
            this.btnCerrar.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnCerrar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCerrar.Location = new System.Drawing.Point(389, 222);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(75, 23);
            this.btnCerrar.TabIndex = 1;
            this.btnCerrar.Text = "Cerrar";
            this.btnCerrar.UseVisualStyleBackColor = false;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Location = new System.Drawing.Point(308, 222);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "Profesores";
            this.button1.UseVisualStyleBackColor = false;
            // 
            // ProfesoresForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(470, 251);
            this.ControlBox = false;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnCerrar);
            this.Controls.Add(this.dataGridProfesores);
            this.Font = new System.Drawing.Font("Calibri", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ProfesoresForm";
            this.Text = "LISTADO DE PROFESORES ASOCIADOS";
            this.Load += new System.EventHandler(this.ProfesoresForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridProfesores)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridProfesores;
        private System.Windows.Forms.Button btnCerrar;
        private System.Windows.Forms.Button button1;

    }
}