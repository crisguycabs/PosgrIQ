namespace PosgrIQ
{
    partial class AddEstudiantesDoctForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddEstudiantesDoctForm));
            this.label2 = new System.Windows.Forms.Label();
            this.numCod = new System.Windows.Forms.NumericUpDown();
            this.txtNombre = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.txtCorreo = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtCiudad = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.numCedula = new System.Windows.Forms.NumericUpDown();
            this.cmbCondicion = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.cmbDirector = new System.Windows.Forms.ComboBox();
            this.btnAddProfesor = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.cmbCodirector1 = new System.Windows.Forms.ComboBox();
            this.label11 = new System.Windows.Forms.Label();
            this.cmbCodirector2 = new System.Windows.Forms.ComboBox();
            this.label12 = new System.Windows.Forms.Label();
            this.cmbReglamentos = new System.Windows.Forms.ComboBox();
            this.btnAddReglamento = new System.Windows.Forms.Button();
            this.label13 = new System.Windows.Forms.Label();
            this.txtTema = new System.Windows.Forms.TextBox();
            this.dateTema = new System.Windows.Forms.DateTimePicker();
            this.cmbConceptoTema = new System.Windows.Forms.ComboBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.txtRutaTema = new System.Windows.Forms.TextBox();
            this.btnRutaTema = new System.Windows.Forms.Button();
            this.label17 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.cmbSolicitarQualify = new System.Windows.Forms.ComboBox();
            this.cmbAprobarQualify = new System.Windows.Forms.ComboBox();
            this.btnVerArchivoTema = new System.Windows.Forms.Button();
            this.chkTema = new System.Windows.Forms.CheckBox();
            this.label21 = new System.Windows.Forms.Label();
            this.txtNivel = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.numCod)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numCedula)).BeginInit();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Calibri", 9F);
            this.label2.Location = new System.Drawing.Point(6, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 14);
            this.label2.TabIndex = 14;
            this.label2.Text = "Nombre:";
            // 
            // numCod
            // 
            this.numCod.Font = new System.Drawing.Font("Calibri", 9F);
            this.numCod.Location = new System.Drawing.Point(84, 8);
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
            this.txtNombre.Location = new System.Drawing.Point(84, 35);
            this.txtNombre.Name = "txtNombre";
            this.txtNombre.Size = new System.Drawing.Size(223, 22);
            this.txtNombre.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 9F);
            this.label1.Location = new System.Drawing.Point(6, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 14);
            this.label1.TabIndex = 11;
            this.label1.Text = "Codigo:";
            // 
            // label5
            // 
            this.label5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label5.Location = new System.Drawing.Point(8, 119);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(595, 2);
            this.label5.TabIndex = 15;
            // 
            // btnAdd
            // 
            this.btnAdd.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnAdd.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAdd.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAdd.Location = new System.Drawing.Point(528, 400);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 24;
            this.btnAdd.Text = "Agregar";
            this.btnAdd.UseVisualStyleBackColor = false;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancel.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnCancel.Location = new System.Drawing.Point(447, 400);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 23;
            this.btnCancel.Text = "Cancelar";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // txtCorreo
            // 
            this.txtCorreo.Font = new System.Drawing.Font("Calibri", 9F);
            this.txtCorreo.Location = new System.Drawing.Point(84, 62);
            this.txtCorreo.Name = "txtCorreo";
            this.txtCorreo.Size = new System.Drawing.Size(223, 22);
            this.txtCorreo.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Calibri", 9F);
            this.label3.Location = new System.Drawing.Point(6, 66);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(45, 14);
            this.label3.TabIndex = 14;
            this.label3.Text = "Correo:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Calibri", 9F);
            this.label4.Location = new System.Drawing.Point(325, 12);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(51, 14);
            this.label4.TabIndex = 14;
            this.label4.Text = "Cedula: ";
            // 
            // txtCiudad
            // 
            this.txtCiudad.Font = new System.Drawing.Font("Calibri", 9F);
            this.txtCiudad.Location = new System.Drawing.Point(378, 35);
            this.txtCiudad.Name = "txtCiudad";
            this.txtCiudad.Size = new System.Drawing.Size(223, 22);
            this.txtCiudad.TabIndex = 5;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Calibri", 9F);
            this.label6.Location = new System.Drawing.Point(325, 39);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(48, 14);
            this.label6.TabIndex = 14;
            this.label6.Text = "Ciudad:";
            // 
            // numCedula
            // 
            this.numCedula.Font = new System.Drawing.Font("Calibri", 9F);
            this.numCedula.Location = new System.Drawing.Point(378, 8);
            this.numCedula.Maximum = new decimal(new int[] {
            1410065407,
            2,
            0,
            0});
            this.numCedula.Name = "numCedula";
            this.numCedula.Size = new System.Drawing.Size(82, 22);
            this.numCedula.TabIndex = 4;
            this.numCedula.Value = new decimal(new int[] {
            1410065407,
            2,
            0,
            0});
            // 
            // cmbCondicion
            // 
            this.cmbCondicion.FormattingEnabled = true;
            this.cmbCondicion.Location = new System.Drawing.Point(512, 62);
            this.cmbCondicion.Name = "cmbCondicion";
            this.cmbCondicion.Size = new System.Drawing.Size(89, 22);
            this.cmbCondicion.TabIndex = 7;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Calibri", 9F);
            this.label7.Location = new System.Drawing.Point(444, 66);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(67, 14);
            this.label7.TabIndex = 14;
            this.label7.Text = "Condicion: ";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Calibri", 9F);
            this.label8.Location = new System.Drawing.Point(325, 66);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(41, 14);
            this.label8.TabIndex = 14;
            this.label8.Text = "Nivel: ";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Calibri", 9F);
            this.label9.Location = new System.Drawing.Point(6, 131);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(56, 14);
            this.label9.TabIndex = 14;
            this.label9.Text = "Director: ";
            // 
            // cmbDirector
            // 
            this.cmbDirector.FormattingEnabled = true;
            this.cmbDirector.Location = new System.Drawing.Point(84, 127);
            this.cmbDirector.Name = "cmbDirector";
            this.cmbDirector.Size = new System.Drawing.Size(223, 22);
            this.cmbDirector.TabIndex = 10;
            // 
            // btnAddProfesor
            // 
            this.btnAddProfesor.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnAddProfesor.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAddProfesor.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAddProfesor.Location = new System.Drawing.Point(179, 155);
            this.btnAddProfesor.Name = "btnAddProfesor";
            this.btnAddProfesor.Size = new System.Drawing.Size(128, 23);
            this.btnAddProfesor.TabIndex = 13;
            this.btnAddProfesor.Text = "Agregar Profesor";
            this.btnAddProfesor.UseVisualStyleBackColor = false;
            this.btnAddProfesor.Click += new System.EventHandler(this.btnAddProfesor_Click);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Calibri", 9F);
            this.label10.Location = new System.Drawing.Point(325, 131);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(74, 14);
            this.label10.TabIndex = 14;
            this.label10.Text = "Codirector 1:";
            // 
            // cmbCodirector1
            // 
            this.cmbCodirector1.FormattingEnabled = true;
            this.cmbCodirector1.Location = new System.Drawing.Point(406, 127);
            this.cmbCodirector1.Name = "cmbCodirector1";
            this.cmbCodirector1.Size = new System.Drawing.Size(195, 22);
            this.cmbCodirector1.TabIndex = 11;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Calibri", 9F);
            this.label11.Location = new System.Drawing.Point(325, 160);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(74, 14);
            this.label11.TabIndex = 14;
            this.label11.Text = "Codirector 2:";
            // 
            // cmbCodirector2
            // 
            this.cmbCodirector2.FormattingEnabled = true;
            this.cmbCodirector2.Location = new System.Drawing.Point(406, 156);
            this.cmbCodirector2.Name = "cmbCodirector2";
            this.cmbCodirector2.Size = new System.Drawing.Size(195, 22);
            this.cmbCodirector2.TabIndex = 12;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Calibri", 9F);
            this.label12.Location = new System.Drawing.Point(6, 93);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(79, 14);
            this.label12.TabIndex = 14;
            this.label12.Text = "Reglamento: ";
            // 
            // cmbReglamentos
            // 
            this.cmbReglamentos.FormattingEnabled = true;
            this.cmbReglamentos.Location = new System.Drawing.Point(84, 89);
            this.cmbReglamentos.Name = "cmbReglamentos";
            this.cmbReglamentos.Size = new System.Drawing.Size(89, 22);
            this.cmbReglamentos.TabIndex = 8;
            // 
            // btnAddReglamento
            // 
            this.btnAddReglamento.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnAddReglamento.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAddReglamento.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAddReglamento.Location = new System.Drawing.Point(179, 89);
            this.btnAddReglamento.Name = "btnAddReglamento";
            this.btnAddReglamento.Size = new System.Drawing.Size(128, 23);
            this.btnAddReglamento.TabIndex = 9;
            this.btnAddReglamento.Text = "Agregar Reglamento";
            this.btnAddReglamento.UseVisualStyleBackColor = false;
            this.btnAddReglamento.Click += new System.EventHandler(this.btnAddReglamento_Click);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Calibri", 9F);
            this.label13.Location = new System.Drawing.Point(5, 220);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(39, 14);
            this.label13.TabIndex = 14;
            this.label13.Text = "Tema:";
            // 
            // txtTema
            // 
            this.txtTema.Enabled = false;
            this.txtTema.Font = new System.Drawing.Font("Calibri", 9F);
            this.txtTema.Location = new System.Drawing.Point(83, 215);
            this.txtTema.Multiline = true;
            this.txtTema.Name = "txtTema";
            this.txtTema.Size = new System.Drawing.Size(518, 51);
            this.txtTema.TabIndex = 15;
            // 
            // dateTema
            // 
            this.dateTema.Enabled = false;
            this.dateTema.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTema.Location = new System.Drawing.Point(83, 272);
            this.dateTema.Name = "dateTema";
            this.dateTema.Size = new System.Drawing.Size(96, 22);
            this.dateTema.TabIndex = 16;
            // 
            // cmbConceptoTema
            // 
            this.cmbConceptoTema.Enabled = false;
            this.cmbConceptoTema.FormattingEnabled = true;
            this.cmbConceptoTema.Location = new System.Drawing.Point(263, 272);
            this.cmbConceptoTema.Name = "cmbConceptoTema";
            this.cmbConceptoTema.Size = new System.Drawing.Size(89, 22);
            this.cmbConceptoTema.TabIndex = 17;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Calibri", 9F);
            this.label14.Location = new System.Drawing.Point(5, 276);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(42, 14);
            this.label14.TabIndex = 14;
            this.label14.Text = "Fecha:";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Calibri", 9F);
            this.label15.Location = new System.Drawing.Point(202, 276);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(60, 14);
            this.label15.TabIndex = 14;
            this.label15.Text = "Concepto:";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Calibri", 9F);
            this.label16.Location = new System.Drawing.Point(5, 303);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(72, 14);
            this.label16.TabIndex = 19;
            this.label16.Text = "Documento:";
            // 
            // txtRutaTema
            // 
            this.txtRutaTema.Enabled = false;
            this.txtRutaTema.Font = new System.Drawing.Font("Calibri", 9F);
            this.txtRutaTema.Location = new System.Drawing.Point(83, 300);
            this.txtRutaTema.Name = "txtRutaTema";
            this.txtRutaTema.Size = new System.Drawing.Size(518, 22);
            this.txtRutaTema.TabIndex = 18;
            // 
            // btnRutaTema
            // 
            this.btnRutaTema.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnRutaTema.Enabled = false;
            this.btnRutaTema.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRutaTema.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnRutaTema.Location = new System.Drawing.Point(83, 328);
            this.btnRutaTema.Name = "btnRutaTema";
            this.btnRutaTema.Size = new System.Drawing.Size(128, 23);
            this.btnRutaTema.TabIndex = 19;
            this.btnRutaTema.Text = "Seleccionar Archivo";
            this.btnRutaTema.UseVisualStyleBackColor = false;
            this.btnRutaTema.Click += new System.EventHandler(this.btnRutaTema_Click);
            // 
            // label17
            // 
            this.label17.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label17.Location = new System.Drawing.Point(8, 358);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(595, 2);
            this.label17.TabIndex = 20;
            // 
            // label18
            // 
            this.label18.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label18.Location = new System.Drawing.Point(8, 393);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(595, 2);
            this.label18.TabIndex = 15;
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Font = new System.Drawing.Font("Calibri", 9F);
            this.label19.Location = new System.Drawing.Point(6, 369);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(93, 14);
            this.label19.TabIndex = 14;
            this.label19.Text = "Solicitó Qualify:";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Font = new System.Drawing.Font("Calibri", 9F);
            this.label20.Location = new System.Drawing.Point(325, 369);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(91, 14);
            this.label20.TabIndex = 14;
            this.label20.Text = "Aprobó Qualify:";
            // 
            // cmbSolicitarQualify
            // 
            this.cmbSolicitarQualify.FormattingEnabled = true;
            this.cmbSolicitarQualify.Items.AddRange(new object[] {
            "No",
            "Si"});
            this.cmbSolicitarQualify.Location = new System.Drawing.Point(105, 365);
            this.cmbSolicitarQualify.Name = "cmbSolicitarQualify";
            this.cmbSolicitarQualify.Size = new System.Drawing.Size(38, 22);
            this.cmbSolicitarQualify.TabIndex = 21;
            // 
            // cmbAprobarQualify
            // 
            this.cmbAprobarQualify.FormattingEnabled = true;
            this.cmbAprobarQualify.Items.AddRange(new object[] {
            "No",
            "Si"});
            this.cmbAprobarQualify.Location = new System.Drawing.Point(422, 365);
            this.cmbAprobarQualify.Name = "cmbAprobarQualify";
            this.cmbAprobarQualify.Size = new System.Drawing.Size(38, 22);
            this.cmbAprobarQualify.TabIndex = 22;
            // 
            // btnVerArchivoTema
            // 
            this.btnVerArchivoTema.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnVerArchivoTema.Enabled = false;
            this.btnVerArchivoTema.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnVerArchivoTema.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnVerArchivoTema.Location = new System.Drawing.Point(217, 328);
            this.btnVerArchivoTema.Name = "btnVerArchivoTema";
            this.btnVerArchivoTema.Size = new System.Drawing.Size(128, 23);
            this.btnVerArchivoTema.TabIndex = 20;
            this.btnVerArchivoTema.Text = "Ver Archivo";
            this.btnVerArchivoTema.UseVisualStyleBackColor = false;
            this.btnVerArchivoTema.Click += new System.EventHandler(this.btnVerArchivoTema_Click);
            // 
            // chkTema
            // 
            this.chkTema.AutoSize = true;
            this.chkTema.Location = new System.Drawing.Point(8, 193);
            this.chkTema.Name = "chkTema";
            this.chkTema.Size = new System.Drawing.Size(113, 18);
            this.chkTema.TabIndex = 14;
            this.chkTema.Text = "Tema Entregado";
            this.chkTema.UseVisualStyleBackColor = true;
            this.chkTema.CheckedChanged += new System.EventHandler(this.chkTema_CheckedChanged);
            // 
            // label21
            // 
            this.label21.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label21.Location = new System.Drawing.Point(8, 185);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(595, 2);
            this.label21.TabIndex = 21;
            // 
            // txtNivel
            // 
            this.txtNivel.Enabled = false;
            this.txtNivel.Font = new System.Drawing.Font("Calibri", 9F);
            this.txtNivel.Location = new System.Drawing.Point(378, 62);
            this.txtNivel.Name = "txtNivel";
            this.txtNivel.Size = new System.Drawing.Size(43, 22);
            this.txtNivel.TabIndex = 2;
            // 
            // AddEstudiantesDoctForm
            // 
            this.AcceptButton = this.btnAdd;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(608, 429);
            this.ControlBox = false;
            this.Controls.Add(this.label21);
            this.Controls.Add(this.chkTema);
            this.Controls.Add(this.txtTema);
            this.Controls.Add(this.txtRutaTema);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.cmbAprobarQualify);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.cmbSolicitarQualify);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.dateTema);
            this.Controls.Add(this.cmbCodirector2);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.cmbCodirector1);
            this.Controls.Add(this.btnRutaTema);
            this.Controls.Add(this.cmbDirector);
            this.Controls.Add(this.btnVerArchivoTema);
            this.Controls.Add(this.cmbReglamentos);
            this.Controls.Add(this.cmbConceptoTema);
            this.Controls.Add(this.cmbCondicion);
            this.Controls.Add(this.btnAddReglamento);
            this.Controls.Add(this.btnAddProfesor);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.label18);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label20);
            this.Controls.Add(this.label19);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.numCedula);
            this.Controls.Add(this.numCod);
            this.Controls.Add(this.txtCorreo);
            this.Controls.Add(this.txtCiudad);
            this.Controls.Add(this.txtNivel);
            this.Controls.Add(this.txtNombre);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Calibri", 9F);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "AddEstudiantesDoctForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "AGREGAR ESTUDIANTE DE DOCTORADO";
            this.Load += new System.EventHandler(this.AddEstudiantesDoctForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numCod)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numCedula)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown numCod;
        private System.Windows.Forms.TextBox txtNombre;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.TextBox txtCorreo;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtCiudad;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.NumericUpDown numCedula;
        private System.Windows.Forms.ComboBox cmbCondicion;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox cmbDirector;
        private System.Windows.Forms.Button btnAddProfesor;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ComboBox cmbCodirector1;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.ComboBox cmbCodirector2;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.ComboBox cmbReglamentos;
        private System.Windows.Forms.Button btnAddReglamento;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txtTema;
        private System.Windows.Forms.DateTimePicker dateTema;
        private System.Windows.Forms.ComboBox cmbConceptoTema;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.TextBox txtRutaTema;
        private System.Windows.Forms.Button btnRutaTema;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.ComboBox cmbSolicitarQualify;
        private System.Windows.Forms.ComboBox cmbAprobarQualify;
        private System.Windows.Forms.Button btnVerArchivoTema;
        private System.Windows.Forms.CheckBox chkTema;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.TextBox txtNivel;
    }
}