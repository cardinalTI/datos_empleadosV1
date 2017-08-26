<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.btnimport = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.btnabajas = New System.Windows.Forms.Button()
        Me.btnibajas = New System.Windows.Forms.Button()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.ComboSUELDO = New System.Windows.Forms.ComboBox()
        Me.DataGridView3 = New System.Windows.Forms.DataGridView()
        Me.btnasueldo = New System.Windows.Forms.Button()
        Me.btnisueldo = New System.Windows.Forms.Button()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.Combodepto = New System.Windows.Forms.ComboBox()
        Me.DataGridView5 = New System.Windows.Forms.DataGridView()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TabPage5 = New System.Windows.Forms.TabPage()
        Me.Combopuesto = New System.Windows.Forms.ComboBox()
        Me.DataGridView6 = New System.Windows.Forms.DataGridView()
        Me.btnpuestoa = New System.Windows.Forms.Button()
        Me.btnpuestoi = New System.Windows.Forms.Button()
        Me.TabPage6 = New System.Windows.Forms.TabPage()
        Me.Combobancos = New System.Windows.Forms.ComboBox()
        Me.DataGridView7 = New System.Windows.Forms.DataGridView()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.DataSet11 = New empleados_microsip.DataSet1()
        Me.DataGridView4 = New System.Windows.Forms.DataGridView()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        CType(Me.DataGridView5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage5.SuspendLayout()
        CType(Me.DataGridView6, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage6.SuspendLayout()
        CType(Me.DataGridView7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataSet11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnimport
        '
        Me.btnimport.Location = New System.Drawing.Point(6, 6)
        Me.btnimport.Name = "btnimport"
        Me.btnimport.Size = New System.Drawing.Size(112, 23)
        Me.btnimport.TabIndex = 0
        Me.btnimport.Text = "importar usuarios"
        Me.btnimport.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(3, 34)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(617, 283)
        Me.DataGridView1.TabIndex = 1
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Controls.Add(Me.TabPage5)
        Me.TabControl1.Controls.Add(Me.TabPage6)
        Me.TabControl1.Location = New System.Drawing.Point(0, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(634, 349)
        Me.TabControl1.TabIndex = 3
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage1.Controls.Add(Me.Button5)
        Me.TabPage1.Controls.Add(Me.btnimport)
        Me.TabPage1.Controls.Add(Me.DataGridView1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(626, 323)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Agregar Empleados"
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(135, 4)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(157, 23)
        Me.Button5.TabIndex = 2
        Me.Button5.Text = "Agregar empleados a txt"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage2.Controls.Add(Me.DataGridView2)
        Me.TabPage2.Controls.Add(Me.btnabajas)
        Me.TabPage2.Controls.Add(Me.btnibajas)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(626, 323)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Bajas de empleados"
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Location = New System.Drawing.Point(0, 36)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.Size = New System.Drawing.Size(620, 281)
        Me.DataGridView2.TabIndex = 4
        '
        'btnabajas
        '
        Me.btnabajas.Location = New System.Drawing.Point(133, 7)
        Me.btnabajas.Name = "btnabajas"
        Me.btnabajas.Size = New System.Drawing.Size(108, 23)
        Me.btnabajas.TabIndex = 3
        Me.btnabajas.Text = "Agregar bajas"
        Me.btnabajas.UseVisualStyleBackColor = True
        '
        'btnibajas
        '
        Me.btnibajas.Location = New System.Drawing.Point(6, 6)
        Me.btnibajas.Name = "btnibajas"
        Me.btnibajas.Size = New System.Drawing.Size(110, 23)
        Me.btnibajas.TabIndex = 0
        Me.btnibajas.Text = "Importar bajas"
        Me.btnibajas.UseVisualStyleBackColor = True
        '
        'TabPage3
        '
        Me.TabPage3.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage3.Controls.Add(Me.ComboSUELDO)
        Me.TabPage3.Controls.Add(Me.DataGridView3)
        Me.TabPage3.Controls.Add(Me.btnasueldo)
        Me.TabPage3.Controls.Add(Me.btnisueldo)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(626, 323)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Cambio de sueldo"
        '
        'ComboSUELDO
        '
        Me.ComboSUELDO.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboSUELDO.FormattingEnabled = True
        Me.ComboSUELDO.Items.AddRange(New Object() {"FOLDUR", "MORGET", "GRUPO CONISAL", "WIPSI", "IT TELECOM", "MORGET SEMANAL", "MORGET CATORCENAL", "MORGET QUINCENAL", "MORGET MENSUAL", "MORGET INTERNA", "AICEL", "NUBULA", "CONSORCIO ATERAP SA DE CV", "CROTEC SA DE CV", "PEPSAT SA DE CV", "INFORMATION THECNOLOGY", "UPHETILOLI 2"})
        Me.ComboSUELDO.Location = New System.Drawing.Point(144, 6)
        Me.ComboSUELDO.Name = "ComboSUELDO"
        Me.ComboSUELDO.Size = New System.Drawing.Size(121, 21)
        Me.ComboSUELDO.TabIndex = 3
        '
        'DataGridView3
        '
        Me.DataGridView3.AllowUserToAddRows = False
        Me.DataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView3.Location = New System.Drawing.Point(3, 36)
        Me.DataGridView3.Name = "DataGridView3"
        Me.DataGridView3.Size = New System.Drawing.Size(617, 281)
        Me.DataGridView3.TabIndex = 2
        '
        'btnasueldo
        '
        Me.btnasueldo.Location = New System.Drawing.Point(280, 7)
        Me.btnasueldo.Name = "btnasueldo"
        Me.btnasueldo.Size = New System.Drawing.Size(108, 23)
        Me.btnasueldo.TabIndex = 1
        Me.btnasueldo.Text = "Actualizar sueldos"
        Me.btnasueldo.UseVisualStyleBackColor = True
        '
        'btnisueldo
        '
        Me.btnisueldo.Location = New System.Drawing.Point(6, 6)
        Me.btnisueldo.Name = "btnisueldo"
        Me.btnisueldo.Size = New System.Drawing.Size(108, 23)
        Me.btnisueldo.TabIndex = 0
        Me.btnisueldo.Text = "Importar sueldos"
        Me.btnisueldo.UseVisualStyleBackColor = True
        '
        'TabPage4
        '
        Me.TabPage4.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage4.Controls.Add(Me.Combodepto)
        Me.TabPage4.Controls.Add(Me.DataGridView5)
        Me.TabPage4.Controls.Add(Me.Button2)
        Me.TabPage4.Controls.Add(Me.Button1)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(626, 323)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Agregar Departamentos"
        '
        'Combodepto
        '
        Me.Combodepto.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combodepto.FormattingEnabled = True
        Me.Combodepto.Items.AddRange(New Object() {"FOLDUR", "MORGET", "GRUPO CONISAL", "WIPSI", "IT TELECOM", "MORGET SEMANAL", "MORGET CATORCENAL", "MORGET QUINCENAL", "MORGET MENSUAL", "MORGET INTERNA", "AICEL", "NUBULA", "CONSORCIO ATERAP SA DE CV", "CROTEC SA DE CV", "PEPSAT SA DE CV", "INFORMATION THECNOLOGY", "UPHETILOLI 2"})
        Me.Combodepto.Location = New System.Drawing.Point(224, 7)
        Me.Combodepto.Name = "Combodepto"
        Me.Combodepto.Size = New System.Drawing.Size(154, 21)
        Me.Combodepto.TabIndex = 3
        '
        'DataGridView5
        '
        Me.DataGridView5.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView5.Location = New System.Drawing.Point(6, 35)
        Me.DataGridView5.Name = "DataGridView5"
        Me.DataGridView5.Size = New System.Drawing.Size(615, 271)
        Me.DataGridView5.TabIndex = 2
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(431, 6)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(166, 23)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Agregar Departamentos"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(6, 6)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(199, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Importar Departamentos"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TabPage5
        '
        Me.TabPage5.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage5.Controls.Add(Me.Combopuesto)
        Me.TabPage5.Controls.Add(Me.DataGridView6)
        Me.TabPage5.Controls.Add(Me.btnpuestoa)
        Me.TabPage5.Controls.Add(Me.btnpuestoi)
        Me.TabPage5.ForeColor = System.Drawing.Color.Black
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage5.Size = New System.Drawing.Size(626, 323)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "Agregar puesto"
        '
        'Combopuesto
        '
        Me.Combopuesto.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combopuesto.FormattingEnabled = True
        Me.Combopuesto.Items.AddRange(New Object() {"FOLDUR", "MORGET", "GRUPO CONISAL", "WIPSI", "IT TELECOM", "MORGET SEMANAL", "MORGET CATORCENAL", "MORGET QUINCENAL", "MORGET MENSUAL", "MORGET INTERNA", "AICEL", "NUBULA", "CONSORCIO ATERAP SA DE CV", "CROTEC SA DE CV", "PEPSAT SA DE CV", "INFORMATION THECNOLOGY", "UPHETILOLI 2"})
        Me.Combopuesto.Location = New System.Drawing.Point(218, 7)
        Me.Combopuesto.Name = "Combopuesto"
        Me.Combopuesto.Size = New System.Drawing.Size(153, 21)
        Me.Combopuesto.TabIndex = 3
        '
        'DataGridView6
        '
        Me.DataGridView6.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView6.Location = New System.Drawing.Point(9, 36)
        Me.DataGridView6.Name = "DataGridView6"
        Me.DataGridView6.Size = New System.Drawing.Size(611, 281)
        Me.DataGridView6.TabIndex = 2
        '
        'btnpuestoa
        '
        Me.btnpuestoa.Location = New System.Drawing.Point(428, 7)
        Me.btnpuestoa.Name = "btnpuestoa"
        Me.btnpuestoa.Size = New System.Drawing.Size(153, 23)
        Me.btnpuestoa.TabIndex = 1
        Me.btnpuestoa.Text = "Agregar Puesto"
        Me.btnpuestoa.UseVisualStyleBackColor = True
        '
        'btnpuestoi
        '
        Me.btnpuestoi.Location = New System.Drawing.Point(8, 6)
        Me.btnpuestoi.Name = "btnpuestoi"
        Me.btnpuestoi.Size = New System.Drawing.Size(203, 23)
        Me.btnpuestoi.TabIndex = 0
        Me.btnpuestoi.Text = "Importar puesto"
        Me.btnpuestoi.UseVisualStyleBackColor = True
        '
        'TabPage6
        '
        Me.TabPage6.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage6.Controls.Add(Me.Combobancos)
        Me.TabPage6.Controls.Add(Me.DataGridView7)
        Me.TabPage6.Controls.Add(Me.Button4)
        Me.TabPage6.Controls.Add(Me.Button3)
        Me.TabPage6.Location = New System.Drawing.Point(4, 22)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage6.Size = New System.Drawing.Size(626, 323)
        Me.TabPage6.TabIndex = 5
        Me.TabPage6.Text = "Cuentas de Banco"
        '
        'Combobancos
        '
        Me.Combobancos.FormattingEnabled = True
        Me.Combobancos.Items.AddRange(New Object() {"FOLDUR", "MORGET", "GRUPO CONISAL", "WIPSI", "IT TELECOM", "MORGET SEMANAL", "MORGET CATORCENAL", "MORGET QUINCENAL", "MORGET MENSUAL", "MORGET INTERNA", "AICEL", "NUBULA", "CONSORCIO ATERAP SA DE CV", "CROTEC SA DE CV", "PEUGEOT", "PEPSAT SA DE CV", "INFORMATION THECNOLOGY", "UPHETILOLI 2"})
        Me.Combobancos.Location = New System.Drawing.Point(134, 7)
        Me.Combobancos.Name = "Combobancos"
        Me.Combobancos.Size = New System.Drawing.Size(121, 21)
        Me.Combobancos.TabIndex = 3
        '
        'DataGridView7
        '
        Me.DataGridView7.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView7.Location = New System.Drawing.Point(10, 35)
        Me.DataGridView7.Name = "DataGridView7"
        Me.DataGridView7.Size = New System.Drawing.Size(610, 282)
        Me.DataGridView7.TabIndex = 2
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(279, 5)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(117, 23)
        Me.Button4.TabIndex = 1
        Me.Button4.Text = "Actualizar Cuentas"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(9, 7)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(108, 23)
        Me.Button3.TabIndex = 0
        Me.Button3.Text = "Importar Cuentas"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'DataSet11
        '
        Me.DataSet11.DataSetName = "DataSet1"
        Me.DataSet11.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'DataGridView4
        '
        Me.DataGridView4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView4.Location = New System.Drawing.Point(0, 357)
        Me.DataGridView4.Name = "DataGridView4"
        Me.DataGridView4.Size = New System.Drawing.Size(625, 44)
        Me.DataGridView4.TabIndex = 4
        Me.DataGridView4.Visible = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(644, 354)
        Me.Controls.Add(Me.DataGridView4)
        Me.Controls.Add(Me.TabControl1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.Text = "Modulo de Empleados"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        CType(Me.DataGridView5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage5.ResumeLayout(False)
        CType(Me.DataGridView6, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage6.ResumeLayout(False)
        CType(Me.DataGridView7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataSet11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnimport As System.Windows.Forms.Button
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents btnabajas As System.Windows.Forms.Button
    Friend WithEvents btnibajas As System.Windows.Forms.Button
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents DataGridView3 As System.Windows.Forms.DataGridView
    Friend WithEvents btnasueldo As System.Windows.Forms.Button
    Friend WithEvents btnisueldo As System.Windows.Forms.Button
    Friend WithEvents DataSet11 As empleados_microsip.DataSet1
    Friend WithEvents DataGridView4 As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents DataGridView5 As System.Windows.Forms.DataGridView
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents DataGridView6 As System.Windows.Forms.DataGridView
    Friend WithEvents btnpuestoa As System.Windows.Forms.Button
    Friend WithEvents btnpuestoi As System.Windows.Forms.Button
    Friend WithEvents TabPage6 As System.Windows.Forms.TabPage
    Friend WithEvents DataGridView7 As System.Windows.Forms.DataGridView
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ComboSUELDO As System.Windows.Forms.ComboBox
    Friend WithEvents Combobancos As System.Windows.Forms.ComboBox
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Combodepto As System.Windows.Forms.ComboBox
    Friend WithEvents Combopuesto As System.Windows.Forms.ComboBox

End Class
