Imports System.Data.OleDb
Imports System.Data.Odbc
Imports FirebirdSql.Data.FirebirdClient
Imports System.Data.SqlClient

Public Class Form1
    Private cdatos As datoshandler
    Private arreDatos As ArrayList

    Dim Conexion As New FbConnection
    Private m_con As String
    Private m_ConnODBC As OdbcConnection


    Private Sub btnimport_Click(sender As System.Object, e As System.EventArgs) Handles btnimport.Click

        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        Try
            Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
            Dim cnConex As New OleDbConnection(stConexion)
            Dim Cmd As New OleDbCommand("Select  [numero],[nombre],[apellido_paterno],[apellido_materno],[clave_puesto],[clave_depto],[clave_frecuencia_pago],[num_reg_patronal],[forma_pago],[contrato],[jornada],[regimen_fiscal],[fecha_ingreso],[estatus],[tipo_salario],[salario_diario],[salario_integrado],[rfc],[curp],[reg_imss],[dirección],[cp],[contrato_sat] From [Hoja1$]")
            Dim Ds As New DataSet
            Dim Da As New OleDbDataAdapter
            Dim Dt As New DataTable
            cnConex.Open()
            Cmd.Connection = cnConex
            Da.SelectCommand = Cmd
            Da.Fill(Ds)
            Dt = Ds.Tables(0)
            Me.DataGridView1.Columns.Clear()
            Me.DataGridView1.DataSource = Dt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub




    Private Sub Agregar_Click(sender As System.Object, e As System.EventArgs)

        Me.delete()
        Try

            For i As Integer = 0 To Me.DataGridView1.Rows.Count - 1
                With Me.DataGridView1.Rows(i)

                    If .Cells(0) Is DBNull.Value Then
                        MsgBox("No hay mas datos")
                    End If

                    Me.Agregausuario(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, .Cells(3).Value, .Cells(4).Value, .Cells(5).Value, .Cells(6).Value, .Cells(7).Value, .Cells(8).Value, .Cells(9).Value, .Cells(10).Value, .Cells(11).Value, .Cells(12).Value, .Cells(13).Value, .Cells(14).Value, .Cells(15).Value, .Cells(16).Value, .Cells(17).Value, .Cells(18).Value, .Cells(19).Value, .Cells(20).Value, .Cells(21).Value, .Cells(22).Value)
                    'empleadoid = Me.buscardatos(.Cells(0).Value)
                    'rol = Me.buscarrol(.Cells())
                    'If empleadoid <> "" Then

                    '    Me.Agregarclave(empleadoid, ("I" + .Cells(0).Value), rol)
                    'End If
                End With
            Next

            MsgBox("usuarios almacenados correctamente")


        Catch ex As Exception
            MsgBox("Error numero 1")
        End Try

        Muestradatos()
        MsgBox("El txt fue creado correctamente")
    End Sub

    Public Sub delete()



        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= microsip; user id=sicossadmi;password=ipp2012;")

        'Dim DBCon As New SqlConnection("data source= 192.168.2.83; Initial Catalog= microsip;user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= microsip;user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("Server=USUARIO-PC\SQLEXPRESS;Database=incidencias;integrated security=true")


        Dim consulta As String
        consulta = ("delete from usuario")


        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                End With
                DBCon.Open()

                comm.ExecuteNonQuery()

            End Using

        Catch ex As SqlException

            MsgBox("ERROR EN LA CONEXION")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try
    End Sub


    Public Sub Agregausuario(ByVal numero As String, ByVal nombre As String, ByVal apellido_paterno As String, ByVal apellido_materno As String, ByVal clave_puesto As String, ByVal clave_depto As String, ByVal clave_frecuencia_pago As String,
                              ByVal num_reg_patronal As String, ByVal forma_pago As String, ByVal contrato As String, ByVal jornada As String, ByVal regimen_fiscal As String,
                                 ByVal fecha_ingreso As String, ByVal estatus As String, ByVal tipo_salario As String, ByVal salario_diario As String, ByVal salario_integrado As String,
                                 ByVal rfc As String, ByVal curp As String, ByVal registro_imss As String, ByVal direccion As String, ByVal cp As String, ByVal contratosat As String)


        Dim DBCon As New SqlConnection("data source= 192.168.2.82; Initial Catalog= microsip; user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 192.168.2.83; Initial Catalog= microsip;user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("data source= 189.190.172.169; Initial Catalog= microsip;user id=sicossadmi;password=ipp2012;")
        'Dim DBCon As New SqlConnection("Server=USUARIO-PC\SQLEXPRESS;Database=incidencias;integrated security=true")



        Dim consulta As String
        consulta = ("insert into usuario (numero,nombres,apellido_paterno,apellido_materno,clave_puesto,clave_depto,clave_frecuencia_pago,num_reg_patronal,forma_pago,contrato,jornada,regimen_fiscal,fecha_ingreso,estatus,tipo_salario,salario_diario,salario_integrado,rfc,curp,registro_imss,direccion,cp,contrato_sat) " + _
                   "values ('" & numero & "','" & nombre & "','" & apellido_paterno & "','" & apellido_materno & "','" & clave_puesto & "','" & clave_depto & "','" & clave_frecuencia_pago & "','" & num_reg_patronal & "','" & forma_pago & "','" & contrato & "','" & jornada & "','" & regimen_fiscal & "','" & fecha_ingreso & "','" & estatus & "','" & tipo_salario & "','" & salario_diario & "','" & salario_integrado & "','" & rfc & "','" & curp & "','" & registro_imss & "','" & direccion & "','" & cp & "','" & contratosat & "')")




        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                End With
                DBCon.Open()

                comm.ExecuteNonQuery()

            End Using

        Catch ex As SqlException

            MsgBox("ERROR EN LA CONEXION")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try
    End Sub

    Private Sub btnibajas_Click(sender As System.Object, e As System.EventArgs) Handles btnibajas.Click
        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        Try
            Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
            Dim cnConex As New OleDbConnection(stConexion)
            Dim Cmd As New OleDbCommand("Select [numero],[Registro_patronal],[Tipo],[Fecha],[Causa_baja] From [Hoja1$]")
            Dim Ds As New DataSet
            Dim Da As New OleDbDataAdapter
            Dim Dt As New DataTable
            cnConex.Open()
            Cmd.Connection = cnConex
            Da.SelectCommand = Cmd
            Da.Fill(Ds)
            Dt = Ds.Tables(0)
            Me.DataGridView2.Columns.Clear()
            Me.DataGridView2.DataSource = Dt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnabajas_Click(sender As System.Object, e As System.EventArgs) Handles btnabajas.Click
        Dim contador As Integer = 0
        Dim nempleado As String
        Dim registro As String
        Try

            For i As Integer = 0 To Me.DataGridView2.Rows.Count - 1
                With Me.DataGridView2.Rows(i)

                    nempleado = Me.buscarempleado(.Cells(0).Value)
                    If nempleado <> "" Then
                        registro = Me.buscarpatronal(.Cells(1).Value)
                        If registro <> "" Then
                            Me.agbaja(nempleado, registro, .Cells(2).Value, .Cells(3).Value, .Cells(4).Value)
                            contador = contador + 1
                        End If
                    End If
                End With

            Next
            MsgBox("El total de usuarios actualizados fueron " & contador)
        Catch ex As Exception
            MsgBox("Error No Controlado 351: " & ex.Message)
        End Try
    End Sub


    Public Sub Agregarclave(ByVal empleadoid As Integer, ByVal clavesegunda As String, ByVal rol As Integer)


        Dim DBCon As New SqlConnection("data source= 192.168.2.82 ; Initial Catalog= microsip; user id=sicossadmi;password=ipp2012; ")
        ' Dim DBCon As New SqlConnection("data source= 192.168.2.83; Initial Catalog= microsip;user id=sicossadmi;password=ipp2012;")



        Dim consulta As String
        consulta = ("insert into usuario (empleado_id,clave_empleado_id,clave_empleado,rol_clave_emp_id) " + _
                   "values ('" & empleadoid & "','" & empleadoid + 1 & "','" & clavesegunda & "','" & rol & "')")




        Try
            'Abrimos la conexión y comprobamos que no hay error

            Using comm As New SqlCommand(consulta, DBCon)
                With comm
                    .CommandType = CommandType.Text

                End With
                DBCon.Open()

                comm.ExecuteNonQuery()

            End Using

        Catch ex As SqlException

            MsgBox("terminado")
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try
    End Sub


    Public Function buscardatos(ByVal numero As String) As String
        Dim DBCon As OdbcConnection
        Dim cadenaODBC As String

        '      cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        '   ";PWD=8244Ata;DBNAME=192.168.2.21" & _
        '":C:\microsip datos\PRUEBA NOMINAS.FDB"




        cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
       ";PWD=ata8244;DBNAME=192.168.2.83" & _
    ":C:\microsip datos\NEXTEL.FDB"

        DBCon = New OdbcConnection(cadenaODBC)

        Dim consulta As String
        Dim resultado As String
        consulta = "select empleado_id  from empleados " & _
           "where numero  = '" & numero & "' "



        Try
            'Abrimos la conexión y comprobamos que no hay error
            Using comm As New OdbcCommand(consulta, DBCon)
                With comm

                    .CommandType = CommandType.Text

                    '.Parameters.Add(Bempleado)
                End With

                DBCon.Open()
                resultado = comm.ExecuteScalar()

            End Using
            ' MsgBox("Conexion realizada satsfactoriamente")
            If resultado Is Nothing Then
                resultado = ""
            End If


            Return resultado.ToLower
        Catch ex As Odbc.OdbcException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox(ex.Message.ToString)
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try
    End Function

    Public Function buscarpatronal(ByVal registro As String) As String
        Dim DBCon As OdbcConnection
        Dim cadenaODBC As String

        If ComboSUELDO.Text = "NUBULA" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\NUBULA SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "INFORMATION THECNOLOGY INDUSTRIES" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\INFORMATION THECNOLOGY.FDB"
        End If

        If ComboSUELDO.Text = "FOLDUR" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\NEXTEL.FDB"
        End If

        If ComboSUELDO.Text = "MORGET" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\MORGET.FDB"
        End If

        If ComboSUELDO.Text = "GRUPO CONISAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\GRUPO CONISAL.FDB"
        End If


        If ComboSUELDO.Text = "WIPSI" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\WIPSI A C.FDB"
        End If


        If ComboSUELDO.Text = "MORGET SEMANAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
      ";PWD=ata8244;DBNAME=192.168.2.83" & _
   ":C:\microsip datos\1 MORGET SEMANAL.FDB"
        End If
        If ComboSUELDO.Text = "MORGET CATORCENAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
      ";PWD=ata8244;DBNAME= 192.168.2.83" & _
   ":C:\microsip datos\2 MORGET CATORCENAL.FDB"
        End If
        If ComboSUELDO.Text = "MORGET QUINCENAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
      ";PWD=ata8244;DBNAME=192.168.2.83" & _
   ":C:\microsip datos\3  MORGET QUINCENAL.FDB"
        End If
        If ComboSUELDO.Text = "MORGET MENSUAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\4 MORGET MENSUAL.FDB"
        End If

        ''agosto

        If ComboSUELDO.Text = "IT TELECOM" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\IT RESOURCES TELECOM SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "MORGET INTERNA" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\5 MORGET INTERNA.FDB"
        End If

        If ComboSUELDO.Text = "AICEL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\AICEL.FDB"
        End If

        If ComboSUELDO.Text = "CONSORCIO ATERAP SA DE CV" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\CONSORCIO ATERAP SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "CROTEC SA DE CV" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\CROTEC SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "PEPSAT SA DE CV" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\PEPSAT SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "UPHETILOLI 2" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\UPHETILOLI 2.FDB"
        End If




        DBCon = New OdbcConnection(cadenaODBC)

        Dim consulta As String
        Dim resultado As String
        consulta = "select reg_patronal_id from reg_patronales " & _
           "where num_reg_patronal  = '" & registro & "'"



        Try
            'Abrimos la conexión y comprobamos que no hay error
            Using comm As New OdbcCommand(consulta, DBCon)
                With comm

                    .CommandType = CommandType.Text

                    '.Parameters.Add(Bempleado)
                End With

                DBCon.Open()
                resultado = comm.ExecuteScalar()

            End Using
            ' MsgBox("Conexion realizada satsfactoriamente")
            If resultado Is Nothing Then
                resultado = ""
            End If


            Return resultado.ToLower
        Catch ex As Odbc.OdbcException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox(ex.Message.ToString)
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Function

    Public Function buscarempleado(ByVal empleado As String) As String
        Dim DBCon As OdbcConnection
        Dim cadenaODBC As String



        If ComboSUELDO.Text = "NUBULA" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\NUBULA SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "INFORMATION THECNOLOGY INDUSTRIES" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\INFORMATION THECNOLOGY.FDB"
        End If

        If ComboSUELDO.Text = "FOLDUR" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\NEXTEL.FDB"
        End If

        If ComboSUELDO.Text = "MORGET" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\MORGET.FDB"
        End If

        If ComboSUELDO.Text = "GRUPO CONISAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\GRUPO CONISAL.FDB"
        End If


        If ComboSUELDO.Text = "WIPSI" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\WIPSI A C.FDB"
        End If


        If ComboSUELDO.Text = "MORGET SEMANAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
      ";PWD=ata8244;DBNAME=192.168.2.83" & _
   ":C:\microsip datos\1 MORGET SEMANAL.FDB"
        End If
        If ComboSUELDO.Text = "MORGET CATORCENAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
      ";PWD=ata8244;DBNAME= 192.168.2.83" & _
   ":C:\microsip datos\2 MORGET CATORCENAL.FDB"
        End If
        If ComboSUELDO.Text = "MORGET QUINCENAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
      ";PWD=ata8244;DBNAME=192.168.2.83" & _
   ":C:\microsip datos\3  MORGET QUINCENAL.FDB"
        End If
        If ComboSUELDO.Text = "MORGET MENSUAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\4 MORGET MENSUAL.FDB"
        End If

        ''agosto

        If ComboSUELDO.Text = "IT TELECOM" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\IT RESOURCES TELECOM SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "MORGET INTERNA" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\5 MORGET INTERNA.FDB"
        End If

        If ComboSUELDO.Text = "AICEL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\AICEL.FDB"
        End If

        If ComboSUELDO.Text = "CONSORCIO ATERAP SA DE CV" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\CONSORCIO ATERAP SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "CROTEC SA DE CV" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\CROTEC SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "PEPSAT SA DE CV" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\PEPSAT SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "UPHETILOLI 2" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\UPHETILOLI 2.FDB"
        End If


        DBCon = New OdbcConnection(cadenaODBC)

        Dim consulta As String
        Dim resultado As String
        consulta = "select empleado_id  from claves_empleados " & _
           "where clave_empleado  = '" & empleado & "'"



        Try
            'Abrimos la conexión y comprobamos que no hay error
            Using comm As New OdbcCommand(consulta, DBCon)
                With comm

                    .CommandType = CommandType.Text

                    '.Parameters.Add(Bempleado)
                End With

                DBCon.Open()
                resultado = comm.ExecuteScalar()

            End Using
            ' MsgBox("Conexion realizada satsfactoriamente")
            If resultado Is Nothing Then
                resultado = ""
            End If


            Return resultado.ToLower
        Catch ex As Odbc.OdbcException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox(ex.Message.ToString)
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Function



    Public Function agbaja(ByVal id_empleado As Int32, ByVal registo_patronal As String, ByVal tipo As String, ByVal fecha As String, ByVal causa_baja As String)
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction
        Try
            Dim cadenaODBC As String

            If Combopuesto.Text = "NUBULA" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\NUBULA SA DE CV.FDB"
            End If

            If Combopuesto.Text = "INFORMATION THECNOLOGY INDUSTRIES" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\INFORMATION THECNOLOGY.FDB"
            End If



            If ComboSUELDO.Text = "FOLDUR" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\NEXTEL.FDB"
            End If

            If ComboSUELDO.Text = "MORGET" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\MORGET.FDB"
            End If

            If ComboSUELDO.Text = "GRUPO CONISAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\GRUPO CONISAL.FDB"
            End If


            If ComboSUELDO.Text = "WIPSI" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\WIPSI A C.FDB"
            End If




            '    cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
            '   ";PWD=ata8244;DBNAME=192.168.2.83" & _
            '":C:\microsip datos\NEXTEL.FDB"



            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)


                .Append("insert into incidencias (INCIDENCIA_ID,EMPLEADO_ID,REG_PATRONAL_ID,TIPO,FECHA,CAUSA_BAJA,SALINT_DEFAULT,FORMA_EMITIDA)" + _
                        "values (GEN_ID(ID_DOCTOS,1),'" & id_empleado & "','" & registo_patronal & "','" & tipo & "','" & fecha & "','" & causa_baja & "','" & "S" & "', '" & "N" & "')")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    Private Sub btnisueldo_Click(sender As System.Object, e As System.EventArgs) Handles btnisueldo.Click
        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        Try
            Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
            Dim cnConex As New OleDbConnection(stConexion)
            Dim Cmd As New OleDbCommand("Select [numero],[Registro_patronal],[Tipo],[Fecha],[Salario_hora],[Salario_diario],[Salario_integrado]From [Hoja1$]")
            Dim Ds As New DataSet
            Dim Da As New OleDbDataAdapter
            Dim Dt As New DataTable
            cnConex.Open()
            Cmd.Connection = cnConex
            Da.SelectCommand = Cmd
            Da.Fill(Ds)
            Dt = Ds.Tables(0)
            Me.DataGridView3.Columns.Clear()
            Me.DataGridView3.DataSource = Dt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Muestradatos()


        Dim c As New datoshandler
        arreDatos = New ArrayList

        arreDatos = c.regresadatos()
        Try
            For i As Integer = 0 To Me.arreDatos.Count - 1

                With CType(Me.arreDatos(i), datos)

                    'If .mensaje3 <> "" Then

                    Me.DataSet11.datos.AdddatosRow(.mensaje)

                    'End If
                    Me.DataGridView4.DataSource = DataSet11.Tables(0).DefaultView
                End With

            Next

        Catch ex As Exception
            MsgBox("error 2")
        End Try

        c.gridatxt(DataGridView4)
    End Sub




    Private Sub btnasueldo_Click(sender As System.Object, e As System.EventArgs) Handles btnasueldo.Click
        Dim contador As Integer = 0
        Dim nempleado As String
        Dim registro As String
        Try

            For i As Integer = 0 To Me.DataGridView3.Rows.Count - 1
                With Me.DataGridView3.Rows(i)

                    nempleado = Me.buscarempleado(.Cells(0).Value)
                    If nempleado <> "" Then
                        registro = Me.buscarpatronal(.Cells(1).Value)
                        If registro <> "" Then
                            Me.agsueldos(nempleado, registro, .Cells(2).Value, .Cells(3).Value, .Cells(4).Value, .Cells(5).Value, .Cells(6).Value)
                            contador = contador + 1
                        End If
                    End If
                End With
            Next
            MsgBox("El total de usuarios actualizados fueron " & contador)
        Catch ex As Exception
            MsgBox("Error No Controlado 351: " & ex.Message)
        End Try
    End Sub


    Public Function agsueldos(ByVal id_empleado As Int32, ByVal registo_patronal As String, ByVal tipo As String, ByVal fecha As String, ByVal salario_hora As Double, ByVal salario_diario As Double, ByVal salario_integrado As Double)
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction
        Try
            Dim cadenaODBC As String

            If ComboSUELDO.Text = "NUBULA" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\NUBULA SA DE CV.FDB"
            End If

            If ComboSUELDO.Text = "INFORMATION THECNOLOGY INDUSTRIES" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\INFORMATION THECNOLOGY.FDB"
            End If

            If ComboSUELDO.Text = "FOLDUR" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\NEXTEL.FDB"
            End If

            If ComboSUELDO.Text = "MORGET" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\MORGET.FDB"
            End If

            If ComboSUELDO.Text = "GRUPO CONISAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\GRUPO CONISAL.FDB"
            End If


            If ComboSUELDO.Text = "WIPSI" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\WIPSI A C.FDB"
            End If


            If ComboSUELDO.Text = "MORGET SEMANAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME=192.168.2.83" & _
       ":C:\microsip datos\1 MORGET SEMANAL.FDB"
            End If
            If ComboSUELDO.Text = "MORGET CATORCENAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME= 192.168.2.83" & _
       ":C:\microsip datos\2 MORGET CATORCENAL.FDB"
            End If
            If ComboSUELDO.Text = "MORGET QUINCENAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME=192.168.2.83" & _
       ":C:\microsip datos\3  MORGET QUINCENAL.FDB"
            End If
            If ComboSUELDO.Text = "MORGET MENSUAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\4 MORGET MENSUAL.FDB"
            End If

            ''agosto

            If ComboSUELDO.Text = "IT TELECOM" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\IT RESOURCES TELECOM SA DE CV.FDB"
            End If

            If ComboSUELDO.Text = "MORGET INTERNA" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\5 MORGET INTERNA.FDB"
            End If

            If ComboSUELDO.Text = "AICEL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\AICEL.FDB"
            End If

            If ComboSUELDO.Text = "CONSORCIO ATERAP SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\CONSORCIO ATERAP SA DE CV.FDB"
            End If

            If ComboSUELDO.Text = "CROTEC SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\CROTEC SA DE CV.FDB"
            End If

            If ComboSUELDO.Text = "PEPSAT SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\PEPSAT SA DE CV.FDB"
            End If

            If ComboSUELDO.Text = "UPHETILOLI 2" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\UPHETILOLI 2.FDB"
            End If


            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)


                .Append("insert into incidencias (INCIDENCIA_ID,EMPLEADO_ID,REG_PATRONAL_ID,TIPO,FECHA,SALARIO_DIARIO,SALARIO_HORA,SALARIO_INTEG,SALINT_DEFAULT,FORMA_EMITIDA)" + _
                        "values (GEN_ID(ID_DOCTOS,1),'" & id_empleado & "','" & registo_patronal & "','" & tipo & "','" & fecha & "','" & salario_diario & "','" & salario_hora & "','" & salario_integrado & "','" & "S" & "', '" & "N" & "')")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    ''depto
    Public Function agdepto(ByVal nombre As String)
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction
        Try
            Dim cadenaODBC As String

            If Combodepto.Text = "NUBULA" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\NUBULA SA DE CV.FDB"
            End If

            If Combodepto.Text = "INFORMATION THECNOLOGY INDUSTRIES" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\INFORMATION THECNOLOGY.FDB"
            End If

            If Combodepto.Text = "FOLDUR" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\NEXTEL.FDB"
            End If

            If Combodepto.Text = "MORGET" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\MORGET.FDB"
            End If

            If Combodepto.Text = "GRUPO CONISAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\GRUPO CONISAL.FDB"
            End If


            If Combodepto.Text = "WIPSI" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\WIPSI A C.FDB"
            End If


            If Combodepto.Text = "MORGET SEMANAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME=192.168.2.83" & _
       ":C:\microsip datos\1 MORGET SEMANAL.FDB"
            End If
            If Combodepto.Text = "MORGET CATORCENAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME= 192.168.2.83" & _
       ":C:\microsip datos\2 MORGET CATORCENAL.FDB"
            End If
            If Combodepto.Text = "MORGET QUINCENAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME=192.168.2.83" & _
       ":C:\microsip datos\3  MORGET QUINCENAL.FDB"
            End If
            If Combodepto.Text = "MORGET MENSUAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\4 MORGET MENSUAL.FDB"
            End If

            ''agosto

            If Combodepto.Text = "IT TELECOM" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\IT RESOURCES TELECOM SA DE CV.FDB"
            End If

            If Combodepto.Text = "MORGET INTERNA" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\5 MORGET INTERNA.FDB"
            End If

            If Combodepto.Text = "AICEL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\AICEL.FDB"
            End If

            If Combodepto.Text = "CONSORCIO ATERAP SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\CONSORCIO ATERAP SA DE CV.FDB"
            End If

            If Combodepto.Text = "CROTEC SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\CROTEC SA DE CV.FDB"
            End If

            If Combodepto.Text = "PEPSAT SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\PEPSAT SA DE CV.FDB"
            End If

            If Combodepto.Text = "UPHETILOLI 2" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\UPHETILOLI 2.FDB"
            End If


            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)





                .Append("insert into DEPTOS_NO (DEPTO_NO_ID,NOMBRE)" + _
                      "values (GEN_ID(ID_CATALOGOS,1),'" & nombre & "')")

            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    ''puesto
    Public Function agpuesto(ByVal nombre As String, ByVal diario As String, ByVal diariom As String)
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction
        Try
            Dim cadenaODBC As String


            If Combopuesto.Text = "NUBULA" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\NUBULA SA DE CV.FDB"
            End If

            If Combopuesto.Text = "INFORMATION THECNOLOGY INDUSTRIES" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\INFORMATION THECNOLOGY.FDB"
            End If

            If Combopuesto.Text = "FOLDUR" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\NEXTEL.FDB"
            End If

            If Combopuesto.Text = "MORGET" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\MORGET.FDB"
            End If

            If Combopuesto.Text = "GRUPO CONISAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\GRUPO CONISAL.FDB"
            End If


            If Combopuesto.Text = "WIPSI" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\WIPSI A C.FDB"
            End If


            If Combopuesto.Text = "MORGET SEMANAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME=192.168.2.83" & _
       ":C:\microsip datos\1 MORGET SEMANAL.FDB"
            End If
            If Combopuesto.Text = "MORGET CATORCENAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME= 192.168.2.83" & _
       ":C:\microsip datos\2 MORGET CATORCENAL.FDB"
            End If
            If Combopuesto.Text = "MORGET QUINCENAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME=192.168.2.83" & _
       ":C:\microsip datos\3  MORGET QUINCENAL.FDB"
            End If
            If Combopuesto.Text = "MORGET MENSUAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\4 MORGET MENSUAL.FDB"
            End If

            ''agosto

            If Combopuesto.Text = "IT TELECOM" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\IT RESOURCES TELECOM SA DE CV.FDB"
            End If

            If Combopuesto.Text = "MORGET INTERNA" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\5 MORGET INTERNA.FDB"
            End If

            If Combopuesto.Text = "AICEL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\AICEL.FDB"
            End If

            If Combopuesto.Text = "CONSORCIO ATERAP SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\CONSORCIO ATERAP SA DE CV.FDB"
            End If

            If Combopuesto.Text = "CROTEC SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\CROTEC SA DE CV.FDB"
            End If

            If Combopuesto.Text = "PEPSAT SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\PEPSAT SA DE CV.FDB"
            End If

            If Combopuesto.Text = "UPHETILOLI 2" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\UPHETILOLI 2.FDB"
            End If


            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)


                .Append("insert into PUESTOS_NO (PUESTO_NO_ID,NOMBRE,SUELDO_DIARIO,SUELDO_DIARIO_MAX)" + _
                        "values (GEN_ID(ID_CATALOGOS,1),'" & nombre & "','" & diario & "','" & diariom & "')")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    Public Function agdeptoc(ByVal tipo_depto As String, ByVal clave As String)
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction
        Try
            Dim cadenaODBC As String

            ''bases por empresa

            If Combodepto.Text = "NUBULA" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\NUBULA SA DE CV.FDB"
            End If

            If Combodepto.Text = "INFORMATION THECNOLOGY INDUSTRIES" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\INFORMATION THECNOLOGY.FDB"
            End If

            If Combodepto.Text = "FOLDUR" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\NEXTEL.FDB"
            End If

            If Combodepto.Text = "MORGET" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\MORGET.FDB"
            End If

            If Combodepto.Text = "GRUPO CONISAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\GRUPO CONISAL.FDB"
            End If


            If Combodepto.Text = "WIPSI" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\WIPSI A C.FDB"
            End If


            If Combodepto.Text = "MORGET SEMANAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME=192.168.2.83" & _
       ":C:\microsip datos\1 MORGET SEMANAL.FDB"
            End If
            If Combodepto.Text = "MORGET CATORCENAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME= 192.168.2.83" & _
       ":C:\microsip datos\2 MORGET CATORCENAL.FDB"
            End If
            If Combodepto.Text = "MORGET QUINCENAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME=192.168.2.83" & _
       ":C:\microsip datos\3  MORGET QUINCENAL.FDB"
            End If
            If Combodepto.Text = "MORGET MENSUAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\4 MORGET MENSUAL.FDB"
            End If

            ''agosto

            If Combodepto.Text = "IT TELECOM" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\IT RESOURCES TELECOM SA DE CV.FDB"
            End If

            If Combodepto.Text = "MORGET INTERNA" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\5 MORGET INTERNA.FDB"
            End If

            If Combodepto.Text = "AICEL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\AICEL.FDB"
            End If

            If Combodepto.Text = "CONSORCIO ATERAP SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\CONSORCIO ATERAP SA DE CV.FDB"
            End If

            If Combodepto.Text = "CROTEC SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\CROTEC SA DE CV.FDB"
            End If

            If Combodepto.Text = "PEPSAT SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\PEPSAT SA DE CV.FDB"
            End If

            If Combodepto.Text = "UPHETILOLI 2" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\UPHETILOLI 2.FDB"
            End If

            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)


                .Append("insert into CLAVES_CAT_SEC (NOMBRE_TABLA,ELEM_ID,CLAVE)" + _
                        "values ('" & "DEPTOS_NO" & "','" & tipo_depto & "','" & clave & "')")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    Public Function agpuestoc(ByVal tipo_depto As String, ByVal clave As String)
        'Dim tr As FbTransaction
        Dim trODBC As OdbcTransaction
        Try
            Dim cadenaODBC As String


            If Combopuesto.Text = "NUBULA" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\NUBULA SA DE CV.FDB"
            End If

            If Combopuesto.Text = "INFORMATION THECNOLOGY INDUSTRIES" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\INFORMATION THECNOLOGY.FDB"
            End If

            If Combopuesto.Text = "FOLDUR" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\NEXTEL.FDB"
            End If

            If Combopuesto.Text = "MORGET" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\MORGET.FDB"
            End If

            If Combopuesto.Text = "GRUPO CONISAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\GRUPO CONISAL.FDB"
            End If


            If Combopuesto.Text = "WIPSI" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\WIPSI A C.FDB"
            End If


            If Combopuesto.Text = "MORGET SEMANAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME=192.168.2.83" & _
       ":C:\microsip datos\1 MORGET SEMANAL.FDB"
            End If
            If Combopuesto.Text = "MORGET CATORCENAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME= 192.168.2.83" & _
       ":C:\microsip datos\2 MORGET CATORCENAL.FDB"
            End If
            If Combopuesto.Text = "MORGET QUINCENAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME=192.168.2.83" & _
       ":C:\microsip datos\3  MORGET QUINCENAL.FDB"
            End If
            If Combopuesto.Text = "MORGET MENSUAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\4 MORGET MENSUAL.FDB"
            End If

            ''agosto

            If Combopuesto.Text = "IT TELECOM" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\IT RESOURCES TELECOM SA DE CV.FDB"
            End If

            If Combopuesto.Text = "MORGET INTERNA" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\5 MORGET INTERNA.FDB"
            End If

            If Combopuesto.Text = "AICEL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\AICEL.FDB"
            End If

            If Combopuesto.Text = "CONSORCIO ATERAP SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\CONSORCIO ATERAP SA DE CV.FDB"
            End If

            If Combopuesto.Text = "CROTEC SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\CROTEC SA DE CV.FDB"
            End If

            If Combopuesto.Text = "PEPSAT SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\PEPSAT SA DE CV.FDB"
            End If

            If Combopuesto.Text = "UPHETILOLI 2" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\UPHETILOLI 2.FDB"
            End If

            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)


                .Append("insert into CLAVES_CAT_SEC (NOMBRE_TABLA,ELEM_ID,CLAVE)" + _
                        "values ('" & "PUESTOS_NO" & "','" & tipo_depto & "','" & clave & "')")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        Try
            Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
            Dim cnConex As New OleDbConnection(stConexion)
            Dim Cmd As New OleDbCommand("Select [nombre],[clave]From [Hoja1$]")
            Dim Ds As New DataSet
            Dim Da As New OleDbDataAdapter
            Dim Dt As New DataTable
            cnConex.Open()
            Cmd.Connection = cnConex
            Da.SelectCommand = Cmd
            Da.Fill(Ds)
            Dt = Ds.Tables(0)
            Me.DataGridView5.Columns.Clear()
            Me.DataGridView5.DataSource = Dt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim contador As Integer = 0
        Dim depto As String
        Dim registro As String
        Try

            For i As Integer = 0 To Me.DataGridView5.Rows.Count - 1
                With Me.DataGridView5.Rows(i)


                    Me.agdepto(.Cells(0).Value)

                    depto = Me.buscardepto(.Cells(0).Value)
                    If depto <> "" Then
                        Me.agdeptoc(depto, .Cells(1).Value)

                    Else
                        MessageBox.Show("No se encontro el departamento")
                    End If
                    contador = contador + 1

                End With
            Next
            MsgBox("El total de usuarios actualizados fueron " & contador)
        Catch ex As Exception
            MsgBox("Error No Controlado 351: " & ex.Message)
        End Try
    End Sub

    Private Sub btnpuestoi_Click(sender As System.Object, e As System.EventArgs) Handles btnpuestoi.Click
        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        Try
            Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
            Dim cnConex As New OleDbConnection(stConexion)
            Dim Cmd As New OleDbCommand("Select [Puesto],[Clave],[Sueldo Diario],[Sueldo Diario Maximo]From [Hoja1$]")
            Dim Ds As New DataSet
            Dim Da As New OleDbDataAdapter
            Dim Dt As New DataTable
            cnConex.Open()
            Cmd.Connection = cnConex
            Da.SelectCommand = Cmd
            Da.Fill(Ds)
            Dt = Ds.Tables(0)
            Me.DataGridView6.Columns.Clear()
            Me.DataGridView6.DataSource = Dt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnpuestoa_Click(sender As System.Object, e As System.EventArgs) Handles btnpuestoa.Click
        Dim contador As Integer = 0
        Dim puesto As String

        Try

            For i As Integer = 0 To Me.DataGridView6.Rows.Count - 1
                With Me.DataGridView6.Rows(i)

                    Me.agpuesto(.Cells(0).Value, .Cells(2).Value, .Cells(3).Value)

                    puesto = Me.buscarpuesto(.Cells(0).Value)
                    If puesto <> "" Then
                        Me.agpuestoc(puesto, .Cells(1).Value)
                    End If
                    contador = contador + 1

                End With
            Next
            MsgBox("El total de usuarios actualizados fueron " & contador)
        Catch ex As Exception
            MsgBox("Error No Controlado 351: " & ex.Message)
        End Try
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Dim stRuta As String = ""
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = "Seleccionar archivos"
            .Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                stRuta = .FileName
            End If
        End With
        Try
            Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (stRuta & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
            Dim cnConex As New OleDbConnection(stConexion)
            Dim Cmd As New OleDbCommand("Select [Empleado],[Cuenta],[Banco]From [Hoja1$]")
            Dim Ds As New DataSet
            Dim Da As New OleDbDataAdapter
            Dim Dt As New DataTable
            cnConex.Open()
            Cmd.Connection = cnConex
            Da.SelectCommand = Cmd
            Da.Fill(Ds)
            Dt = Ds.Tables(0)
            Me.DataGridView7.Columns.Clear()
            Me.DataGridView7.DataSource = Dt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Dim contador As Integer = 0
        Dim nempleado As String
        Dim registro As String
        Try

            For i As Integer = 0 To Me.DataGridView7.Rows.Count - 1
                With Me.DataGridView7.Rows(i)



                    Me.agcuenta(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value)
                    Me.agcuentab(.Cells(0).Value, .Cells(2).Value)
                    contador = contador + 1


                End With
            Next
            MsgBox("El total de usuarios actualizados fueron " & contador)
        Catch ex As Exception
            MsgBox("Error No Controlado 351: " & ex.Message)
        End Try
    End Sub

    Public Function agcuenta(ByVal numero As String, ByVal cuenta As String, ByVal banco As String)
        'Dim tr As FbTransaction


        Dim numerouno As String = (cuenta.Substring(0, 11))

        Dim trODBC As OdbcTransaction




        Try
            Dim cadenaODBC As String

            If Combobancos.Text = "PEUGEOT" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" &
         ";PWD=ata8244;DBNAME=201.139.106.58" &
      ":C:\microsip datos\ITI  PEUGEOT.FDB"
            End If



            If Combobancos.Text = "NUBULA" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\NUBULA SA DE CV.FDB"
            End If

            If Combobancos.Text = "INFORMATION THECNOLOGY INDUSTRIES" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\INFORMATION THECNOLOGY.FDB"
            End If

            If Combobancos.Text = "FOLDUR" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\NEXTEL.FDB"
            End If

            If Combobancos.Text = "MORGET" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\MORGET.FDB"
            End If

            If Combobancos.Text = "GRUPO CONISAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\GRUPO CONISAL.FDB"
            End If


            If Combobancos.Text = "WIPSI" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\WIPSI A C.FDB"
            End If


            If Combobancos.Text = "MORGET SEMANAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME=192.168.2.83" & _
       ":C:\microsip datos\1 MORGET SEMANAL.FDB"
            End If
            If Combobancos.Text = "MORGET CATORCENAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME= 192.168.2.83" & _
       ":C:\microsip datos\2 MORGET CATORCENAL.FDB"
            End If
            If Combobancos.Text = "MORGET QUINCENAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME=192.168.2.83" & _
       ":C:\microsip datos\3  MORGET QUINCENAL.FDB"
            End If
            If Combobancos.Text = "MORGET MENSUAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\4 MORGET MENSUAL.FDB"
            End If

            ''agosto

            If Combobancos.Text = "IT TELECOM" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\IT RESOURCES TELECOM SA DE CV.FDB"
            End If

            If Combobancos.Text = "MORGET INTERNA" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\5 MORGET INTERNA.FDB"
            End If

            If Combobancos.Text = "AICEL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=201.139.106.58" & _
     ":C:\microsip datos\AICEL.FDB"
            End If

            If Combobancos.Text = "CONSORCIO ATERAP SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\CONSORCIO ATERAP SA DE CV.FDB"
            End If

            If Combobancos.Text = "CROTEC SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\CROTEC SA DE CV.FDB"
            End If

            If Combobancos.Text = "PEPSAT SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\PEPSAT SA DE CV.FDB"
            End If

            If Combobancos.Text = "UPHETILOLI 2" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\UPHETILOLI 2.FDB"
            End If


            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("update empleados set NUM_CTABAN_PAGO_ELECT = '" & numerouno & "' where numero = '" & numero & "' ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    ''BANCO
    Public Function agcuentab(ByVal numero As String, ByVal banco As String)

        ''NEXTEL

        If Combobancos.Text = "FOLDUR" Then
            If banco = "SANTANDER" Then
                banco = 13479
            End If

            If banco = "BANCOMER" Then
                banco = 7516
            End If
        End If
        ''NEXTEL

        ''AICEL

        If Combobancos.Text = "AICEL" Then

            If banco = "AZTECA" Then
                banco = 17416
            End If

            If banco = "BANAMEX" Then
                banco = 17417
            End If
            If banco = "BANBAJIO" Then
                banco = 17418
            End If
            If banco = "BANORTE" Then
                banco = 17419
            End If
            If banco = "BBVA" Then
                banco = 17420
            End If
            If banco = "HSBC" Then
                banco = 17421
            End If
            If banco = "INBURSA" Then
                banco = 17422
            End If
            If banco = "SANTANDER" Then
                banco = 17423
            End If
            If banco = "IXE BANCO" Then
                banco = 17510
            End If
            If banco = "SCOTIABANK" Then
                banco = 17511
            End If
            If banco = "CHEQUE" Then
                banco = 19398
            End If
            If banco = "BANCO DEL BAJIO" Then
                banco = 19399
            End If
            If banco = "BANCOPPEL" Then
                banco = 19400
            End If
            If banco = "AUTOFIN" Then
                banco = 19401
            End If
            If banco = "AFIRME" Then
                banco = 19403
            End If
            If banco = "BANREGIO" Then
                banco = 19609
            End If

        End If


        ''AICEL

        'MORGET INTERNA

        If Combobancos.Text = "MORGET INTERNA" Then

            If banco = "BANAMEX" Then
                banco = 796
            End If
            If banco = "HSBC" Then
                banco = 805
            End If
            If banco = "BANCO AZTECA" Then
                banco = 813
            End If
            If banco = "BANORTE" Then
                banco = 821
            End If
            If banco = "BBVA BANCOMER" Then
                banco = 829
            End If
            If banco = "SANTANDER" Then
                banco = 839
            End If
            If banco = "SCOTIABANK" Then
                banco = 847
            End If

        End If
        'MORGET INTERNA


        'MORGET SEMANAL

        If Combobancos.Text = "MORGET SEMANAL" Then

            If banco = "AFIRME" Then
                banco = 5944
            End If
            If banco = "AUTOFIN" Then
                banco = 5945
            End If
            If banco = "CITIBANAMEX" Then
                banco = 5946
            End If
            If banco = "BANBAJIO" Then
                banco = 5947
            End If
            If banco = "BANCO AZTECA" Then
                banco = 5949
            End If
            If banco = "BANCOMER" Then
                banco = 5950
            End If
            If banco = "BANCOPPEL" Then
                banco = 5951
            End If
            If banco = "BANORTE" Then
                banco = 5952
            End If
            If banco = "HSBC" Then
                banco = 5953
            End If
            If banco = "IXE" Then
                banco = 5954
            End If
            If banco = "SANTANDER" Then
                banco = 5955
            End If
            If banco = "SCOTIABANK" Then
                banco = 5956
            End If
            If banco = "BANREGIO" Then
                banco = 5957
            End If

        End If

        'MORGET SEMANAL

        'MORGET CATORCENAL

        If Combobancos.Text = "MORGET CATORCENAL" Then

            If banco = "AFIRME" Then
                banco = 2249
            End If
            If banco = "AUTOFIN" Then
                banco = 2250
            End If
            If banco = "CITIBANAMEX" Then
                banco = 2251
            End If
            If banco = "BANBAJIO" Then
                banco = 2252
            End If
            If banco = "BANCO AZTECA" Then
                banco = 2253
            End If
            If banco = "BANCOMER" Then
                banco = 2255
            End If
            If banco = "BANCOPPEL" Then
                banco = 2254
            End If
            If banco = "BANORTE" Then
                banco = 2256
            End If
            If banco = "HSBC" Then
                banco = 2257
            End If
            If banco = "IXE" Then
                banco = 2258
            End If
            If banco = "SANTANDER" Then
                banco = 2259
            End If
            If banco = "SCOTIABANK" Then
                banco = 2260
            End If
            If banco = "BANREGIO" Then
                banco = 2261
            End If

        End If


        'MORGET CATORCENAL


        'MORGET QUINCENAL

        If Combobancos.Text = "MORGET QUINCENAL" Then

            If banco = "AFIRME" Then
                banco = 3828
            End If
            If banco = "AUTOFIN" Then
                banco = 3829
            End If
            If banco = "CITIBANAMEX" Then
                banco = 3830
            End If
            If banco = "BANBAJIO" Then
                banco = 3831
            End If
            If banco = "BANCO AZTECA" Then
                banco = 3832
            End If
            If banco = "BANCOMER" Then
                banco = 3833
            End If
            If banco = "BANCOPPEL" Then
                banco = 3834

            End If
            If banco = "BANORTE" Then
                banco = 3538


            End If
            If banco = "HSBC" Then
                banco = 3836

            End If
            If banco = "IXE" Then
                banco = 3837

            End If
            If banco = "SANTANDER" Then
                banco = 3838

            End If
            If banco = "SCOTIABANK" Then
                banco = 3839

            End If
            If banco = "BANREGIO" Then
                banco = 3840


            End If
        End If

        'MORGET QUINCENAL


        'MORGET MENSUAL

        If Combobancos.Text = "MORGET MENSUAL" Then

            If banco = "AFIRME" Then
                banco = 2646
            End If
            If banco = "AUTOFIN" Then
                banco = 2647
            End If
            If banco = "CITIBANAMEX" Then
                banco = 2648
            End If
            If banco = "BANBAJIO" Then
                banco = 2649
            End If
            If banco = "BANCO AZTECA" Then
                banco = 2650
            End If
            If banco = "BANCOMER" Then
                banco = 2651
            End If
            If banco = "BANCOPPEL" Then
                banco = 2652

            End If
            If banco = "BANORTE" Then
                banco = 2653


            End If
            If banco = "HSBC" Then
                banco = 2654

            End If
            If banco = "IXE" Then
                banco = 2655

            End If
            If banco = "SANTANDER" Then
                banco = 2656

            End If
            If banco = "SCOTIABANK" Then
                banco = 2657

            End If
            If banco = "BANREGIO" Then
                banco = 2658


            End If
        End If

        'MORGET MENSUAL


        If Combobancos.Text = "MORGET" Then

            If banco = "BANCOMER" Then
                banco = 6942
            End If
            If banco = "AFIRME" Then
                banco = 7012
            End If
            If banco = "AUTOFIN" Then
                banco = 7021
            End If
            If banco = "BANAMEX" Then
                banco = 7029
            End If
            If banco = "BANBAJIO" Then
                banco = 7030
            End If
            If banco = "BANCO AZTECA" Then
                banco = 7031
            End If
            If banco = "BANCO DEL BAJIO" Then
                banco = 7032
            End If
            If banco = "BANCOPPEL" Then
                banco = 7033
            End If
            If banco = "BANORTE" Then
                banco = 7034
            End If
            If banco = "CHEQUE" Then
                banco = 7035
            End If
            If banco = "HSBC" Then
                banco = 7036
            End If
            If banco = "IXE" Then
                banco = 7037
            End If
            If banco = "SANTANDER" Then
                banco = 7038
            End If
            If banco = "SCOTIABANK" Then
                banco = 7039
            End If
        End If



        If Combobancos.Text = "GRUPO CONISAL" Then

            If banco = "BBVA" Then
                banco = 416
            End If
            If banco = "BANAMEX" Then
                banco = 426
            End If
            If banco = "INBURSA" Then
                banco = 435
            End If
            If banco = "BANBAJIO" Then
                banco = 444
            End If
            If banco = "AZTECA" Then
                banco = 8568
            End If
            If banco = "BANORTE" Then
                banco = 8569
            End If
            If banco = "HSBC" Then
                banco = 8570
            End If
            If banco = "SANTANDER" Then
                banco = 8571
            End If

        End If

        ''wipsi
        If Combobancos.Text = "WIPSI" Then


            If banco = "BBVA" Then
                banco = 17420
            End If
            If banco = "BANAMEX" Then
                banco = 17417
            End If
            If banco = "INBURSA" Then
                banco = 17422
            End If
            If banco = "BANBAJIO" Then
                banco = 17418
            End If
            If banco = "AZTECA" Then
                banco = 17416
            End If
            If banco = "BANORTE" Then
                banco = 17419
            End If
            If banco = "HSBC" Then
                banco = 17421
            End If
            If banco = "SANTANDER" Then
                banco = 17423
            End If
            If banco = "IXE BANCO" Then
                banco = 17510
            End If
            If banco = "SCOTIABANK" Then
                banco = 17511
            End If

            If banco = "CHEQUE" Then
                banco = 19398
            End If

            If banco = "BANCO DEL BAJIO" Then
                banco = 19399
            End If

            If banco = "BANCOPPEL" Then
                banco = 19400
            End If

            If banco = "AUTOFIN" Then
                banco = 19401
            End If
            If banco = "AFIRME" Then
                banco = 19403
            End If

        End If


        'IT TELECOM

        If Combobancos.Text = "IT TELECOM" Then

            If banco = "BBVA" Then
                banco = 9175
            End If
            If banco = "BANAMEX" Then
                banco = 9172
            End If
            If banco = "INBURSA" Then
                banco = 9177
            End If
            If banco = "BANBAJIO" Then
                banco = 9173
            End If
            If banco = "IXE BANCO" Then
                banco = 9178
            End If
            If banco = "BANORTE" Then
                banco = 9174
            End If
            If banco = "HSBC" Then
                banco = 9176
            End If
            If banco = "SANTANDER" Then
                banco = 9179
            End If
            If banco = "SCOTIABANK" Then
                banco = 9180
            End If

        End If


        ''UPHETILOLI

        If Combobancos.Text = "UPHETILOLI 2" Then

            If banco = "BANORTE" Then
                banco = 373
            End If

            If banco = "AFIRME" Then
                banco = 382
            End If

            If banco = "AUTOFIN" Then
                banco = 390
            End If

            If banco = "BANBAJIO" Then
                banco = 398
            End If

            If banco = "BANCO AZTECA" Then
                banco = 406
            End If

            If banco = "BANCOMER" Then
                banco = 414
            End If

            If banco = "BANCOPPEL" Then
                banco = 422
            End If

            If banco = "BANREGIO" Then
                banco = 430
            End If

            If banco = "CITIBANAMEX" Then
                banco = 438
            End If

            If banco = "HSBC" Then
                banco = 446
            End If

            If banco = "IXE" Then
                banco = 454
            End If

            If banco = "SANTANDER" Then
                banco = 462
            End If

            If banco = "SCOTIABANK" Then
                banco = 470
            End If

        End If

        'UPHETILOLI

        If Combobancos.Text = "NUBULA" Then

            If banco = "BANCOMER" Then
                banco = 502
            End If

            If banco = "SANTANDER" Then
                banco = 505
            End If

            If banco = "BANAMEX" Then
                banco = 508
            End If

            If banco = "HSBC" Then
                banco = 520
            End If

        End If

        If Combobancos.Text = "INFORMATION THECNOLOGY" Then

            If banco = "BANCOMER" Then
                banco = 500
            End If
            If banco = "BANCO AZTECA" Then
                banco = 503
            End If
            If banco = "HSBC" Then
                banco = 514
            End If
            If banco = "SANTANDER" Then
                banco = 521
            End If

        End If

        If Combobancos.Text = "PEUGEOT" Then

            If banco = "BANORTE" Then
                banco = 671
            End If
            If banco = "BANCOMER" Then
                banco = 680
            End If
            If banco = "BANAMEX" Then
                banco = 688
            End If
            If banco = "SCOTIABANK" Then
                banco = 696
            End If
            If banco = "SANTANDER" Then
                banco = 704
            End If
            If banco = "HSBC" Then
                banco = 712
            End If


        End If


        Dim trODBC As OdbcTransaction
        Try
            Dim cadenaODBC As String


            If Combobancos.Text = "PEUGEOT" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" &
         ";PWD=ata8244;DBNAME=201.139.106.58" &
      ":C:\microsip datos\ITI  PEUGEOT.FDB"
            End If


            If Combobancos.Text = "NUBULA" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\NUBULA SA DE CV.FDB"
            End If

            If Combobancos.Text = "INFORMATION THECNOLOGY INDUSTRIES" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\INFORMATION THECNOLOGY.FDB"
            End If

            If Combobancos.Text = "FOLDUR" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\NEXTEL.FDB"
            End If

            If Combobancos.Text = "MORGET" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\MORGET.FDB"
            End If

            If Combobancos.Text = "GRUPO CONISAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\GRUPO CONISAL.FDB"
            End If


            If Combobancos.Text = "WIPSI" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
         ";PWD=ata8244;DBNAME=192.168.2.83" & _
      ":C:\microsip datos\WIPSI A C.FDB"
            End If


            If Combobancos.Text = "MORGET SEMANAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME=192.168.2.83" & _
       ":C:\microsip datos\1 MORGET SEMANAL.FDB"
            End If
            If Combobancos.Text = "MORGET CATORCENAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME= 192.168.2.83" & _
       ":C:\microsip datos\2 MORGET CATORCENAL.FDB"
            End If
            If Combobancos.Text = "MORGET QUINCENAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
          ";PWD=ata8244;DBNAME=192.168.2.83" & _
       ":C:\microsip datos\3  MORGET QUINCENAL.FDB"
            End If
            If Combobancos.Text = "MORGET MENSUAL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\4 MORGET MENSUAL.FDB"
            End If

            ''agosto

            If Combobancos.Text = "IT TELECOM" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\IT RESOURCES TELECOM SA DE CV.FDB"
            End If

            If Combobancos.Text = "MORGET INTERNA" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=201.139.106.58" & _
     ":C:\microsip datos\5 MORGET INTERNA.FDB"
            End If

            If Combobancos.Text = "AICEL" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\AICEL.FDB"
            End If

            If Combobancos.Text = "CONSORCIO ATERAP SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\CONSORCIO ATERAP SA DE CV.FDB"
            End If

            If Combobancos.Text = "CROTEC SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\CROTEC SA DE CV.FDB"
            End If

            If Combobancos.Text = "PEPSAT SA DE CV" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\PEPSAT SA DE CV.FDB"
            End If

            If Combobancos.Text = "UPHETILOLI 2" Then
                cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
        ";PWD=ata8244;DBNAME=192.168.2.83" & _
     ":C:\microsip datos\UPHETILOLI 2.FDB"
            End If

            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()

            Me.m_ConnODBC = New OdbcConnection(cadenaODBC)
            Me.m_ConnODBC.Open()
            trODBC = Me.m_ConnODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New OdbcCommand("", Me.m_ConnODBC, trODBC)
            With strQuery
                .Remove(0, .Length)



                .Append("update empleados set GRUPO_PAGO_ELECT_ID = '" & banco & "' where numero = '" & numero & "' ")


            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            commODBC.ExecuteNonQuery()
            trODBC.Commit()
            'Me.m_conn.Close()
            Me.m_ConnODBC.Close()

        Catch ex As Exception
            Try

                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox(ex.Message)
            End Try
        Finally
            Try
                Me.m_ConnODBC.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Try
    End Function

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Me.delete()
        Try

            For i As Integer = 0 To Me.DataGridView1.Rows.Count - 1
                With Me.DataGridView1.Rows(i)

                    If .Cells(0) Is DBNull.Value Then
                        MsgBox("No hay mas datos")
                    End If

                    Me.Agregausuario(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value, .Cells(3).Value, .Cells(4).Value, .Cells(5).Value, .Cells(6).Value, .Cells(7).Value, .Cells(8).Value, .Cells(9).Value, .Cells(10).Value, .Cells(11).Value, .Cells(12).Value, .Cells(13).Value, .Cells(14).Value, .Cells(15).Value, .Cells(16).Value, .Cells(17).Value, .Cells(18).Value, .Cells(19).Value, .Cells(20).Value, .Cells(21).Value, .Cells(22).Value)
                    'empleadoid = Me.buscardatos(.Cells(0).Value)
                    'rol = Me.buscarrol(.Cells())
                    'If empleadoid <> "" Then

                    '    Me.Agregarclave(empleadoid, ("I" + .Cells(0).Value), rol)
                    'End If
                End With
            Next

            MsgBox("usuarios almacenados correctamente")


        Catch ex As Exception
            MsgBox("Error numero 1")
        End Try

        Muestradatos()
        MsgBox("El txt fue creado correctamente")
    End Sub


    Public Function buscardepto(ByVal depto As String) As String
        Dim DBCon As OdbcConnection
        Dim cadenaODBC As String



        If ComboSUELDO.Text = "NUBULA" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\NUBULA SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "INFORMATION THECNOLOGY INDUSTRIES" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\INFORMATION THECNOLOGY.FDB"
        End If

        If ComboSUELDO.Text = "FOLDUR" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\NEXTEL.FDB"
        End If

        If ComboSUELDO.Text = "MORGET" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\MORGET.FDB"
        End If

        If ComboSUELDO.Text = "GRUPO CONISAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\GRUPO CONISAL.FDB"
        End If


        If ComboSUELDO.Text = "WIPSI" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\WIPSI A C.FDB"
        End If


        If ComboSUELDO.Text = "MORGET SEMANAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
      ";PWD=ata8244;DBNAME=192.168.2.83" & _
   ":C:\microsip datos\1 MORGET SEMANAL.FDB"
        End If
        If ComboSUELDO.Text = "MORGET CATORCENAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
      ";PWD=ata8244;DBNAME= 192.168.2.83" & _
   ":C:\microsip datos\2 MORGET CATORCENAL.FDB"
        End If
        If ComboSUELDO.Text = "MORGET QUINCENAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
      ";PWD=ata8244;DBNAME=192.168.2.83" & _
   ":C:\microsip datos\3  MORGET QUINCENAL.FDB"
        End If
        If ComboSUELDO.Text = "MORGET MENSUAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\4 MORGET MENSUAL.FDB"
        End If

        ''agosto

        If ComboSUELDO.Text = "IT TELECOM" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\IT RESOURCES TELECOM SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "MORGET INTERNA" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\5 MORGET INTERNA.FDB"
        End If

        If ComboSUELDO.Text = "AICEL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\AICEL.FDB"
        End If

        If ComboSUELDO.Text = "CONSORCIO ATERAP SA DE CV" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\CONSORCIO ATERAP SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "CROTEC SA DE CV" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\CROTEC SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "PEPSAT SA DE CV" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\PEPSAT SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "UPHETILOLI 2" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\UPHETILOLI 2.FDB"
        End If


        DBCon = New OdbcConnection(cadenaODBC)

        Dim consulta As String
        Dim resultado As String
        consulta = "select DEPTO_NO_ID  from DEPTOS_NO " & _
           "where NOMBRE  = '" & depto & "'"



        Try
            'Abrimos la conexión y comprobamos que no hay error
            Using comm As New OdbcCommand(consulta, DBCon)
                With comm

                    .CommandType = CommandType.Text

                    '.Parameters.Add(Bempleado)
                End With

                DBCon.Open()
                resultado = comm.ExecuteScalar()

            End Using
            ' MsgBox("Conexion realizada satsfactoriamente")
            If resultado Is Nothing Then
                resultado = ""
            End If


            Return resultado.ToLower
        Catch ex As Odbc.OdbcException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox(ex.Message.ToString)
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Function


    Public Function buscarpuesto(ByVal puesto As String) As String
        Dim DBCon As OdbcConnection
        Dim cadenaODBC As String



        If ComboSUELDO.Text = "NUBULA" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\NUBULA SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "INFORMATION THECNOLOGY INDUSTRIES" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\INFORMATION THECNOLOGY.FDB"
        End If

        If ComboSUELDO.Text = "FOLDUR" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\NEXTEL.FDB"
        End If

        If ComboSUELDO.Text = "MORGET" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\MORGET.FDB"
        End If

        If ComboSUELDO.Text = "GRUPO CONISAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\GRUPO CONISAL.FDB"
        End If


        If ComboSUELDO.Text = "WIPSI" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
     ";PWD=ata8244;DBNAME=192.168.2.83" & _
  ":C:\microsip datos\WIPSI A C.FDB"
        End If


        If ComboSUELDO.Text = "MORGET SEMANAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
      ";PWD=ata8244;DBNAME=192.168.2.83" & _
   ":C:\microsip datos\1 MORGET SEMANAL.FDB"
        End If
        If ComboSUELDO.Text = "MORGET CATORCENAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
      ";PWD=ata8244;DBNAME= 192.168.2.83" & _
   ":C:\microsip datos\2 MORGET CATORCENAL.FDB"
        End If
        If ComboSUELDO.Text = "MORGET QUINCENAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
      ";PWD=ata8244;DBNAME=192.168.2.83" & _
   ":C:\microsip datos\3  MORGET QUINCENAL.FDB"
        End If
        If ComboSUELDO.Text = "MORGET MENSUAL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\4 MORGET MENSUAL.FDB"
        End If

        ''agosto

        If ComboSUELDO.Text = "IT TELECOM" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\IT RESOURCES TELECOM SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "MORGET INTERNA" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\5 MORGET INTERNA.FDB"
        End If

        If ComboSUELDO.Text = "AICEL" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\AICEL.FDB"
        End If

        If ComboSUELDO.Text = "CONSORCIO ATERAP SA DE CV" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\CONSORCIO ATERAP SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "CROTEC SA DE CV" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\CROTEC SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "PEPSAT SA DE CV" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\PEPSAT SA DE CV.FDB"
        End If

        If ComboSUELDO.Text = "UPHETILOLI 2" Then
            cadenaODBC = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA" & _
    ";PWD=ata8244;DBNAME=192.168.2.83" & _
 ":C:\microsip datos\UPHETILOLI 2.FDB"
        End If


        DBCon = New OdbcConnection(cadenaODBC)

        Dim consulta As String
        Dim resultado As String
        consulta = "select PUESTO_NO_ID  from PUESTOS_NO " & _
           "where NOMBRE  = '" & puesto & "'"



        Try
            'Abrimos la conexión y comprobamos que no hay error
            Using comm As New OdbcCommand(consulta, DBCon)
                With comm

                    .CommandType = CommandType.Text

                    '.Parameters.Add(Bempleado)
                End With

                DBCon.Open()
                resultado = comm.ExecuteScalar()

            End Using
            ' MsgBox("Conexion realizada satsfactoriamente")
            If resultado Is Nothing Then
                resultado = ""
            End If


            Return resultado.ToLower
        Catch ex As Odbc.OdbcException
            'Si hubiese error en la conexión mostramos el texto de la descripción
            MsgBox(ex.Message.ToString)
        Finally
            DBCon.Close()
            DBCon.Dispose()
        End Try


    End Function

End Class
