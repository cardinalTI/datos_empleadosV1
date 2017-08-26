
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO

Public Class datoshandler

    Private m_con As String
    Private m_connODBC As SqlConnection



    Public Function regresadatos() As ArrayList

       Dim trODBC As SqlTransaction
        Try
            Dim cadenaODBC As String

            cadenaODBC = Me.m_con

            Dim OdbcDr As SqlDataReader
            Dim arreDatos As New ArrayList
            Dim strQuery As New System.Text.StringBuilder()
            'Me.m_conn.Open()
            Me.m_connODBC = New SqlConnection("data source= 192.168.2.82 ; Initial Catalog= microsip; user id=sicossadmi;password=ipp2012;")
            'Me.m_connODBC = New SqlConnection("data source= 192.168.2.83; Initial Catalog= microsip; user id=sicossadmi;password=ipp2012;")
            'Me.m_connODBC = New SqlConnection("data source= 189.190.172.169; Initial Catalog= microsip; user id=sicossadmi;password=ipp2012;")
            'Me.m_connODBC = New SqlConnection("Server=USUARIO-PC\SQLEXPRESS;Database=incidencias;integrated security=true")

            'Me.m_connODBC = New SqlConnection("data source= 192.168.2.83; Initial Catalog= incidencias; user id=sicossadmi;password=ipp2012;")
            'Me.m_connODBC = New SqlConnection("Server=(local)\sqlexpress10;Database=incidencias;integrated security=true")
            Me.m_connODBC.Open()
            trODBC = Me.m_connODBC.BeginTransaction(IsolationLevel.Serializable)
            Dim commODBC As New SqlCommand("", Me.m_connODBC, trODBC)
            With strQuery
                .Append("SELECT numero as numero,nombres as nombre ,apellido_paterno as app ,apellido_materno as apm  ,clave_puesto as cpt ,clave_depto as cld ,clave_frecuencia_pago as clfp ,num_reg_patronal as numrp ,forma_pago as formap ,contrato as contrato ,jornada as jornada ,regimen_fiscal as regimenf ,fecha_ingreso as fechai ,estatus as estatus  ,tipo_salario as tipos ,salario_diario as sald ,salario_integrado as sali ,rfc as rfc ,curp as curp ,registro_imss as regimss,direccion as direccion,cp as cp, contrato_sat as contratosat FROM usuario ")

            End With
            'comm.CommandText = strQuery.ToString
            commODBC.CommandText = strQuery.ToString
            'dr = com.ExecuteReader()
            OdbcDr = commODBC.ExecuteReader()
            While OdbcDr.Read()
                Dim c As New datos
                c.numero = OdbcDr("numero")
                c.nombres = OdbcDr("nombre")
                c.apellido_paterno = OdbcDr("app")
                c.apellido_pmaterno = OdbcDr("apm")
                c.clavepuesto = OdbcDr("cpt")
                c.clavedepto = OdbcDr("cld")
                c.clavefrecuenciapago = OdbcDr("clfp")
                c.regpatronal = OdbcDr("numrp")
                c.formapago = OdbcDr("formap")
                c.contrato = OdbcDr("contrato")
                c.jornada = OdbcDr("jornada")
                c.regimenfiscal = OdbcDr("regimenf")
                c.fechaingreso = OdbcDr("fechai")
                c.estatus = OdbcDr("estatus")
                c.tiposalario = OdbcDr("tipos")

                c.salariodiario = OdbcDr("sald")
                c.salariointegrado = OdbcDr("sali")


                c.rfc = OdbcDr("rfc")
                c.curp = OdbcDr("curp")
                c.registroimss = OdbcDr("regimss")
                c.direccion = OdbcDr("direccion")
                If c.direccion = "CIUDAD DE MEXICO" Or c.direccion = "MEXICO" Then
                    c.direccion = "DIF"
                ElseIf c.direccion = "PUEBLA" Then
                    c.direccion = "PUE"
                ElseIf c.direccion = "OAXACA" Then
                    c.direccion = "OAX"
                End If

                c.cp = OdbcDr("cp")
                '' c.cuenta = OdbcDr("cuenta")
                ' c.centro = OdbcDr("centro")
                c.contratosat = OdbcDr("contratosat")

                c.mensaje = c.numero + "," + c.nombres + "," + c.apellido_paterno + "," + c.apellido_pmaterno + "," + c.clavepuesto + "," + c.clavedepto + "," + c.clavefrecuenciapago + "," + c.regpatronal + "," + c.formapago + "," + c.contrato + "," + c.jornada + "," + c.regimenfiscal + "," + c.fechaingreso + "," + c.estatus + "," + c.tiposalario + "," + c.salariodiario + "," + c.salariointegrado + "," + c.rfc + "," + c.curp + "," + c.registroimss + "," + c.direccion + "," + c.cp + "," + c.contratosat
                'Me.mensaje3 = c.mensaje1 + c.mensaje1
                'MsgBox(c.mensaje)
                arreDatos.Add(c)
            End While
            Me.m_connODBC.Close()
            Return (arreDatos)
        Catch ex As Exception
            Try
                trODBC.Rollback()
            Catch ex1 As Exception
                MsgBox("error 3")
            End Try
        Finally
            Try
                Me.m_connODBC.Close()
            Catch ex As Exception
                MsgBox("error 4")
            End Try
        End Try
    End Function

    Function gridatxt(ByVal Grid As DataGridView) As Boolean
        Dim texto As StreamWriter
        Dim escribo As String
        Dim filas As Integer
        Dim columnas As Integer
        Dim titulo As String = ""
        Dim palabra As String
        Dim letras As Integer
        Dim nuevo As String
        filas = Grid.RowCount - 1
        columnas = Grid.ColumnCount - 1
        texto = New StreamWriter("c:\exportaciones\empleados.txt", True)
        'texto = New StreamWriter("c:\Users\soporte\Desktop\empleados.txt", True)
        'texto = New StreamWriter("c:\Users\Soporte\Desktop\incidencias.txt", True)
        Dim tamaño(columnas) As Integer
        For i = 0 To columnas
            tamaño(i) = 0
        Next
        For a = 0 To filas
            Grid.CurrentCell = Grid.Rows(a).Cells(0)
            For b = 0 To columnas
                titulo = Grid.Columns(b).Name
                If IsDBNull(Grid.CurrentRow.Cells.Item(titulo).Value) Then
                    'palabra = "NULL"
                    palabra = vbCrLf

                Else
                    palabra = Grid.CurrentRow.Cells.Item(titulo).Value
                End If
                letras = palabra.Length
                If letras > tamaño(b) Then
                    tamaño(b) = letras

                End If
            Next
        Next
        For a = 0 To filas
            escribo = ""
            Grid.CurrentCell = Grid.Rows(a).Cells(0)
            For b = 0 To columnas
                titulo = Grid.Columns(b).Name
                If b = 0 Then
                    If IsDBNull(Grid.CurrentRow.Cells.Item(titulo).Value) Then
                        'escribo = "NULL"
                        palabra = vbCrLf

                    Else
                        escribo = Grid.CurrentRow.Cells.Item(titulo).Value
                        letras = escribo.Length
                    End If
                Else
                    If IsDBNull(Grid.CurrentRow.Cells.Item(titulo).Value) Then
                        'escribo = "NULL"
                        palabra = vbCrLf
                    Else
                        nuevo = Grid.CurrentRow.Cells.Item(titulo).Value
                        letras = nuevo.Length
                        Do While letras < tamaño(b)
                            'nuevo = nuevo & ""
                            letras = letras + 1
                        Loop
                        escribo = escribo & "" & nuevo
                    End If
                End If
            Next
            texto.WriteLine(escribo)
        Next
        texto.Close()
    End Function

End Class
