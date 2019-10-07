Imports System.Data.SqlClient

Public Class Form1
    Dim tVentanaY As Integer = 133
    Dim tVentanaX As Integer = 292
    Dim anombre() As String
    Dim afechaini() As String
    Dim afechafin() As String
    Dim m As Integer
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'animacionventana()
        'Timer2.Start()
        'Timer1.Start()
        Dim nombre As String = ""
        Dim fecha As Date
        Dim fechaInicio As String = ""
        Dim lector As SqlDataReader
        Dim bandera As Boolean = False
        Dim n As Integer
        Try
            m = 1
            n = 0
            Using cnx = New SqlConnection("Data Source=192.168.3.254;Initial Catalog=Vistas;User ID=sa;Password=inventumc762$")
                cnx.Open()
                Using cmd As New SqlCommand()
                    cmd.Connection = cnx
                    cmd.CommandText = "SET DATEFORMAT dmy   select Nombre, fechaFin, fechaInicio from RH_periodoPrueba where fechaFin>=getdate() -1 and fechaFin<=getdate() + 5 order by fechaFin asc"


                    lector = cmd.ExecuteReader
                    While lector.Read
                        nombre = CStr(lector(0).ToString)
                        fecha = CStr(lector(1).ToString)
                        fechaInicio = CStr(CDate(lector(2).ToString))
                        Dim numdias As Integer = DateDiff(DateInterval.Day, Today, CDate(fecha))


                        If numdias <= 5 And numdias >= 0 Then
                            'Dim sonido = New System.Media.SoundPlayer
                            Select Case numdias
                                Case 0
                                    Label1.Text = "Hoy terminó el periodo de prueba de:"
                                Case 1
                                    Label1.Text = "Falta un dia para que finalize periodo de prueba de:"
                                Case 2
                                    Label1.Text = "Faltan dos dias para que finalize el periodo de prueba de:"
                                Case 3
                                    Label1.Text = "Faltan tres dias para que finalize el periodo de prueba de:"
                                Case 4
                                    Label1.Text = "Faltan cuatro dias para que finalize el periodo de prueba de:"
                                Case 5
                                    Label1.Text = "Faltan cinco dias para que finalize el periodo de prueba de:"
                            End Select

                            ReDim Preserve anombre(n)
                            ReDim Preserve afechaini(n)
                            ReDim Preserve afechafin(n)
                            anombre(n) = nombre
                            afechaini(n) = fechaInicio
                            afechafin(n) = Label1.Text
                            n = n + 1

                            Label2.Text = nombre
                            Label9.Text = fechaInicio
                            bandera = True
                            'sonido = New System.Media.SoundPlayer(My.Application.Info.DirectoryPath + "\Sound_WAV.wav")
                            'sonido.Play()
                        Else

                        End If
                    End While
                    lector.Close()
                    If bandera = False Then
                        Me.Close()
                    End If
                End Using
                cnx.Close()
            End Using
            If anombre.Length >= 1 Then
                otravez(0)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "HA ocurrido un error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Function animacionventana()
        Dim desktopSize As Size
        desktopSize = System.Windows.Forms.SystemInformation.PrimaryMonitorSize
        Dim height As Integer = desktopSize.Height
        Dim width As Integer = desktopSize.Width
        Me.Location = New Point(width - tVentanaX, height - tVentanaY + 180)

        tVentanaY = tVentanaY + 5
        If tVentanaY >= 354 Then
            Timer1.Stop()
        End If
        Return Nothing
    End Function

    Function otravez(ByVal veces As Integer)
        Dim sonido As System.Media.SoundPlayer
        sonido = New System.Media.SoundPlayer(My.Application.Info.DirectoryPath + "\Sound_WAV.wav")
        tVentanaY = 133
        animacionventana()
        Timer2.Start()
        Timer1.Start()
        Label1.Text = afechafin(veces)
        Label2.Text = anombre(veces)
        Label9.Text = afechaini(veces)

        sonido.Play()

        Return Nothing
    End Function

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        animacionventana()
    End Sub

    Private Sub Label7_Click(sender As System.Object, e As System.EventArgs) Handles Label7.Click

        If anombre.Length > 1 And m <= anombre.Length - 1 Then
            Me.Visible = True
            otravez(m)
            m = m + 1
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub PictureBox1_Click(sender As System.Object, e As System.EventArgs) Handles PictureBox1.Click
        If anombre.Length > 1 And m <= anombre.Length - 1 Then
            Me.Visible = True
            otravez(m)
            m = m + 1
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub Timer2_Tick(sender As System.Object, e As System.EventArgs) Handles Timer2.Tick
        Me.Visible = True
    End Sub
End Class
