Public Class Form1
    Dim posicion As Integer
    Dim temp As String
    Dim comilla As Char
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        comilla = Chr(34)
        posicion = 0
        temp = ""

    End Sub
    Public Sub S0()
        Dim actual As Char
        actual = RichTextBox1.Text(posicion)
       
        Select Case actual

            Case comilla
                RichTextBox2.Text = actual
                temp = temp + actual

                posicion = posicion + 1

                S1()
        End Select

    End Sub

    Public Sub S1()
        Dim actual As Char
        actual = RichTextBox1.Text(posicion)
     
        Label1.Text = actual
        Select Case actual
            Case comilla
                temp = temp + actual
                posicion = posicion + 1
                RichTextBox2.Text = actual
                RichTextBox2.Text = ("Ingreso Una cadena de Texto " + temp)
                S2()

            Case Else
                temp += actual
                posicion = posicion + 1
              
                RichTextBox2.Text = ("Ingreso Una cadena de Texto " + temp)

                S1()
        End Select
    End Sub

    Public Sub S2()

        RichTextBox2.Text = "Finalizo"
        temp = ""
        posicion = posicion + 1
        S0()

    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim i As Integer
        Dim cadena As String = RichTextBox1.Text
        For i = 0 To RichTextBox1.TextLength - 11
            If cadena.Substring(i, 11) = "<TOPOLOGIA>" Then
                RichTextBox1.Select(i, 11)
                RichTextBox1.SelectionColor = Color.Green
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 12
            If cadena.Substring(i, 12) = "</TOPOLOGIA>" Then
                RichTextBox1.Select(i, 12)
                RichTextBox1.SelectionColor = Color.Green

                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 6
            If cadena.Substring(i, 6) = "<HOST>" Then
                RichTextBox1.Select(i, 6)
                RichTextBox1.SelectionColor = Color.Blue
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 7
            If cadena.Substring(i, 7) = "</HOST>" Then
                RichTextBox1.Select(i, 7)
                RichTextBox1.SelectionColor = Color.Blue
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 4
            If cadena.Substring(i, 4) = "<ID>" Then
                RichTextBox1.Select(i, 4)
                RichTextBox1.SelectionColor = Color.Pink
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 5
            If cadena.Substring(i, 5) = "</ID>" Then
                RichTextBox1.Select(i, 5)
                RichTextBox1.SelectionColor = Color.Pink
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 9
            If cadena.Substring(i, 9) = "<USUARIO>" Then
                RichTextBox1.Select(i, 9)
                RichTextBox1.SelectionColor = Color.Pink
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 10
            If cadena.Substring(i, 10) = "</USUARIO>" Then
                RichTextBox1.Select(i, 10)
                RichTextBox1.SelectionColor = Color.Pink
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 6
            If cadena.Substring(i, 6) = "<TIPO>" Then
                RichTextBox1.Select(i, 6)
                RichTextBox1.SelectionColor = Color.Pink
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 7
            If cadena.Substring(i, 7) = "</TIPO>" Then
                RichTextBox1.Select(i, 7)
                RichTextBox1.SelectionColor = Color.Pink
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 8
            If cadena.Substring(i, 8) = "<SWITCH>" Then
                RichTextBox1.Select(i, 8)
                RichTextBox1.SelectionColor = Color.Blue
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 9
            If cadena.Substring(i, 9) = "</SWITCH>" Then
                RichTextBox1.Select(i, 9)
                RichTextBox1.SelectionColor = Color.Blue
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 12
            If cadena.Substring(i, 12) = "<CONECTADOS>" Then
                RichTextBox1.Select(i, 12)
                RichTextBox1.SelectionColor = Color.Orange
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 13
            If cadena.Substring(i, 13) = "</CONECTADOS>" Then
                RichTextBox1.Select(i, 13)
                RichTextBox1.SelectionColor = Color.Orange
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 1
            If cadena.Substring(i, 1) = "1" Then
                RichTextBox1.Select(i, 1)
                RichTextBox1.SelectionColor = Color.Red
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 1
            If cadena.Substring(i, 1) = "2" Then
                RichTextBox1.Select(i, 1)
                RichTextBox1.SelectionColor = Color.Red
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 1
            If cadena.Substring(i, 1) = "3" Then
                RichTextBox1.Select(i, 1)
                RichTextBox1.SelectionColor = Color.Red
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 1
            If cadena.Substring(i, 1) = "4" Then
                RichTextBox1.Select(i, 1)
                RichTextBox1.SelectionColor = Color.Red
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 1
            If cadena.Substring(i, 1) = "5" Then
                RichTextBox1.Select(i, 1)
                RichTextBox1.SelectionColor = Color.Red
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 1
            If cadena.Substring(i, 1) = "6" Then
                RichTextBox1.Select(i, 1)
                RichTextBox1.SelectionColor = Color.Red
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 1
            If cadena.Substring(i, 1) = "7" Then
                RichTextBox1.Select(i, 1)
                RichTextBox1.SelectionColor = Color.Red
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 1
            If cadena.Substring(i, 1) = "8" Then
                RichTextBox1.Select(i, 1)
                RichTextBox1.SelectionColor = Color.Red
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 1
            If cadena.Substring(i, 1) = "9" Then
                RichTextBox1.Select(i, 1)
                RichTextBox1.SelectionColor = Color.Red
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 1
            If cadena.Substring(i, 1) = "0" Then
                RichTextBox1.Select(i, 1)
                RichTextBox1.SelectionColor = Color.Red
                i = i + 1
            End If
        Next
        For i = 0 To RichTextBox1.TextLength - 1
            If cadena.Substring(i, 1) = "," Then
                RichTextBox1.Select(i, 1)
                RichTextBox1.SelectionColor = Color.Red
                i = i + 1
            End If

        Next
        S0()



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim SaveFileDialog1 As New SaveFileDialog

        SaveFileDialog1.Filter = ".LFP | *.LFP"

        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            RichTextBox1.SaveFile(SaveFileDialog1.FileName, RichTextBoxStreamType.RichText)
            '
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim OpenFileDialog1 As New OpenFileDialog

        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop


        OpenFileDialog1.Filter = ".LFP |*.LFP"



        If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then
            Dim FileName As String = OpenFileDialog1.FileName
            Me.RichTextBox1.LoadFile(OpenFileDialog1.FileName, RichTextBoxStreamType.PlainText)

        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        End
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        MsgBox("Cristian Enrique Uz García" + vbCrLf + "201318609", MsgBoxStyle.OkOnly + vbExclamation, "Estudiante")

    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged


    End Sub
End Class
