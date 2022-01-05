Public Class Form_Note
    Public TypeNote As String = ""

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim filename1 As String
        Dim filenum As Integer

        filename1 = Form_Spec.FullName_Index_Path & "\Index_Note.txt"
        filenum = FreeFile()
        FileOpen(filenum, filename1, OpenMode.Output)




        'PrintLine(filenum, Form_Spec_CHK_Top.Default_Title & " - " & TextBox1.Text)
        'PrintLine(filenum, Form_Spec_All_Pc.Default_Title & " - " & TextBox1.Text)
        PrintLine(filenum, TextBox1.Text)

        'Form_Spec.CMB_Compare.Text
        PrintLine(filenum, Form_Spec.CMB_Compare.Text)
        FileClose(filenum)

        Read_FormTitle()

        Form_SetX_List.Read_Title()
        Form_SetX_List.Read_Title2()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Read_FormTitle()
    End Sub
    Public Sub Read_FormTitle()
        Dim filename1 As String


        filename1 = Form_Spec.FullName_Index_Path & "\Index_Note.txt"

        If (IO.File.Exists(filename1)) Then
            Dim Parameter_Line As String = ""

            '開啟證件
            Dim fs As IO.FileStream = New IO.FileStream(filename1, IO.FileMode.Open)
            Using sr As IO.StreamReader = New IO.StreamReader(fs, _
                                                        System.Text.Encoding.Default)

                '送出相關設定
                Parameter_Line = sr.ReadLine()                                  '讀取一行資料
                TypeNote = Parameter_Line
                Form_Spec_CHK_Top.Text = Form_Spec_CHK_Top.Default_Title & " - " & Parameter_Line
                Form_Spec_All_Pc.Text = Form_Spec_All_Pc.Default_Title & " - " & Parameter_Line
                Form_Spec_CHK_Ctl.Text = Form_Spec_CHK_Ctl.Default_Title & " - " & Parameter_Line

                Parameter_Line = sr.ReadLine()                                  '讀取一行資料

                If Parameter_Line <> "" Then
                    Form_Spec.CMB_Compare.Text = Parameter_Line
                    Form_Spec.Get_Compare_Type()
                    Form_Spec.Page_Now()
                    Debug.Print("NFS_8")
                    Form_Spec.Update_for_Sync()

                End If

                sr.Close()                                                          '關閉檔案

            End Using
        Else
            TypeNote = ""
            Form_Spec_CHK_Top.Text = Form_Spec_CHK_Top.Default_Title
            Form_Spec_CHK_Ctl.Text = Form_Spec_CHK_Ctl.Default_Title
            Form_Spec_All_Pc.Text = Form_Spec_All_Pc.Default_Title
        End If

    End Sub
    Public Function GetTitle() As String
        Return TypeNote
    End Function
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = Form_Spec_CHK_Top.Text
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox2.Text = Form_Spec.CMB_Compare.Text

    End Sub
End Class