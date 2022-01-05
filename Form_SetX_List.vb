Public Class Form_SetX_List

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Read_Title()
        Read_Title2()
    End Sub
    Public Sub Read_Title()
        Dim filename1 As String
        Dim Line_1 As String
        Dim Line_2 As String


        ListBox1.Items.Clear()

        For i = 0 To (Form_Spec_CHK_Ctl.CMB_SaveType.Items.Count - 1)
            If i = 0 Then
                filename1 = Form_Spec.Get_file_Path_x & "\Index\Index_Note.txt"
            Else
                filename1 = Form_Spec.Get_file_Path_x & "\" & Form_Spec_CHK_Ctl.CMB_SaveType.Items(i) & "\Index\Index_Note.txt"
            End If


            If (IO.File.Exists(filename1)) Then


                '開啟證件
                Dim fs As IO.FileStream = New IO.FileStream(filename1, IO.FileMode.Open)
                Using sr As IO.StreamReader = New IO.StreamReader(fs, _
                                                            System.Text.Encoding.Default)

                    '送出相關設定
                    Line_1 = sr.ReadLine()                                  '讀取一行資料



                    Line_2 = sr.ReadLine()                                  '讀取一行資料
                    'ListBox1.Items.Add(Line_1 & ",   " & Line_2)
                    ListBox1.Items.Add(i & ": " & Line_1)


                    sr.Close()                                                          '關閉檔案

                End Using
            Else
                ListBox1.Items.Add(i & ": ---")
            End If
        Next i
    End Sub

    Public Sub Read_Title2()
        Dim filename1 As String
        Dim Line_1 As String
        Dim Line_2 As String

        ListBox2.Items.Clear()


        For i = 0 To (Form_Spec_CHK_Ctl.CMB_SaveType.Items.Count - 1)
            If i = 0 Then
                filename1 = Form_Spec.Get_file_Path_x & "\Index\compare\Index_Note.txt"
            Else
                filename1 = Form_Spec.Get_file_Path_x & "\" & Form_Spec_CHK_Ctl.CMB_SaveType.Items(i) & "\Index\compare\Index_Note.txt"
            End If


            If (IO.File.Exists(filename1)) Then


                '開啟證件
                Dim fs As IO.FileStream = New IO.FileStream(filename1, IO.FileMode.Open)
                Using sr As IO.StreamReader = New IO.StreamReader(fs, _
                                                            System.Text.Encoding.Default)

                    '送出相關設定
                    Line_1 = sr.ReadLine()                                  '讀取一行資料



                    Line_2 = sr.ReadLine()                                  '讀取一行資料
                    ListBox2.Items.Add(i & ": " & Line_1 & ",   " & Line_2)


                    sr.Close()                                                          '關閉檔案

                End Using
            Else
                ListBox2.Items.Add(i & ": ---")
            End If
        Next i
    End Sub

    Private Sub Form_SetX_List_Load(sender As Object, e As EventArgs) Handles Me.Load
        Read_Title()
        Read_Title2()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Form_Spec_CHK_Ctl.CMB_SaveType.SelectedIndex = ListBox1.SelectedIndex
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form_Note.Show()
    End Sub
End Class