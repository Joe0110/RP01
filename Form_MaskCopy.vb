Public Class Form_MaskCopy
    Public Full_Target_Path As String
    Public Log_File As String
    '================================
    '   過濾Item
    '================================

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i As Integer
        Dim PageName As String
        Dim op_src As String
        Dim op_trg As String
        Dim xx_array() As String


        If CKB_CopyALL.Checked = True Then
            RadioButton2.Checked = True

        End If

        TB_Color.BackColor = Color.Red

        '=================================================
        '   來源路徑                  =>   目標路徑
        '=================================================
        'Form_Spec.FullName_item_Path => Full_Target_Path
        '   確認檔案是否存在
        '   讀檔
        '   過濾
        '   寫入檔案
        '
        '   頁數: Form_Spec.Page_Default
        '-------------------------------------------------
        If Not IO.Directory.Exists(Full_Target_Path) Then
            '如不存在，建立資料夾
            IO.Directory.CreateDirectory(Full_Target_Path)
        End If

        For i = 1 To Form_Spec.Page_Default
            PageName = "Page_" & i.ToString("000") & ".csv"
            op_src = Form_Spec.FullName_item_Path & "\" & PageName
            op_trg = Full_Target_Path & "\" & PageName




            If (IO.File.Exists(op_src)) Then
                '------------------
                '   讀取檔案 - Start
                '------------------
                Dim Parameter_Line As String = ""



                '開啟證件
                Dim fs As IO.FileStream = New IO.FileStream(op_src, IO.FileMode.Open)
                Using sr As IO.StreamReader = New IO.StreamReader(fs, _
                                                            System.Text.Encoding.Default)


                    Dim TARGET_filenum As Integer

                    TARGET_filenum = FreeFile()
                    If RadioButton1.Checked = True Then
                        FileOpen(TARGET_filenum, op_trg, OpenMode.Append)
                    Else
                        FileOpen(TARGET_filenum, op_trg, OpenMode.Output)
                    End If



                    While (sr.EndOfStream <> True)

                        '送出相關設定
                        Parameter_Line = sr.ReadLine()                                    '讀取一行資料

                        If CKB_CopyALL.Checked = True Then
                            '--------------------
                            '   全部備份
                            '--------------------
                            PrintLine(TARGET_filenum, Parameter_Line)
                        Else

                            If RadioButton3.Checked = True Then


                                '-----------------
                                '   過濾文字
                                '-----------------
                                If Parameter_Line.Contains(TextBox1.Text) Then
                                    'PrintLine(TARGET_filenum, Parameter_Line)

                                    '-----------------
                                    '   代換顏色
                                    '-----------------
                                    If CheckBox1.Checked = True Then
                                        Dim NewStr As String = ""
                                        Dim ii As Integer
                                        xx_array = Parameter_Line.Split(",")
                                        For ii = 0 To (UBound(xx_array) - 1)
                                            If ii <> 13 Then
                                                NewStr = NewStr & xx_array(ii) & ","
                                            Else
                                                NewStr = NewStr & ComboBox2.SelectedIndex & ","
                                            End If

                                        Next
                                        NewStr = NewStr & xx_array(ii)
                                        PrintLine(TARGET_filenum, NewStr)
                                    Else
                                        PrintLine(TARGET_filenum, Parameter_Line)
                                    End If
                                End If
                            Else
                                '-----------------
                                '   過濾顏色
                                '-----------------
                                xx_array = Parameter_Line.Split(",")



                                If xx_array(13) = Val(TextBox1.Text) Then

                                    '-----------------
                                    '   代換顏色
                                    '-----------------
                                    If CheckBox1.Checked = True Then
                                        Dim NewStr As String = ""
                                        Dim ii As Integer

                                        For ii = 0 To (UBound(xx_array) - 1)
                                            If ii <> 13 Then
                                                NewStr = NewStr & xx_array(ii) & ","
                                            Else
                                                NewStr = NewStr & ComboBox2.SelectedIndex & ","
                                            End If

                                        Next
                                        NewStr = NewStr & xx_array(ii)
                                        PrintLine(TARGET_filenum, NewStr)
                                    Else
                                        PrintLine(TARGET_filenum, Parameter_Line)
                                    End If


                                End If

                            End If
                        End If
                    End While

                    FileClose(TARGET_filenum)


                    sr.Close()                                                          '關閉檔案

                End Using

                '------------------
                '   讀取檔案 - End
                '------------------

            End If
        Next

        TB_Color.BackColor = Color.YellowGreen

    End Sub

    Private Sub Form_MaskCopy_Load(sender As Object, e As EventArgs) Handles Me.Load
        'CMB_SaveType
        ComboBox_Copy(Form_Spec_CHK_Ctl.CMB_SaveType, ComboBox1)
        ComboBox_Copy(Form_Spec_CHK_Ctl.CMB_BG_Color, ComboBox2)
    End Sub

    Public Sub ComboBox_Copy(ByVal CSrc As ComboBox, ByRef CTrg As ComboBox)

        '

        'For i = 0 To CSrc.Items.Count - 1
        '    CTrg.Items.Add(CSrc.Items(i))
        'Next

        CTrg.Items.AddRange(CSrc.Items.Cast(Of String).ToArray())
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem = "" Then
            Full_Target_Path = Form_Spec.Get_file_Path_x & "\DATA_CSV"
        Else
            Full_Target_Path = Form_Spec.Get_file_Path_x & "\" & ComboBox1.SelectedItem & "\DATA_CSV"
        End If


    End Sub
   


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim i As Integer
        Dim PageName As String
        Dim op_src As String
        Dim op_trg As String
        Dim xx_array() As String


        If CKB_CopyALL.Checked = True Then
            RadioButton2.Checked = True

        End If

        TB_Color.BackColor = Color.Red

        '=================================================
        '   來源路徑                  =>   目標路徑
        '=================================================
        'Form_Spec.FullName_item_Path => Full_Target_Path
        '   確認檔案是否存在
        '   讀檔
        '   過濾
        '   寫入檔案
        '
        '   頁數: Form_Spec.Page_Default
        '-------------------------------------------------
        If Not IO.Directory.Exists(Full_Target_Path) Then
            '如不存在，建立資料夾
            IO.Directory.CreateDirectory(Full_Target_Path)
        End If


        For i = 1 To Form_Spec.Page_Default
            PageName = "Page_" & i.ToString("000") & ".csv"
            op_src = Form_Spec.FullName_item_Path & "\" & PageName
            op_trg = Full_Target_Path & "\" & PageName




            If (IO.File.Exists(op_src)) Then
                '------------------
                '   讀取檔案 - Start
                '------------------
                Dim Parameter_Line As String = ""



                '開啟證件
                Dim fs As IO.FileStream = New IO.FileStream(op_src, IO.FileMode.Open)
                Using sr As IO.StreamReader = New IO.StreamReader(fs, _
                                                            System.Text.Encoding.Default)


                    Dim TARGET_filenum As Integer

                    TARGET_filenum = FreeFile()
                    If RadioButton1.Checked = True Then
                        FileOpen(TARGET_filenum, op_trg, OpenMode.Append)
                    Else
                        FileOpen(TARGET_filenum, op_trg, OpenMode.Output)
                    End If



                    While (sr.EndOfStream <> True)

                        '送出相關設定
                        Parameter_Line = sr.ReadLine()                                    '讀取一行資料

                        If CKB_CopyALL.Checked = True Then
                            '--------------------
                            '   全部備份
                            '--------------------
                            PrintLine(TARGET_filenum, Parameter_Line)
                        Else

                            '--------------------
                            '   條件備份
                            '--------------------
                            If RadioButton3.Checked = True Then


                                '-----------------
                                '   過濾文字
                                '-----------------
                                If Parameter_Line.Contains(TextBox1.Text) Then
                                    'PrintLine(TARGET_filenum, Parameter_Line)

                                    '-----------------
                                    '   代換顏色
                                    '-----------------
                                    If CheckBox1.Checked = True Then
                                        Dim NewStr As String = ""
                                        Dim ii As Integer
                                        xx_array = Parameter_Line.Split(",")

                                        For ii = 0 To (UBound(xx_array) - 1)
                                            If ii <> 13 Then
                                                NewStr = NewStr & xx_array(ii) & ","
                                            Else
                                                '-----------
                                                '   顏色欄位
                                                '-----------
                                                NewStr = NewStr & ComboBox2.SelectedIndex & ","
                                            End If

                                        Next

                                        NewStr = NewStr & xx_array(ii)
                                        PrintLine(TARGET_filenum, NewStr)
                                    Else

                                        PrintLine(TARGET_filenum, Parameter_Line)
                                    End If
                                End If
                            Else
                                '-----------------
                                '   過濾顏色
                                '-----------------
                                xx_array = Parameter_Line.Split(",")



                                If xx_array(13) = Val(TextBox1.Text) Then

                                    '-----------------
                                    '   代換顏色
                                    '-----------------
                                    If CheckBox1.Checked = True Then
                                        Dim NewStr As String = ""
                                        Dim ii As Integer

                                        For ii = 0 To (UBound(xx_array) - 1)
                                            If ii <> 13 Then
                                                NewStr = NewStr & xx_array(ii) & ","
                                            Else
                                                NewStr = NewStr & ComboBox2.SelectedIndex & ","
                                            End If

                                        Next
                                        NewStr = NewStr & xx_array(ii)
                                        PrintLine(TARGET_filenum, NewStr)
                                    Else
                                        PrintLine(TARGET_filenum, Parameter_Line)
                                    End If


                                End If

                            End If
                        End If
                    End While

                    FileClose(TARGET_filenum)


                    sr.Close()                                                          '關閉檔案

                End Using

                '------------------
                '   讀取檔案 - End
                '------------------

            End If
        Next

        TB_Color.BackColor = Color.YellowGreen

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '--------------------
        '   產生markdown
        '--------------------
        MsgBox("產生MarkDown檔案(未完成)")
    End Sub
End Class