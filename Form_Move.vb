Public Class Form_Move

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Create_All_Path(TB_Target.Text)




    End Sub
    '============================
    '   確認路徑是否存在
    '   不存在就會一個一個建立
    '============================
    Public Sub Create_All_Path(ByVal tFullPath As String)
        Dim full_path As String
        Dim x_path() As String
        full_path = tFullPath
        Dim temp_path As String = ""

        If full_path.Contains("\") Then
            x_path = full_path.Split("\")

            Debug.Print(LBound(x_path))
            Debug.Print(UBound(x_path))

            For i = 0 To (UBound(x_path) - 1) Step 2
                If i + 1 <= UBound(x_path) Then
                    temp_path = temp_path & x_path(i) & "\" & x_path(i + 1)
                Else
                    temp_path = temp_path & x_path(i)
                End If


                Debug.Print(temp_path)

                If Not IO.Directory.Exists(temp_path) Then

                    '如不存在，建立資料夾
                    IO.Directory.CreateDirectory(temp_path)
                End If

                If i + 2 <= UBound(x_path) Then
                    temp_path = temp_path & "\"
                End If
            Next

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'FolderBrowserDialog1.RootFolder = Form_Spec.Get_Compare_Path
        'Debug.Print("111:" & Environment.SpecialFolder.Startup)
        'FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.LocalApplicationData
        FolderBrowserDialog1.ShowDialog()
        TB_Source.Text = FolderBrowserDialog1.SelectedPath
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Read_Spec_List_file_to_Directory(Form_Spec.Get_Source_DIR & "\APP\Spec_List.txt", False, 12)
        Read_Spec_List_file_to_Directory(Form_Spec.Get_Source_DIR & "\APP\Spec_List.txt", False, 0)
    End Sub
    Public Sub Read_Spec_List_file_to_Directory(ByVal SourceFile As String, ByVal Dmsg As Boolean, ByVal EndNumber As Integer)       '[PH00]
        Dim Item_Counter As Integer = 0

        ' Form_PDF As Boolean = False

        Dim line_counter As Integer

        Dim Parameter_Line As String = ""

        Dim strArr2() As String

        line_counter = 0

        If (IO.File.Exists(SourceFile)) Then


            Dim fs_Parameter As IO.FileStream = New IO.FileStream(SourceFile, IO.FileMode.Open)
            Using sr_Parameter As IO.StreamReader = New IO.StreamReader(fs_Parameter, _
                                                        System.Text.Encoding.Default)




                While (sr_Parameter.EndOfStream <> True)





                    Parameter_Line = sr_Parameter.ReadLine()                                  '讀取一行資料 略過標題
                    'CMB_Spec_List.Items.Add(Parameter_Line)
                    ' CMB_Compare.Items.Add(Parameter_Line)
                    strArr2 = Parameter_Line.Split(",")


                    For ii = 0 To EndNumber
                        Dim temp_Target As String
                        Dim temp_Target2 As String
                        Dim temp_source3 As String

                        temp_Target = TB_Target.Text & "\spec\" & strArr2(0) & "\Set" & ii

                        ' temp_Target3 = TB_Source.Text & "\spec\" & strArr2(0) & "\Set" & ii
                        'Create_All_Path(TB_Target.Text & "\spec\" & strArr2(0) & "\Set" & ii)

                        Create_All_Path(temp_Target)



                        temp_source3 = Form_Spec.Get_file_Path_x() & "\set0\Normal"


                        temp_Target2 = temp_Target & "\Normal"

                        Debug.Print("1" & temp_Target2)
                        Debug.Print("2" & temp_source3)

                        If IO.Directory.Exists(temp_Target2) = True Then

                            My.Computer.FileSystem.MoveDirectory(temp_source3, temp_Target2, 3, 3)
                        Else
                            IO.Directory.CreateDirectory(temp_Target2)
                            My.Computer.FileSystem.MoveDirectory(temp_source3, temp_Target2, 3, 3)
                            'My.Computer.FileSystem.CopyDirectory(temp_Target2, temp_Target)
                            ' My.Computer.FileSystem.DeleteDirectory(temp_Target2, FileIO.DeleteDirectoryOption.DeleteAllContents)
                        End If

                        'temp_Target2 = temp_Target & "\All"
                        'If IO.Directory.Exists(temp_Target2) = True Then
                        '    Debug.Print("1" & temp_Target2)
                        '    Debug.Print("2" & temp_Target)
                        '    My.Computer.FileSystem.CopyDirectory(temp_Target2, temp_Target)
                        '    My.Computer.FileSystem.DeleteDirectory(temp_Target2, FileIO.DeleteDirectoryOption.DeleteAllContents)
                        'End If

                        'temp_Target2 = temp_Target & "\DATA_CSV"
                        'If IO.Directory.Exists(temp_Target2) = True Then
                        '    My.Computer.FileSystem.CopyDirectory(temp_Target2, temp_Target)
                        '    My.Computer.FileSystem.DeleteDirectory(temp_Target2, FileIO.DeleteDirectoryOption.DeleteAllContents)
                        'End If

                        'temp_Target2 = temp_Target & "\Index"
                        'If IO.Directory.Exists(temp_Target2) = True Then
                        '    My.Computer.FileSystem.CopyDirectory(temp_Target2, temp_Target)
                        '    My.Computer.FileSystem.DeleteDirectory(temp_Target2, FileIO.DeleteDirectoryOption.DeleteAllContents)
                        'End If

                        'temp_Target2 = temp_Target & "\LOG"
                        'If IO.Directory.Exists(temp_Target2) = True Then
                        '    My.Computer.FileSystem.CopyDirectory(temp_Target2, temp_Target)
                        '    My.Computer.FileSystem.DeleteDirectory(temp_Target2, FileIO.DeleteDirectoryOption.DeleteAllContents)
                        'End If

                        'temp_Target2 = temp_Target & "\Normal"
                        'If IO.Directory.Exists(temp_Target2) = True Then
                        '    My.Computer.FileSystem.CopyDirectory(temp_Target2, temp_Target)
                        '    My.Computer.FileSystem.DeleteDirectory(temp_Target2, FileIO.DeleteDirectoryOption.DeleteAllContents)
                        'End If

                        'temp_Target2 = temp_Target & "\sp"
                        'If IO.Directory.Exists(temp_Target2) = True Then
                        '    My.Computer.FileSystem.CopyDirectory(temp_Target2, temp_Target)
                        '    My.Computer.FileSystem.DeleteDirectory(temp_Target2, FileIO.DeleteDirectoryOption.DeleteAllContents)
                        'End If
                    Next ii



                End While


                sr_Parameter.Close()                                                          '關閉檔案



            End Using

        Else



            If Dmsg = True Then
                MsgBox("Please check: " & SourceFile & ", 此為Log檔案的路徑")
            End If

        End If



    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub Form_Move_Load(sender As Object, e As EventArgs) Handles Me.Load

        If Form_Spec.Get_Spec_Path_file_Version = 2 Then
            TB_BasePath_New.Text = Form_Spec.Get_Source_DIR
            TB_WorkPath_New.Text = Form_Spec.Get_work_DIR
            TB_ExportPath_New.Text = Form_Spec.Get_Export_DIR
        Else
            TB_BasePath.Text = Form_Spec.Get_Source_DIR
            TB_WorkPath.Text = Form_Spec.Get_work_DIR
            TB_ExportPath.Text = Form_Spec.Get_Export_DIR
        End If

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Form_status.Open_X_file(LB_Setup_Filename.Text)
    End Sub
End Class