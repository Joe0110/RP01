Imports System.IO '新增此命名空間
Imports System.Math

Public Class Form_Report
    Public History_FileName As String

    Public NoteMsg1 As String = ""
    Public Have_SVN_Diffe As Boolean = False
    Public SVN_Diffe_Name As String

    'Public ReleaseFileName As String = "d:\s_List\Release_LOG.csv"
    'Public ReleaseFileName As String = Form_Main.JPCB_LOG_Path & "\Release_LOG.csv"
    'Public ReleaseFileName As String = Form_Main.JPCB_LOG_Path & "\REC00001_Release_LOG.csv"
    Public ReportFileName As String = ""

    Public Last_v1 As Integer = 0
    Public Last_v2 As Integer = 0
    Public Last_v3 As Integer = 0
    Public Last_v4 As Integer
    Public Chart_color As Integer = 0


    Public Trg_Filename As String


    Public total_OK As Integer
    Public total_NG As Integer
    Public total_Time As Integer



    Private Sub Form_Report_Load(sender As Object, e As EventArgs) Handles Me.Load



    End Sub

    Public Sub Chart_Add_Point(ByRef dChart As DataVisualization.Charting.Series, ByVal dType As String, ByVal dValeu As Integer, ByVal dLimit As Integer)
        Dim dtemp As Single

        dtemp = dValeu

        If dChart.ChartType = DataVisualization.Charting.SeriesChartType.Column Then
            'Debug.Print("1111")

            If dtemp = 0 Then
                dtemp = 1
            End If
        End If


        If dType <> "" Then
            dChart.Points.AddXY(dType, dtemp)
            If dChart.Points.Count = dLimit Then
                dChart.Points.RemoveAt(0)
            End If
        Else
            dChart.Points.AddY(dtemp)
            If dChart.Points.Count = dLimit Then
                dChart.Points.RemoveAt(0)
            End If
        End If

    End Sub
    Public Sub Dele_chart(ByVal Limit As Integer)
        For i = 0 To Limit Step 1

            Chart_Remove_Point(Chart1.Series(0))
            Chart_Remove_Point(Chart1.Series(1))
            Chart_Remove_Point(Chart1.Series(2))
            Chart_Remove_Point(Chart1.Series(3))
            Chart_Remove_Point(Chart1.Series(4))
            Chart_Remove_Point(Chart1.Series(5))
            Chart_Remove_Point(Chart1.Series(6))
            Chart_Remove_Point(Chart1.Series(7))
            Chart_Remove_Point(Chart1.Series(8))
            Chart_Remove_Point(Chart1.Series(9))
            Chart_Remove_Point(Chart1.Series(10))
            Chart_Remove_Point(Chart1.Series(11))
            Chart_Remove_Point(Chart1.Series(12))
            Chart_Remove_Point(Chart1.Series(13))
            Chart_Remove_Point(Chart1.Series(14))
            Chart_Remove_Point(Chart1.Series(15))
            Chart_Remove_Point(Chart1.Series(16))
            Chart_Remove_Point(Chart1.Series(17))
            Chart_Remove_Point(Chart1.Series(18))
            Chart_Remove_Point(Chart1.Series(19))
            Chart_Remove_Point(Chart1.Series(20))
            Chart_Remove_Point(Chart1.Series(21))
            Chart_Remove_Point(Chart1.Series(22))
            Chart_Remove_Point(Chart1.Series(23))
            Chart_Remove_Point(Chart1.Series(24))
            Chart_Remove_Point(Chart1.Series(25))
            Chart_Remove_Point(Chart1.Series(26))
            ' Chart_Remove_Point(Chart1.Series(27))
            'Chart_Remove_Point(Chart1.Series(28))
            'Chart_Remove_Point(Chart1.Series(19))
            'Chart_Remove_Point(Chart1.Series(20))
            'Chart_Remove_Point(Chart1.Series(21))

        Next
    End Sub
    Public Sub Chart_Remove_Point(ByRef dChart As DataVisualization.Charting.Series)
        If dChart.Points.Count <> 0 Then
            dChart.Points.RemoveAt(0)
        End If
    End Sub




    Private Function Get_History_FileName(ByVal Ori_str As String) As String
        Dim xx() As String
        Dim NewName As String




        xx = Ori_str.Split(".")

        NewName = xx(0) & "." & xx(1) & "_History." & xx(2)
        Return NewName
    End Function






 





  


    Private Sub Button4_Click(sender As Object, e As EventArgs)
        '=================
        ' 開啟Release Log
        '=================
        ' System.Diagnostics.Process.Start(ReleaseFileName)

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        '===================
        '　開啟Report
        '===================
        'Label18.Text = "Open:" & Trg_Filename
        ReportFileName = Trg_Filename
        'Label17.Text = "Open:" & ReportFileName

        '        ReportFileName

        '============檢查檔案是否存在==========
        If (File.Exists(ReportFileName)) Then



            'If (ReportFileName <> "") Then
            System.Diagnostics.Process.Start(ReportFileName)
            'Else
            '   MsgBox("No (Report) File Name")
            'End If
        Else
            MsgBox("No (Report) File Name")
        End If
    End Sub




    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        'Get_Report_File(0)
        History_FileName = Form_status.Log_File
        Read_Content_to_HTML(History_FileName)
        Chart1.ChartAreas(0).RecalculateAxesScale()
    End Sub
    Private Sub Read_Content_to_HTML(ByVal FNAme As String)
        Dim Parameter_Line As String = ""
        Dim Display_Line As String = ""
        Dim str_array() As String



        'Button21.Enabled = False



        '開啟證件
        Dim fs As FileStream = New FileStream(FNAme, FileMode.Open)
        Using sr As StreamReader = New StreamReader(fs, _
                                                    System.Text.Encoding.Default)


            'Parameter_Line = sr.ReadLine() & vbCrLf                                     '讀取一行資料 略過標題



            While (sr.EndOfStream <> True)



                '送出相關設定
                Parameter_Line = sr.ReadLine() & vbCrLf                                     '讀取一行資料
                ' data_process_PWM(Parameter_Line)


                str_array = Parameter_Line.Split(",")

                'If ((str_array(1) <> Last_v1) And (str_array(2) <> Last_v2) And (str_array(3) <> Last_v3)) Then
                '=================================
                '   相同OK數時, 不重覆繪製的處理
                '=================================
                Dim qqq As Integer
                qqq = UBound(str_array)
                Debug.Print("QQQ" & qqq)
                If qqq >= 27 Then
                    If ((str_array(1).Replace("%", "") <> Last_v1)) Then
                        Chart_Add_Point(Chart1.Series(0), str_array(0), str_array(1).Replace("%", ""), 100)

                        ' qqq = UBound(str_array)
                        ' Debug.Print("QQQ" & qqq)

                        If qqq > 1 Then
                            For i = 1 To (qqq - 1)
                                Chart_Add_Point(Chart1.Series(i), "", x_String_process(str_array(i + 1)), 100)
                            Next
                        End If

                        'Chart_Add_Point(Chart1.Series(1), "", x_String_process(str_array(2)), 100)
                        'Chart_Add_Point(Chart1.Series(2), "", x_String_process(str_array(3)), 100)
                        'Chart_Add_Point(Chart1.Series(3), "", x_String_process(str_array(4)), 100)
                        'Chart_Add_Point(Chart1.Series(4), "", x_String_process(str_array(5)), 100)
                        'Chart_Add_Point(Chart1.Series(5), "", x_String_process(str_array(6)), 100)
                        'Chart_Add_Point(Chart1.Series(6), "", x_String_process(str_array(7)), 100)
                        'Chart_Add_Point(Chart1.Series(7), "", x_String_process(str_array(8)), 100)
                        'Chart_Add_Point(Chart1.Series(8), "", x_String_process(str_array(9)), 100)
                        'Chart_Add_Point(Chart1.Series(9), "", x_String_process(str_array(10)), 100)
                        'Chart_Add_Point(Chart1.Series(10), "", x_String_process(str_array(11)), 100)

                        'Chart_Add_Point(Chart1.Series(11), "", x_String_process(str_array(12)), 100)
                        'Chart_Add_Point(Chart1.Series(12), "", x_String_process(str_array(13)), 100)
                        'Chart_Add_Point(Chart1.Series(13), "", x_String_process(str_array(14)), 100)
                        'Chart_Add_Point(Chart1.Series(14), "", x_String_process(str_array(15)), 100)
                        'Chart_Add_Point(Chart1.Series(15), "", x_String_process(str_array(16)), 100)
                        'Chart_Add_Point(Chart1.Series(16), "", x_String_process(str_array(17)), 100)
                        'Chart_Add_Point(Chart1.Series(17), "", x_String_process(str_array(18)), 100)
                        'Chart_Add_Point(Chart1.Series(18), "", x_String_process(str_array(19)), 100)
                        'Chart_Add_Point(Chart1.Series(19), "", x_String_process(str_array(20)), 100)
                        'Chart_Add_Point(Chart1.Series(20), "", x_String_process(str_array(21)), 100)

                        'Chart_Add_Point(Chart1.Series(21), "", x_String_process(str_array(22)), 100)
                        'Chart_Add_Point(Chart1.Series(22), "", x_String_process(str_array(23)), 100)
                        'Chart_Add_Point(Chart1.Series(23), "", x_String_process(str_array(24)), 100)
                        'Chart_Add_Point(Chart1.Series(24), "", x_String_process(str_array(25)), 100)
                        'Chart_Add_Point(Chart1.Series(25), "", x_String_process(str_array(26)), 100)
                        'Chart_Add_Point(Chart1.Series(26), "", x_String_process(str_array(27)), 100)
                        'Chart_Add_Point(Chart1.Series(27), "", x_String_process(str_array(28)), 100)

                        'Chart_Add_Point(Chart1.Series(18), "", Val(str_array(19)), 100)
                        'Chart_Add_Point(Chart1.Series(19), "", Val(str_array(20)), 100)
                        'Chart_Add_Point(Chart1.Series(20), "", Val(str_array(21)), 100)
                    End If

                    Last_v1 = str_array(1).Replace("%", "")
                    ' Last_v2 = str_array(2)
                    ' Last_v3 = str_array(3)
                End If

            End While




            'Label134.Text = "END"


            sr.Close()                                                          '關閉檔案




        End Using

    End Sub
    Public Function x_String_process(ByVal AAA As String) As ValidationConstraints
        Dim x1 As String
        Dim x1_array() As String
        x1 = AAA
        x1_array = x1.Split("_")

        If RadioButton1.Checked = True Then
            Return Val(x1_array(0))
        Else
            Return Val(x1_array(1))
        End If


    End Function
    ' 遞迴搜尋子目錄內的符合檔案, 取代 Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories 屬性之 Bug
    ' 需一個 ListBox 物件
    ' 取得目錄搜尋列表
    Public Sub GetFileList(ByRef 路徑 As String, ByRef 檔名 As String, ByRef 接收容器 As ListBox, Optional ByRef 深度搜尋 As Boolean = False)
        Dim dirAttr As Microsoft.VisualBasic.FileAttribute
        Try
            For Each fd In My.Computer.FileSystem.GetFiles(路徑, Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, 檔名)
                接收容器.Items.Add(fd.ToString)
            Next

            ' 喘口氣
            My.Application.DoEvents()

            If 深度搜尋 Then
                For Each ff In My.Computer.FileSystem.GetDirectories(路徑)
                    dirAttr = FileSystem.GetAttr(ff.ToString)
                    If (dirAttr And FileAttribute.System) = 0 Then GetFileList(ff.ToString, 檔名, 接收容器, 深度搜尋)
                Next
            End If
        Catch ex As UnauthorizedAccessException
        Catch ex As IOException
        Catch ex As Exception
        End Try
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dele_chart(100)
    End Sub












    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim xfilename As String

        xfilename = Form_Spec.FullName_ALL & "\Status_All_C.png"
        ' End If

        Form_Total_Items.Capture_Screen_to_PNG(Me, xfilename)

        xfilename = Form_Spec.FullName_ALL & "\Status_All_C_" & Now.ToString("yyyyMMdd_HHmmss") & ".Png"


        Form_Total_Items.Capture_Screen_to_PNG(Me, xfilename)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim xfilename As String

        xfilename = Form_Spec.FullName_ALL & "\Status_All_D.png"
        ' End If

        Form_Total_Items.Capture_Screen_to_PNG(Me, xfilename)

        xfilename = Form_Spec.FullName_ALL & "\Status_All_D_" & Now.ToString("yyyyMMdd_HHmmss") & ".Png"


        Form_Total_Items.Capture_Screen_to_PNG(Me, xfilename)
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        'Dele_chart(100)
        'History_FileName = Form_status.Log_File
        'Read_Content_to_HTML(History_FileName)
        'Chart1.ChartAreas(0).RecalculateAxesScale()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Dele_chart(100)
        History_FileName = Form_status.Log_File
        Read_Content_to_HTML(History_FileName)
        Chart1.ChartAreas(0).RecalculateAxesScale()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click


        Chart_color += 1
        If Chart_color >= 7 Then
            Chart_color = 0
        End If


        Select Case Chart_color
            Case 0
                Chart1.ChartAreas(0).BackColor = Color.White
            Case 1
                Chart1.ChartAreas(0).BackColor = Color.Black
            Case 2
                Chart1.ChartAreas(0).BackColor = Color.DarkGray
            Case 3
                Chart1.ChartAreas(0).BackColor = Color.DarkBlue
            Case 4
                Chart1.ChartAreas(0).BackColor = Color.DarkGreen
            Case 5
                Chart1.ChartAreas(0).BackColor = Color.Pink
            Case 6
                Chart1.ChartAreas(0).BackColor = Color.LightPink
        End Select


    End Sub
End Class