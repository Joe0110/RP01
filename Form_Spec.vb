Imports System.IO '新增此命名空間

Public Class Form_Spec
    Public Use_Main_path As Boolean = False
    Public Use_Sub_path As Boolean = False

    Public SetX_Str As String = "Set0"
    Public Spec_Diff As New Form_Spec_All_Pc
    'Public Spec_Diff As Form_Spec_All_Pc
    Public spec_diff_Opened As Boolean = False
    Dim Spec_index_path As String
    Public Page_Exist(256) As Boolean

    Public Const SpecCHK_PathFile As String = "Spec_Path.txt"

    Public Const Title_String As String = "Joe Huang 2021_0626_2350"

    Public Spec_Path_x_file_Version As Integer
    Public Ori_Png_Name As String

    Public Main_Path As String  '優先存取路徑
    Public Sub_Path As String   '備用存取路徑
    Public Ori_Source_DIR As String = "d:\ECU_TEST"

    Public Main_workPath As String
    Public Sub_workPath As String

    Public Main_ExportPath As String
    Public Sub_ExportPath As String

    Public Wo_rk_DIR As String = "d:\ECU_TEST_Work"
    Public Export_DIR As String = "d:\ECU_TEST_Export"
    'Export_DIR

    Public LA_DIR As String
    Public Switch_Page_Done As Boolean = False
    Public Point_Page As Integer = 1

    Public Name_Default As String
    Public Name_Selected As String
    Public Page_Default As String
    Public Page_Selected As String

    Public Redmine_Server As String
    Public Redmine_Server2 As String
    Public Mantis_Server As String
    Public Mantis_Server2 As String

    Public NotePad_path As String

    Public Total_Pages As Integer

    Public Spec_Name As String
    Public Pattern_Type As String

    Public Spec_Compare_Name As String

    Public File_Path As String                              '完整工作路徑(含仕樣別)
    Public Work_Path As String                              '完整工作路徑(含仕樣別)
    Public Export_Path As String                              '完整工作路徑(含仕樣別)
    Public FilePath__Append As String                              '完整工作路徑(含仕樣別)
    Public WorkPath__Append As String                              '完整工作路徑(含仕樣別)
    Public ExportPath__Append As String                              '完整工作路徑(含仕樣別)
    Public compare_Spec_path As String                              '完整工作路徑(含仕樣別)

    Public FullName_item_CSV As String                  '工作的CSV檔
    Public FullName_item_Path As String                  '工作的CSV檔

    Public FullName_ALL As String                   '全部的縮圖檔名
    Public FullName_Index_Path As String                   '全部的Index_Path

    Public Full_Name_Small_Pic As String             '小圖檔名
    Public FullName_XSmall_Path As String             '小圖檔名

    Public Full_Name_CHK As String                   '工作檔名
    Public Full_Name_CHK_Path As String                   '工作檔名



    Public ReportPath As String

    Public Time_UP As Boolean = False
    Public Form_PDF As Boolean = False
    Public Sub Set_Compare_Path(ByVal AAA As String)
        compare_Spec_path = AAA

    End Sub
    Public Function Get_Compare_Path() As String
        Return compare_Spec_path

    End Function
    Public Sub Set_Source_DIR(ByVal AAA As String)
        Debug.Print("-----------------")
        Debug.Print("Set_Source_DIR:" & AAA)
        Debug.Print("-----------------")

        Ori_Source_DIR = AAA

    End Sub
    Public Function Get_Source_DIR() As String
        Return Ori_Source_DIR

    End Function
    Public Sub Set_work_DIR(ByVal AAA As String)
        Debug.Print("-----------------")
        Debug.Print("Set_work_DIR:" & AAA)
        Debug.Print("-----------------")
        Wo_rk_DIR = AAA

    End Sub
    Public Function Get_work_DIR() As String
        Return Wo_rk_DIR

    End Function

    'Get_Export_DIR
    Public Sub Set_Export_DIR(ByVal AAA As String)
        Export_DIR = AAA

    End Sub
    Public Function Get_Export_DIR() As String
        Return Export_DIR

    End Function

    Public Sub Set_Full_PathSmall_Name(ByVal AAA As String)
        FullName_XSmall_Path = AAA
    End Sub
    Public Function Get_Full_PathSmall_Name() As String
        If CKB_compare.Checked = True Then
            Return FullName_XSmall_Path & "\compare"
        Else
            Return FullName_XSmall_Path
        End If

    End Function
    Public Sub Set_Full_Small_Name(ByVal AAA As String)
        Full_Name_Small_Pic = AAA
    End Sub
    Public Function Get_Full_Small_Name() As String
        Return Full_Name_Small_Pic
    End Function
    Public Function Get_ExportPath_x_Append() As String
        Return ExportPath__Append
    End Function
    Public Sub Set_ExportPath_x_Append(ByVal AAA As String)
        ExportPath__Append = AAA
    End Sub
    Public Function Get_filePath_x_Append() As String
        Return FilePath__Append
    End Function
    Public Sub Set_filePath_x_Append(ByVal AAA As String)
        FilePath__Append = AAA
    End Sub
    Public Function Get_WorkPath_x_Append() As String
        Return WorkPath__Append
    End Function
    Public Sub Set_WorkPath_x_Append(ByVal AAA As String)
        WorkPath__Append = AAA
    End Sub
    Public Function Get_file_Path_x() As String
        Return File_Path
    End Function
    Public Sub Set_file_Path_x(ByVal AAA As String)
        File_Path = AAA
    End Sub
    Public Function Get_work_Path_x() As String
        Return Work_Path
    End Function
    Public Sub Set_work_Path_x(ByVal AAA As String)
        Work_Path = AAA
    End Sub
    Public Function Get_Export_Path_x() As String
        Return Export_Path
    End Function
    Public Sub Set_Export_Path_x(ByVal AAA As String)
        Export_Path = AAA
    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll

        Set_Page_Now(TrackBar1.Value)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '==================
        '   由 TB_Page_Now.Text 手動設定頁面 (Set 鍵)
        '==================
        TrackBar1.Value = Val(TB_Page_Now.Text)
        Change_Page(Val(TB_Page_Now.Text))
        Update_pic()

    End Sub
    Public Sub Update_pic()

        Debug.Print("Update__pic_")

        Form_Spec_CHK_Top.Release_LB()



        '-------------
        'Spec_Name
        '-------------

        Set_file_Path_x(Get_Source_DIR() & "\Spec\" & Spec_Name)
        Set_work_Path_x(Get_work_DIR() & "\Spec\" & Spec_Name)
        Set_Export_Path_x(Get_Export_DIR() & "\Spec\" & Spec_Name)

        Set_Compare_Path(Get_Source_DIR() & "\Spec\" & Spec_Compare_Name)



        Form_Spec_CHK_Ctl.ReNew_Working_SetX_Path(False)

        ReportPath = Get_ExportPath_x_Append() & "\Report"

        Debug.Print("ReportPath:" & ReportPath)

        Set_Full_Small_Name(Get_Full_PathSmall_Name() & "\Page_" & Get_Page_Now.ToString("000") & "_s.png")
        Set_CHK_PicName(Get_CHK_PicPath() & "\Page_" & Get_Page_Now.ToString("000") & "_chk.png")


        TB_CHK_PicName.Text = Get_CHK_PicName()




        If CKB_USE_WorkPIC.Checked = True Then
            '-------------------------------
            '   使用工作圖
            '-------------------------------
            If (File.Exists(Get_CHK_PicName)) Then

                PicBox_Work.Load(Get_CHK_PicName)
            Else

                Use_temp_Work_Img()
            End If

        Else
            '-------------------------------
            '   使用原始圖
            '-------------------------------
            PicBox_Work.Load(Get_Full_Name_Ori_Pic(0))
        End If


        TB_FName.Text = Get_Full_Name_Ori_Pic(0)


        Debug.Print("NFS_3")
        Update_for_Sync()


        Form_Spec_CHK_Top.PicBox_Arrow.Visible = False

    End Sub
    Public Function Get_Page_Now() As Integer
        Return Point_Page
    End Function
    Public Sub Set_Page_Now(ByVal AAA As Integer)
        Point_Page = AAA
        TB_Page_Now.Text = AAA
    End Sub
    Public Function Get_Full_Name_Ori_Pic(ByVal UsePage As Integer) As String
        If UsePage = 0 Then
            Return Get_file_Path_x() & "\Page_" & Get_Page_Now.ToString("000") & ".png"
        Else
            Return Get_file_Path_x() & "\Page_" & UsePage.ToString("000") & ".png"
        End If

    End Function
    Public Function Get_Full_Name_Compare_Pic(ByVal UsePage As Integer) As String
        If UsePage = 0 Then
            Return Get_Compare_Path() & "\Page_" & Get_Page_Now.ToString("000") & ".png"
        Else
            Return Get_Compare_Path() & "\Page_" & UsePage.ToString("000") & ".png"
        End If

    End Function
    Public Sub Set_Full_Name_Ori_Pic(ByVal AAA As String)
        Ori_Png_Name = AAA
    End Sub

    '===========================
    '
    '===========================
    Public Sub Set_CHK_PicName(ByVal AAA As String)
        Full_Name_CHK = AAA
    End Sub
    Public Function Get_CHK_PicName() As String
        Return Full_Name_CHK
    End Function
    '===========================
    '
    '===========================
    Public Sub Set_CHK_PicPath(ByVal AAA As String)
        Full_Name_CHK_Path = AAA
    End Sub
    Public Function Get_CHK_PicPath() As String

        If CKB_compare.Checked = True Then
            Return Full_Name_CHK_Path & "\compare"
        Else
            Return Full_Name_CHK_Path
        End If

    End Function

    Public Sub Picbox_Work_Reload()
        If CKB_USE_WorkPIC.Checked = True Then
            If (File.Exists(Get_CHK_PicName)) Then
                'PicBox_Work.Visible = True
                PicBox_Work.Load(Get_CHK_PicName)
            Else
                'PicBox_Work.Visible = False
                Use_temp_Work_Img()
            End If

        Else
            'PicBox_Work.Load(FullName_Ori_Pic)
        End If
    End Sub
    Public Function Get_CSV_Filename() As String
        Dim Rtn_Name As String
        ' Dim filename1 As String
        Dim xpage As Integer

        xpage = Get_Page_Now()

        'Debug.Print("From Get_CSV_Filename")
        'Form_Spec_CHK_Ctl.ReNew_Working_SetX_Path(False)

        ReNew_Working_SetX_Path(False)

        Rtn_Name = FullName_item_Path & "\Page_" & xpage.ToString("000") & ".csv"
        'Debug.Print("Rtn_Name:" & Rtn_Name)
        Return Rtn_Name
    End Function
    '========================================
    '   tCreat_DIR=true => 建立資料夾
    '========================================
    Public Sub ReNew_Working_SetX_Path(ByVal tCreat_DIR As Boolean)

        'Debug.Print("ReNew_Working_SetX_Path:" & Now)


        '-------------------
        '   依設定的工作檔組別
        '-------------------

        ' If CMB_SaveType.Text = "" Then
        '-------------------
        '   以後改為Set_0
        '-------------------
        'Form_Spec.Get_filePath_x_Append = Form_Spec.Get_file_Path_x
        ' Else
        '-------------------
        '   以後改為Set_1 ~ Set_xx
        '-------------------
        If SetX_Str <> "" Then
            Debug.Print("1：" & Get_file_Path_x())
            Debug.Print(Get_file_Path_x)

            Set_filePath_x_Append(Get_file_Path_x() & "\" & SetX_Str)
            Set_WorkPath_x_Append(Get_work_Path_x() & "\" & SetX_Str)
            Set_ExportPath_x_Append(Get_Export_Path_x() & "\" & SetX_Str)
            'Form_Spec.Get_Full_PathSmall_Name = Form_Spec.Get_file_Path_x & "\" & CMB_SaveType.Text
            'Form_Spec.Get_CHK_PicPath = Form_Spec.Get_file_Path_x & "\" & CMB_SaveType.Text
            'Form_Spec.FullName_ALL = Form_Spec.Get_file_Path_x & "\" & CMB_SaveType.Text
            '  End If
            Debug.Print("2:" & Get_filePath_x_Append())

            CHK_DIR_if_None_then_Create(Get_filePath_x_Append)

            Debug.Print(Get_filePath_x_Append() & "\LOG")

            CHK_DIR_if_None_then_Create(Get_filePath_x_Append() & "\LOG")


            FullName_Index_Path = Get_filePath_x_Append() & "\Index"
            FullName_item_Path = Get_filePath_x_Append() & "\DATA_CSV"

            Set_Full_PathSmall_Name(Get_WorkPath_x_Append() & "\sp")

            Set_CHK_PicPath(Get_WorkPath_x_Append() & "\Normal")

            FullName_ALL = Get_filePath_x_Append() & "\ALL"

            'temp_index_dir = Form_Spec.Get_filePath_x_Append & "\Index"

            If tCreat_DIR = True Then
                CHK_DIR_if_None_then_Create(FullName_item_Path)
                CHK_DIR_if_None_then_Create(Get_Full_PathSmall_Name)
                CHK_DIR_if_None_then_Create(Get_CHK_PicPath)
                CHK_DIR_if_None_then_Create(FullName_ALL)
                CHK_DIR_if_None_then_Create(FullName_Index_Path)
            End If


            If CKB_compare.Checked = True Then
                'Dim tempX As String

                FullName_Index_Path &= "\compare"
                FullName_item_Path &= "\compare"

                'Get_Full_PathSmall_Name &= "\compare"


                FullName_ALL &= "\compare"
            End If


            If tCreat_DIR = True Then
                Debug.Print("Idx_path:" & FullName_Index_Path)
                CHK_DIR_if_None_then_Create(FullName_Index_Path)
                CHK_DIR_if_None_then_Create(FullName_item_Path)
                CHK_DIR_if_None_then_Create(Get_Full_PathSmall_Name)
                CHK_DIR_if_None_then_Create(Get_CHK_PicPath)
                CHK_DIR_if_None_then_Create(FullName_ALL)

            End If
        End If
    End Sub
    Public Sub CHK_DIR_if_None_then_Create(ByVal AAA As String)
        If Not IO.Directory.Exists(AAA) Then

            '如不存在，建立資料夾
            IO.Directory.CreateDirectory(AAA)
        End If
    End Sub
    'Debug.Print("NFS_4")
    Public Sub Update_for_Sync()
        'Dim FullName_item_CSV As String
        'Dim Get_file_Path_x As String
        Dim xpage As Integer
        ' Dim xpage2_Offset As Integer
        Dim page_Main_Offset As Integer
        Dim xPage_Offset As Integer



        xpage = Get_Page_Now()

        FullName_item_CSV = Get_CSV_Filename()

        '---------------
        '   讀取工作檔
        '---------------


        TB_Data_CSV.Text = FullName_item_CSV

        Form_Spec_CHK_Ctl.LB_PageNow.Text = "Page_" & xpage.ToString("000")

        'PictureBox1.Image. = Get_file_Path_x & "\" & sss

        '-----------------------
        '   Bottom 讀取工作圖檔  [20201218]
        '-----------------------
        page_Main_Offset = Spec_Main_Offset(xpage)
        'Form_Spec_CHK_Bottom.PicBox_Main.Load(Get_file_Path_x() & "\Page_" & page_Main_Offset.ToString("000") & ".png")

        'Get_Full_Name_Ori_Pic()
        Form_Spec_CHK_Bottom.PicBox_Main.Load(Get_Full_Name_Ori_Pic(page_Main_Offset))



        If CKB_compare.Checked = True Then
            '-----------------------
            '   讀取比較圖檔  [202012]
            '-----------------------
            'Dim xxyy As Point
            'xxyy.X = Form_Spec_CHK_Top.Picbox_Gap + (Form_Spec_CHK_Bottom.PB2_ori_x * Form_Spec_CHK_Top.S_W_Ratio)
            'xxyy.Y = Form_Spec_CHK_Bottom.PB2_ori_y
            Form_Spec_CHK_Bottom.PicBox_Comp.Location = Get_Compare_Location()        '計算新的起點位置

            Form_Spec_CHK_Bottom.PicBox_Comp.Width = Form_Spec_CHK_Top.P1_Ori_Width * Form_Spec_CHK_Top.S_W_Ratio
            Form_Spec_CHK_Bottom.PicBox_Comp.Height = Form_Spec_CHK_Top.P1_Ori_Height * Form_Spec_CHK_Top.S_H_Ratio

            '-----------------------
            '   補差頁的處理
            '-----------------------
            xPage_Offset = Spec_Compare_OffsetPage(xpage)

            If xPage_Offset = 0 Then
                Form_Spec_CHK_Bottom.PicBox_Comp.Visible = False
            Else
                'Get_Full_Name_Compare_Pic()
                'If (File.Exists(Get_Compare_Path & "\Page_" & xPage_Offset.ToString("000") & ".png")) Then
                Dim temp_Offset_pic_filename As String
                temp_Offset_pic_filename = Get_Full_Name_Compare_Pic(xPage_Offset)

                If (File.Exists(temp_Offset_pic_filename)) Then
                    ' If  Then
                    Form_Spec_CHK_Bottom.PicBox_Comp.Visible = True
                    Form_Spec_CHK_Bottom.PicBox_Comp.Load(temp_Offset_pic_filename)
                Else
                    Form_Spec_CHK_Bottom.PicBox_Comp.Visible = False
                End If

            End If

        End If


        Form_Spec_CHK_Top.Graph_Clear()

        '---------------
        '   顯示原始規格圖 (Only)
        '---------------
        If Form_Spec_CHK_Ctl.CKB_Ori_Graphic.Checked = True Then

            Form_Spec_CHK_Top.PicBox_TopMain.Visible = True
            page_Main_Offset = Spec_Main_Offset(xpage)
            Form_Spec_CHK_Top.PicBox_TopMain.Load(Get_Full_Name_Ori_Pic(page_Main_Offset))

            If CKB_compare.Checked = True Then

                Form_Spec_CHK_Bottom.PicBox_Comp.Location = Get_Compare_Location()

                Form_Spec_CHK_Bottom.PicBox_Comp.Width = Form_Spec_CHK_Top.P1_Ori_Width * Form_Spec_CHK_Top.S_W_Ratio
                Form_Spec_CHK_Bottom.PicBox_Comp.Height = Form_Spec_CHK_Top.P1_Ori_Height * Form_Spec_CHK_Top.S_H_Ratio

                xPage_Offset = Spec_Compare_OffsetPage(xpage)

                If xPage_Offset = 0 Then
                    Form_Spec_CHK_Bottom.PicBox_Comp.Visible = False
                Else

                    Dim temp_Offset_pic_filename As String
                    temp_Offset_pic_filename = Get_Full_Name_Compare_Pic(xPage_Offset)


                    If (File.Exists(temp_Offset_pic_filename)) Then
                        Form_Spec_CHK_Top.PicBox_TopSub.Visible = True
                        Form_Spec_CHK_Top.PicBox_TopSub.Load(temp_Offset_pic_filename)
                        Form_Spec_CHK_Top.PicBox_TopSub.Location = Form_Spec_CHK_Bottom.PicBox_Comp.Location
                    Else
                        Form_Spec_CHK_Top.PicBox_TopSub.Visible = False
                    End If

                End If


            End If

        Else

            Form_Spec_CHK_Top.PicBox_TopMain.Visible = False

            If CKB_compare.Checked = True Then
                Form_Spec_CHK_Top.PicBox_TopSub.Visible = False
            End If

        End If




        '--------------------------------
        '   讀取設定檔 並進行方塊的繪製
        '--------------------------------
        Form_Spec_CHK_Ctl.CheckItem_Read_and_Draw(Form_Spec_CHK_Ctl.Get_CSV_Pattern(), False)


        '=====================
        '   輸出的處理
        '=====================

        '----------------------
        '   產生縮圖
        '----------------------
        If CKB_Create_Spic.Checked = True Then

            '---------------
            '   代換縮圖
            '---------------
            If Form_Spec_All_Pc.Created = True Then
                Form_Spec_All_Pc.Use_temp_Small_Img(xpage)    '//建立中使用暫存檔
                Form_Spec_All_Pc.Set_Arror(xpage)

                If CKB_USE_WorkPIC.Checked = True Then
                    Use_temp_Work_Img()
                End If

            End If

            If Spec_Diff.Created = True Then
                Spec_Diff.Set_Arror(xpage)
            End If

            Create_Small_Picture()                      '//產生縮圖

            '---------------
            '   還原縮圖
            '---------------
            If Form_Spec_All_Pc.Created = True Then
                wait_ms(500)
                Form_Spec_All_Pc.Use_Ori_Small_Img(xpage)     '//使用目標縮圖
                Use_Ori_Work_Img()
            End If


        Else
            '---------------------
            '   只調整游標
            '---------------------
            If Form_Spec_All_Pc.Created = True Then
                Form_Spec_All_Pc.Set_Arror(xpage)
            End If

            If Spec_Diff.Created = True Then
                Spec_Diff.Set_Arror(xpage)
            End If

        End If







    End Sub
    Public Function Spec_Main_Offset(ByVal NowPage As Integer) As Integer
        Dim temp_Offset As Integer = 0
        temp_Offset = 0

        'If (NowPage > 87) Then
        '    temp_Offset += 1
        'End If
        'If (NowPage > 88) Then
        '    temp_Offset += 1
        'End If
        'If (NowPage > 76) Then
        '    temp_Offset += 1
        'End If

        Return NowPage - temp_Offset

    End Function
    '====================================
    '   這裡是要補償 舊仕樣沒有的頁面
    '   例如舊的只有95頁, 新的有100頁
    '   這時候舊的需要補5頁的空白
    '   ex空白頁在新仕樣的50,60,70,80,90頁
    '   如果這時候比對設定在62頁時, 因為大於50, 60
    '   所以取得的頁面就要減去2 Page, 故傳回舊仕樣的60頁即可
    '====================================
    Public Function Spec_Compare_OffsetPage(ByVal NowPage As Integer) As Integer
        Dim temp_Offset_Counter As Integer = 0
        temp_Offset_Counter = 0

        'Dim temp_Result As Integer = 0


        For Each Xval As Integer In Form_Spec_CHK_Ctl.Compare_Page_List.Items()
            'Xval()
            If NowPage = Xval Then
                Return 0
            ElseIf NowPage > Xval Then
                temp_Offset_Counter += 1
            Else

            End If
            'Debug.Print("Listbox_Add:" & Xval)
        Next




        Return NowPage - temp_Offset_Counter

    End Function
    Public Sub Use_temp_Work_Img()
        PicBox_Work.Load(Get_Source_DIR() & "\Spec\TempIMG2.png")

    End Sub
    Public Sub Use_Ori_Work_Img()



        If CKB_USE_WorkPIC.Checked = True Then

            If (File.Exists(Get_CHK_PicName)) Then
                PicBox_Work.Load(Get_CHK_PicName)
            Else
                Use_temp_Work_Img()
            End If

        Else
            PicBox_Work.Load(Get_Full_Name_Ori_Pic(0))
        End If


    End Sub

    Public Sub Update_for_Sync_Create_SmallPic()
        'Dim FullName_item_CSV As String
        'Dim Get_file_Path_x As String
        Dim xpage As Integer



        xpage = Get_Page_Now()



        FullName_item_CSV = Get_CSV_Filename()

        TB_Data_CSV.Text = FullName_item_CSV

        Form_Spec_CHK_Ctl.LB_PageNow.Text = "Page_" & xpage.ToString("000")


        Form_Spec_CHK_Bottom.PicBox_Main.Load(Get_Full_Name_Ori_Pic(0))

        If CKB_compare.Checked = True Then

            Dim xPage_Offset As Integer



            Form_Spec_CHK_Bottom.PicBox_Comp.Location = Get_Compare_Location()

            Form_Spec_CHK_Bottom.PicBox_Comp.Width = Form_Spec_CHK_Top.P1_Ori_Width * Form_Spec_CHK_Top.S_W_Ratio
            Form_Spec_CHK_Bottom.PicBox_Comp.Height = Form_Spec_CHK_Top.P1_Ori_Height * Form_Spec_CHK_Top.S_H_Ratio

            xPage_Offset = Spec_Compare_OffsetPage(xpage)

            If xPage_Offset = 0 Then
                '----------------------------------
                '   比對的舊Spec沒有這個頁面, 不顯示
                '----------------------------------
                Form_Spec_CHK_Bottom.PicBox_Comp.Visible = False
            Else
                '----------------------------------
                '   取得Offset後的頁面
                '----------------------------------
                Dim temp_Comp_png As String

                temp_Comp_png = Get_Full_Name_Compare_Pic(xPage_Offset)


                If (File.Exists(temp_Comp_png)) Then
                    Form_Spec_CHK_Bottom.PicBox_Comp.Visible = True
                    Form_Spec_CHK_Bottom.PicBox_Comp.Load(temp_Comp_png)
                Else
                    '----------------------------------
                    '   比對的舊Spec沒有這個頁面, 不顯示
                    '----------------------------------
                    Form_Spec_CHK_Bottom.PicBox_Comp.Visible = False
                End If

            End If

        End If


        Form_Spec_CHK_Top.Graph_Clear()


        If Form_Spec_CHK_Ctl.CKB_Ori_Graphic.Checked = True Then
            Form_Spec_CHK_Top.PicBox_TopMain.Visible = True
            Form_Spec_CHK_Top.PicBox_TopMain.Load(Get_Full_Name_Ori_Pic(0))
        Else
            Form_Spec_CHK_Top.PicBox_TopMain.Visible = False
        End If


        '------------------------
        '   讀取設定檔
        '------------------------

        Form_Spec_CHK_Ctl.CheckItem_Read_and_Draw(Form_Spec_CHK_Ctl.Get_CSV_Pattern(), False)


        '----------------------
        '   產生縮圖
        '----------------------
        'If CKB_Create_Spic.Checked = True Then



        If Form_Spec_All_Pc.Created = True Then
            Form_Spec_All_Pc.Use_temp_Small_Img(xpage)
            Form_Spec_All_Pc.Set_Arror(xpage)
        End If

        If Spec_Diff.Created = True Then
            Spec_Diff.Set_Arror(xpage)
        End If


        Create_Small_Picture()

        If Form_Spec_All_Pc.Created = True Then
            wait_ms(500)
            Form_Spec_All_Pc.Use_Ori_Small_Img(xpage)
        End If


        If Spec_Diff.Created = True Then
            Spec_Diff.Set_Arror(xpage)
        End If






    End Sub
    Public Function Get_Compare_Location() As Point
        Dim xxyy As Point
        Dim x_base As Single

        If CKB_Lo_dpi.Checked = True Then
            x_base = 0.5
        Else
            x_base = 1
        End If

        'xxyy.X = Form_Spec_CHK_Top.Picbox_Gap + (Form_Spec_CHK_Bottom.PB2_ori_x * Form_Spec_CHK_Top.S_W_Ratio * x_base)
        xxyy.X = Form_Spec_CHK_Top.Picbox_Gap + (Form_Spec_CHK_Bottom.PicBox_Main.Location.X + Form_Spec_CHK_Bottom.PicBox_Main.Width)
        'xxyy.Y = Form_Spec_CHK_Bottom.PB2_ori_y * Form_Spec_CHK_Top.S_H_Ratio
        xxyy.Y = Form_Spec_CHK_Bottom.PicBox_Main.Location.Y
        'Debug.Print("compare X:" & xxyy.X)
        Return xxyy
    End Function
    Public Sub Read_Page_Index_file(ByVal SourceFile As String, ByVal Dmsg As Boolean)       '[PH00]
        Dim Item_Counter As Integer = 0

        ' Form_PDF As Boolean = False

        Dim line_counter As Integer

        Dim Parameter_Line As String = ""

        ' Dim strArr2() As String

        CMB_Index.Items.Clear()

        line_counter = 0

        If (File.Exists(SourceFile)) Then


            Dim fs_Parameter As FileStream = New FileStream(SourceFile, FileMode.Open)
            Using sr_Parameter As StreamReader = New StreamReader(fs_Parameter, _
                                                        System.Text.Encoding.Default)


                Parameter_Line = sr_Parameter.ReadLine()                                  '讀取一行資料 略過標題

                Form_PDF = Parameter_Line.Contains("PDF")
                'If Parameter_Line.Contains("PDF") = True Then

                'End If
                'Form_PDF()

                Parameter_Line = sr_Parameter.ReadLine()                                  '讀取一行資料 略過標題
                Parameter_Line = sr_Parameter.ReadLine()                                  '讀取一行資料 略過標題

                While (sr_Parameter.EndOfStream <> True)





                    Parameter_Line = sr_Parameter.ReadLine()                                  '讀取一行資料 略過標題
                    CMB_Index.Items.Add(Parameter_Line)
                    'strArr2 = Parameter_Line.Split(",")







                End While


                sr_Parameter.Close()                                                          '關閉檔案



            End Using

        Else



            If Dmsg = True Then
                MsgBox("Please check: " & SourceFile & ", 此為Log檔案的路徑")
            End If

        End If



    End Sub
    Public Sub Read_Spec_List_file(ByVal SourceFile As String, ByVal Dmsg As Boolean)       '[PH00]
        Dim Item_Counter As Integer = 0

        ' Form_PDF As Boolean = False

        Dim line_counter As Integer

        Dim Parameter_Line As String = ""

        ' Dim strArr2() As String

        line_counter = 0

        If (File.Exists(SourceFile)) Then


            Dim fs_Parameter As FileStream = New FileStream(SourceFile, FileMode.Open)
            Using sr_Parameter As StreamReader = New StreamReader(fs_Parameter, _
                                                        System.Text.Encoding.Default)




                While (sr_Parameter.EndOfStream <> True)





                    Parameter_Line = sr_Parameter.ReadLine()                                  '讀取一行資料 略過標題
                    CMB_Spec_List.Items.Add(Parameter_Line)
                    CMB_Compare.Items.Add(Parameter_Line)
                    'strArr2 = Parameter_Line.Split(",")







                End While


                sr_Parameter.Close()                                                          '關閉檔案



            End Using

        Else



            If Dmsg = True Then
                MsgBox("Please check: " & SourceFile & ", 此為Log檔案的路徑")
            End If

        End If



    End Sub
    Public Sub Read_Spec_Server_file(ByVal SourceFile As String, ByVal Dmsg As Boolean)       '[PH00]
        Dim Item_Counter As Integer = 0

        Dim temp_array() As String

        ' Form_PDF As Boolean = False

        Dim line_counter As Integer

        Dim Parameter_Line As String = ""

        ' Dim strArr2() As String

        line_counter = 0

        If (File.Exists(SourceFile)) Then


            Dim fs_Parameter As FileStream = New FileStream(SourceFile, FileMode.Open)
            Using sr_Parameter As StreamReader = New StreamReader(fs_Parameter, _
                                                        System.Text.Encoding.Default)




                'While (sr_Parameter.EndOfStream <> True)





                Parameter_Line = sr_Parameter.ReadLine()                                  '讀取一行資料 略過標題
                temp_array = Parameter_Line.Split(",")
                Redmine_Server = temp_array(1)

                Parameter_Line = sr_Parameter.ReadLine()                                  '讀取一行資料 略過標題
                temp_array = Parameter_Line.Split(",")
                Redmine_Server2 = temp_array(1)

                Parameter_Line = sr_Parameter.ReadLine()                                  '讀取一行資料 略過標題
                temp_array = Parameter_Line.Split(",")
                Mantis_Server = temp_array(1)

                Parameter_Line = sr_Parameter.ReadLine()                                  '讀取一行資料 略過標題
                temp_array = Parameter_Line.Split(",")
                Mantis_Server2 = temp_array(1)

                Parameter_Line = sr_Parameter.ReadLine()                                  '讀取一行資料 略過標題
                temp_array = Parameter_Line.Split(",")
                NotePad_path = temp_array(1)








                ' End While


                sr_Parameter.Close()                                                          '關閉檔案



            End Using

        Else



            If Dmsg = True Then
                MsgBox("Please check: " & SourceFile & ", 此為Log檔案的路徑")
            End If

        End If



    End Sub
    Private Sub TrackBar1_ValueChanged(sender As Object, e As EventArgs) Handles TrackBar1.ValueChanged

        Change_Page(TrackBar1.Value)

    End Sub
    Public Sub Change_Page(ByVal NewPAGE As Integer)

        '----------------
        '   保留上回頁面
        '----------------

        If NewPAGE <> Get_Page_Now() Then
            Form_Spec_CHK_Ctl.LB_SelectedItem.Text = ""

            Form_Spec_CHK_Ctl.TB_LastPage.Text = Get_Page_Now()


            TrackBar1.Value = NewPAGE
            Set_Page_Now(NewPAGE)
            'Debug.Print("Update_pic_From_Change_Page")
            Update_pic()

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Page_Dn()
    End Sub
    Public Sub Page_Dn()
        If TrackBar1.Value > TrackBar1.Minimum Then
            TrackBar1.Value -= 1
            Change_Page(TrackBar1.Value)
        End If


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Page_Up()

    End Sub
    Public Sub Page_Up()
        If TrackBar1.Value < TrackBar1.Maximum Then
            TrackBar1.Value += 1
            Change_Page(TrackBar1.Value)
        End If

    End Sub
    Public Sub Page_Now()
        'If TrackBar1.Value < TrackBar1.Maximum Then
        'TrackBar1.Value += 1

        Change_Page(TrackBar1.Value)
        'End If
        Debug.Print("Update_pic_From_Page_Now")
        Update_pic()
    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click


        Form_Spec_CHK_Top.Show()
        Form_Spec_CHK_Bottom.Show()
        Form_Spec_CHK_Bottom.Location = Form_Spec_CHK_Top.Location

        Form_Spec_CHK_Bottom.TopMost = True
        Form_Spec_CHK_Top.TopMost = True
        Form_Spec_CHK_Bottom.TopMost = False
        Form_Spec_CHK_Top.TopMost = False

        Update_pic()
    End Sub

    'Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
    '    'Update_for_Sync()
    'End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)
        Form_Spec_CHK_Ctl.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        adj_Opacity()

    End Sub
    Public Sub adj_Opacity()
        Dim xxx As Single
        xxx = Val(TrackBar2.Value) * 0.01

        Label3.Text = TrackBar2.Value & "%"

        Form_Spec_CHK_Top.Opacity = xxx
    End Sub

    Private Sub TrackBar2_Scroll(sender As Object, e As EventArgs) Handles TrackBar2.Scroll

    End Sub

    Private Sub TrackBar2_ValueChanged(sender As Object, e As EventArgs) Handles TrackBar2.ValueChanged
        adj_Opacity()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim Full_target_name As String

        Full_target_name = TB_Data_CSV.Text

        '===================
        '　開啟Report
        '===================

        If (File.Exists(Full_target_name)) Then
            '        ReportFileName
            If (Full_target_name <> "") Then
                System.Diagnostics.Process.Start(Full_target_name)
            Else
                MsgBox("No (Report) File Name")
            End If
        Else
            MsgBox("找不到該檔案:" & Full_target_name)
        End If
    End Sub


    Public Sub Create_Small_Picture()   '[CP00]

        If CKB_compare.Checked = False Then

            Form_Spec_CHK_Top.Capture_ReSize_New_save(100, 160, Get_CHK_PicName, Get_Full_Small_Name)
        Else

            Form_Spec_CHK_Top.Capture_ReSize_New_save(220, 160, Get_CHK_PicName, Get_Full_Small_Name)
        End If

    End Sub
    Public Sub Create_Work_Picture()     '[CP01]

        Form_Spec_CHK_Top.Capture_SavePNG_WorkPic(Get_CHK_PicName)


    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs)

        Scan_All_Page()

    End Sub
    Private Sub Scan_All_Page()
        For i = 1 To Total_Pages
            TrackBar1.Value = i
            Set_Page_Now(i)

            'Do While Switch_Page_Done = False
            '    Application.DoEvents()
            'Loop

            wait_ms(200)  '一定要等待畫面繪製好, 才可以擷取

            Update_pic()



        Next
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs)
        'CKB_Create_Spic.Checked = False

        'Form_Spec_All.Show()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        CKB_Create_Spic.Checked = False

        Form_Spec_All_Pc.Show()
    End Sub
    Public Sub Check_Ver_Path(ByVal AAA As String)
        If (Get_SpecPath_file_lines(AAA) = 5) Then
            Set_Spec_Path_file_Version(1)

            Form_Move.LB_Old_Flag.BackColor = Color.YellowGreen
            Form_Move.LB_New_Flag.BackColor = Color.Black

        Else
            Set_Spec_Path_file_Version(2)
            Form_Move.LB_Old_Flag.BackColor = Color.Black
            Form_Move.LB_New_Flag.BackColor = Color.YellowGreen

        End If
    End Sub
    Private Sub Form_Spec_Load(sender As Object, e As EventArgs) Handles Me.Load
        '[SCHK000]

        Dim xx_index_path As String
        Dim tp_f_path As String

        ComboBox3.SelectedIndex = 0

        'Spec_Diff.Set_ALLPic_Type(True)


        Dim filename1 As String


        Dim xx_input() As String



        Name_Default = "ECU_VV_20200819_t3"
        Page_Default = "120"


        'Get_Spec_path()
        filename1 = Get_SpecCHK_PathFile()

        If (File.Exists(filename1)) Then
            ' MsgBox("找到該路徑:" & filename1)

            '-----------------------------------
            '   先確認 Spec_Path.txt 版本   舊版5行 / 新版9行
            '-----------------------------------
            Check_Ver_Path(filename1)
            'If (Get_SpecPath_file_lines(filename1) = 5) Then
            '    Set_Spec_Path_file_Version(1)

            '    Form_Move.LB_Old_Flag.BackColor = Color.YellowGreen
            '    Form_Move.LB_New_Flag.BackColor = Color.Black

            'Else
            '    Set_Spec_Path_file_Version(2)
            '    Form_Move.LB_Old_Flag.BackColor = Color.Black
            '    Form_Move.LB_New_Flag.BackColor = Color.YellowGreen

            'End If

            Form_Move.LB_PathFile_Version.Text = Get_Spec_Path_file_Version()


            Get_Spec_path(filename1)

            Check_Ver_Path(filename1)

            '-------------------------------------------
            '   自動　轉換到　新版本
            '-------------------------------------------
            'If Get_Spec_Path_file_Version() = 1 Then

            '    Dim response

            '    response = MsgBox("[msg001]此為舊格式,是否轉換至V2.0", vbYesNo)

            '    If response = vbYes Then
            '        Tarns_to_Ver2()      '執行Function

            '    Else

            '    End If
            'End If

            Get_Spec_path(filename1)

        Else
            MsgBox("找不到路徑檔:" & filename1)
            MsgBox("ERR_0001")
            MsgBox("直接結束")
            End
            'Create_New_file()
        End If





        '------------------
        '   預設工作資料夾
        '------------------

        'Spec_Name = "ECU_VV_20200819_t3"
        'TB_Page_Total.Text = "120"

        Spec_Name = Name_Default
        TB_Page_Total.Text = Page_Default


        Set_file_Path_x(Get_Source_DIR() & "\Spec\" & Spec_Name)    '20210526
        Set_work_Path_x(Get_work_DIR() & "\Spec\" & Spec_Name)    '20210526


        '------------------
        '   外部CMD - 代入變數
        '------------------
        If Command() <> "" Then
            xx_input = Command().Split(" ")
            Spec_Name = xx_input(0)
            TB_Page_Total.Text = xx_input(1)
        End If




        Me.Text = "SPCE CHK   " & Spec_Name & "   ( " & TB_Page_Total.Text & " )   " & Title_String

        '------------------
        '   檢查路徑
        '------------------
        tp_f_path = Get_Source_DIR() & "\Spec\" & Spec_Name
        If Directory.Exists(tp_f_path) = False Then
            MsgBox("找不到該路徑:" & tp_f_path)
            'Return
        End If



        Total_Pages = Val(TB_Page_Total.Text)
        TrackBar1.Maximum = Total_Pages

        Page_Dn()

        '---------------------
        '   建立工作資料夾
        '---------------------

        check_and_Create_DIR_ALL()



        '---------------------
        '   讀取索引檔
        '---------------------
        xx_index_path = Get_Source_DIR() & "\Spec\" & Spec_Name & "\Set0\index\index.txt"

        If File.Exists(xx_index_path) = False Then
            MsgBox("找不到該index檔:" & xx_index_path)
            'Return
        Else
            Spec_index_path = xx_index_path
            Read_Page_Index_file(Spec_index_path, True)
        End If




        '---------------------
        '   讀取規格List檔
        '---------------------
        xx_index_path = Get_Source_DIR() & "\APP\Spec_List.txt"

        If File.Exists(xx_index_path) = False Then
            MsgBox("找不到該Spec_List檔:" & xx_index_path)
            'Return
        Else
            Read_Spec_List_file(xx_index_path, True)


        End If

        '---------------------
        '   讀取Server檔
        '---------------------
        xx_index_path = Get_Source_DIR() & "\APP\Spec_Server.txt"

        If File.Exists(xx_index_path) = False Then
            MsgBox("找不到該Spec_Server檔:" & xx_index_path)
            'Return
        Else
            Read_Spec_Server_file(xx_index_path, True)


        End If





    End Sub
    Public Sub Set_Spec_Path_file_Version(ByVal tValue As Integer)
        Spec_Path_x_file_Version = tValue
    End Sub
    Public Function Get_Spec_Path_file_Version() As Integer
        Return Spec_Path_x_file_Version
    End Function
    Public Sub Tarns_to_Ver2()      '[NEW01]
        Dim Parameter_Line As String = ""
        Dim Parameter_Line1 As String = ""
        Dim Parameter_Line2 As String = ""

        Dim src_file As String = ""
        Dim temp_file As String = "op.txt"
        Debug.Print("Tarns_to_Ver2")
        Debug.Print(Get_SpecCHK_PathFile())


        src_file = Get_SpecCHK_PathFile()

        File.Copy(src_file, temp_file, True)


        Dim fs As FileStream = New FileStream(temp_file, FileMode.Open)
        Using sr As StreamReader = New StreamReader(fs, _
                                                    System.Text.Encoding.Default)


            '讀取一行資料 略過標題



            'While (sr.EndOfStream <> True)



            '送出相關設定
            Parameter_Line1 = sr.ReadLine()                                     '讀取一行資料


            ww_back(src_file, Parameter_Line1, True)
            'End While


            Parameter_Line2 = sr.ReadLine()                                     '讀取一行資料
            ww_back(src_file, Parameter_Line2, False)



            ww_back(src_file, Parameter_Line1 & "_work", False)
            ww_back(src_file, Parameter_Line2 & "_work", False)
            ww_back(src_file, Parameter_Line1 & "_Export", False)
            ww_back(src_file, Parameter_Line2 & "_Export", False)

            Parameter_Line = sr.ReadLine()                                     '讀取一行資料
            ww_back(src_file, Parameter_Line, False)

            Parameter_Line = sr.ReadLine()                                     '讀取一行資料
            ww_back(src_file, Parameter_Line, False)

            Parameter_Line = sr.ReadLine()                                     '讀取一行資料
            ww_back(src_file, Parameter_Line, False)

            sr.Close()                                                          '關閉檔案




        End Using



        MsgBox("Done.")
        '-------------------------
        '   舊的
        '-------------------------
        'd:\ECU_TEST
        'c:\Joe\ECU_TEST

        'C:\Program Files (x86)\Acute\LA\LA.exe
        'ECU_UN_20210429_t11()
        '136:

        '-------------------------
        '   新的
        '-------------------------
        'd:\ECU_TEST
        'c:\Joe\ECU_TEST

        'd:\ECU_TEST_work                       <= 新增
        'c:\Joe\ECU_TEST_work                   <= 新增
        'd:\ECU_TEST_Export                     <= 新增
        'c:\Joe\ECU_TEST_Export                 <= 新增

        'C:\Program Files (x86)\Acute\LA\LA.exe
        'ECU_UN_20210429_t11()
        '136:



        '---------------
        '   寫回檔案
        '---------------

    End Sub
    Public Sub ww_back(ByVal tFilename As String, ByVal AAA As String, ByVal tType As Boolean)
        'Dim filename1 As String
        Dim filenum As Integer


        filenum = FreeFile()
        If tType = True Then
            FileOpen(filenum, tFilename, OpenMode.Output)
        Else
            FileOpen(filenum, tFilename, OpenMode.Append)
        End If




        PrintLine(filenum, AAA)


        FileClose(filenum)
    End Sub
    Private Sub check_and_Create_DIR_ALL()
        check_and_Create_DIR2("Index")
        'Debug.Print("")
        check_and_Create_DIR2("All")

        check_and_Create_DIR2B("Normal")
        check_and_Create_DIR2B("sp")

        check_and_Create_DIR2("DATA_CSV")

        'check_and_Create_DIR2("Index")
        check_and_Create_DIR2("All\compare")

        check_and_Create_DIR2B("Normal\compare")
        check_and_Create_DIR2B("sp\compare")

        check_and_Create_DIR2("DATA_CSV\compare")


    End Sub
    Private Sub check_and_Create_DIR2(ByVal SName As String)
        Dim Append_Path As String

        Append_Path = Get_Source_DIR() & "\Spec\" & Spec_Name & "\Set0\" & SName
        If Not IO.Directory.Exists(Append_Path) Then

            '如不存在，建立資料夾
            IO.Directory.CreateDirectory(Append_Path)
        End If
    End Sub

    Private Sub check_and_Create_DIR2B(ByVal SName As String)
        Dim Append_Path As String

        Append_Path = Get_work_DIR() & "\Spec\" & Spec_Name & "\Set0\" & SName
        Debug.Print(Append_Path)
        '-------------------------------
        '但是如果沒有D槽時, 就會發生問題
        '-------------------------------
        Dim temp_op As String = ""
        temp_op = UCase(Append_Path)
        If temp_op.Contains("D:\") Then
            If Not IO.Directory.Exists("d:\") Then
                MsgBox("[MSG210626]_沒有D槽,無法建立:" & Append_Path)
            Else
                If Not IO.Directory.Exists(Append_Path) Then

                    '如不存在，建立資料夾
                    IO.Directory.CreateDirectory(Append_Path)
                End If
            End If
        End If


    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dim Full_target_name As String

        Full_target_name = TB_FName.Text

        '===================
        '　開啟Report
        '===================

        If (File.Exists(Full_target_name)) Then
            '        ReportFileName
            If (Full_target_name <> "") Then
                System.Diagnostics.Process.Start(Full_target_name)
            Else
                MsgBox("No (Report) File Name")
            End If
        Else
            MsgBox("找不到該檔案:" & Full_target_name)
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim Full_target_name As String

        Full_target_name = TB_CHK_PicName.Text

        '===================
        '　開啟Report
        '===================

        If (File.Exists(Full_target_name)) Then
            '        ReportFileName
            If (Full_target_name <> "") Then
                System.Diagnostics.Process.Start(Full_target_name)
            Else
                MsgBox("No (Report) File Name")
            End If
        Else
            MsgBox("找不到該檔案:" & Full_target_name)
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CMB_Index.SelectedIndexChanged
        Dim xx() As String
        Dim array_len As Integer

        If Form_PDF = True Then
            xx = CMB_Index.SelectedItem.ToString.Split(";")
            array_len = UBound(xx)
            Change_Page(Val(xx(array_len)))
        Else
            xx = CMB_Index.SelectedItem.ToString.Split(",")
            Change_Page(Val(xx(0)))
        End If

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Time_UP = True
        Timer1.Enabled = False

    End Sub
    Public Sub wait_ms(ByVal t_Data As Integer)
        Timer1.Interval = t_Data
        Timer1.Enabled = True

        Time_UP = False
        Timer1.Interval = t_Data
        Timer1.Enabled = True
        While (Time_UP = False)
            Application.DoEvents()
        End While





    End Sub
    '====================
    '   [SCHK002] 讀取設定檔 
    '====================
    Public Sub Get_Spec_path(ByVal Flie_Name_ As String)
        'Dim temp_DIR1 As String
        'Dim temp_DIR2 As String
        Debug.Print("[DEG000] FileName: " & Flie_Name_)
        Form_Move.LB_Setup_Filename.Text = Flie_Name_







        Dim fs As FileStream = New FileStream(Flie_Name_, FileMode.Open)
        Using sr As StreamReader = New StreamReader(fs, _
                                                    System.Text.Encoding.Default)


            'Get_Source_DIR = sr.ReadLine() & vbCrLf                                     '讀取一行資料 略過標題
            '=================
            '   讀取工作路徑
            '=================
            Main_Path = sr.ReadLine()
            Sub_Path = sr.ReadLine()



            Debug.Print("Read_Main_Path: " & Main_Path)
            Debug.Print("Read_Sub_Path: " & Sub_Path)




            If Get_Spec_Path_file_Version() = 2 Then
                '------------------------
                '   新版本才有使用
                '------------------------
                Debug.Print("is Ver2")
                Main_workPath = sr.ReadLine()
                Sub_workPath = sr.ReadLine()
                Debug.Print("Ver2_Main_workPath:" & Main_workPath)
                Debug.Print("Ver2_Sub_workPath:" & Sub_workPath)
                'Debug.Print()
                Main_ExportPath = sr.ReadLine()
                Sub_ExportPath = sr.ReadLine()
                Debug.Print("Ver2_Main_ExportPath:" & Main_ExportPath)
                Debug.Print("Ver2_Sub_ExportPath:" & Sub_ExportPath)
            End If




            If Directory.Exists(Main_Path) = True Then
                '------------------
                '   [主要路徑]
                '------------------
                Set_Source_DIR(Main_Path)
                Use_Main_path = True

            Else
                '------------------
                '   檢查 [次要路徑]
                '------------------


                If Directory.Exists(Sub_Path) = True Then
                    Set_Source_DIR(Sub_Path)
                    Use_Sub_path = True
                Else
                    Debug.Print("Why?:" & Sub_Path)
                    '-------------------
                    '   主要／次要，　都不存在
                    '-------------------
                    MsgBox("[MSG000] it's not Exist, Please CHK Path:" & Main_Path & " & " & Sub_Path)
                End If
            End If




            If Get_Spec_Path_file_Version() = 2 Then

                Debug.Print("is Ver2")
                '------------------------
                '   新版本才有
                '------------------------

                If Use_Main_path = True Then
                    If Directory.Exists(Main_workPath) = False Then
                        IO.Directory.CreateDirectory(Main_workPath)
                    End If
                End If

                If Directory.Exists(Main_workPath) = True Then
                    '------------------
                    '   主要工作路徑
                    '------------------
                    Set_work_DIR(Main_workPath)


                Else
                    '------------------
                    '   檢查 次要工作路徑
                    '------------------
                    Debug.Print("Why?:" & Sub_workPath)
                    If Directory.Exists(Sub_workPath) = True Then
                        Set_work_DIR(Sub_workPath)
                    Else

                        '------------------
                        ' 連次要工作路徑都沒有時
                        '   是否應該要建立?
                        '------------------
                        If Use_Sub_path = True Then
                            IO.Directory.CreateDirectory(Sub_workPath)
                        End If

                        MsgBox("[MSG001] it's not Exist,Please CHK Path(New Ver):" & Main_Path & " & " & Sub_Path)



                    End If
                End If



                If Directory.Exists(Main_ExportPath) = False Then
                    '------------------
                    '   檢查 次要 [匯出路徑]
                    '------------------
                    If Directory.Exists(Sub_ExportPath) = True Then
                        Set_Export_DIR(Sub_ExportPath)
                    Else
                        IO.Directory.CreateDirectory(Sub_ExportPath)
                        Set_Export_DIR(Sub_ExportPath)
                        'MsgBox("Please CHK Path:" & Main_Path & " & " & Sub_Path)
                    End If
                Else
                    '------------------
                    '   主要 [匯出路徑]
                    '------------------
                    Set_Export_DIR(Main_ExportPath)
                End If
            End If



            LA_DIR = sr.ReadLine()

            Name_Default = sr.ReadLine()
            Page_Default = sr.ReadLine()


            Debug.Print("Name_Default" & Name_Default)
            Debug.Print("Page_Default:" & Page_Default)



            'Label134.Text = "END"


            sr.Close()                                                          '關閉檔案




        End Using


        '-------------------------------------------
        '   自動　轉換到　新版本
        '-------------------------------------------
        If Get_Spec_Path_file_Version() = 1 Then

            Dim response

            response = MsgBox("[msg001]此為舊格式,是否轉換至V2.0", vbYesNo)

            If response = vbYes Then
                Tarns_to_Ver2()      '執行Function

            Else

            End If
        End If

       


    End Sub
    '====================
    '   [SCHK001_] 讀取設定檔 行數
    '====================
    Public Function Get_SpecPath_file_lines(ByVal Flie_Name_ As String) As Integer
        'Dim temp_DIR1 As String
        'Dim temp_DIR2 As String
        Dim line_Counter As Integer = 0
        Dim Parameter_Line As String

        Dim fs As FileStream = New FileStream(Flie_Name_, FileMode.Open)
        Using sr As StreamReader = New StreamReader(fs, _
                                                    System.Text.Encoding.Default)


            While (sr.EndOfStream <> True)



                '送出相關設定
                Parameter_Line = sr.ReadLine() & vbCrLf                                     '讀取一行資料

                line_Counter += 1




            End While



            sr.Close()                                                          '關閉檔案




        End Using
        Return line_Counter
    End Function

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PicBox_Work.Click

    End Sub

    Private Sub TB_Page_Now_TextChanged(sender As Object, e As EventArgs) Handles TB_Page_Now.TextChanged

    End Sub

    Private Sub CO2019120975ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CO2019120975ToolStripMenuItem.Click
        Switch_Spec("ECU_CO__20191209_7_5", "109")
    End Sub
    Private Sub Switch_Spec(ByVal AA As String, ByVal BB As String)
        Dim xx_index_path As String
        Dim f_path As String

        '------------------
        '   機種切換
        '------------------
        ' If Command() <> "" Then
        'xx_input = Command().Split(" ")

        '-------------------------
        '   保留現在的 Spec,Page
        '-------------------------
        LB_LastSpec.Text = Spec_Name & "," & TB_Page_Total.Text & "," & Get_Page_Now()

        Spec_Name = AA
        TB_Page_Total.Text = BB
        'End If




        Me.Text = "SPCE CHK   " & Spec_Name & "   ( " & TB_Page_Total.Text & " )   " & Title_String

        '------------------
        '   檢查路徑
        '------------------
        f_path = Get_Source_DIR() & "\Spec\" & Spec_Name

        If Directory.Exists(f_path) = False Then
            MsgBox("找不到該路徑:" & f_path)
            'Return
        End If



        Total_Pages = Val(TB_Page_Total.Text)
        TrackBar1.Maximum = Total_Pages

        Page_Dn()

        '---------------------
        '   建立工作資料夾
        '---------------------

        check_and_Create_DIR_ALL()



        '---------------------
        '   讀取索引檔
        '---------------------
        xx_index_path = Get_Source_DIR() & "\Spec\" & Spec_Name & "\index\index.txt"

        If File.Exists(xx_index_path) = False Then
            ' MsgBox("找不到該index檔:" & xx_index_path)
            CMB_Index.Items.Clear()
            Label11.Visible = True

            'Return
        Else
            Spec_index_path = xx_index_path
            Read_Page_Index_file(Spec_index_path, True)
            Label11.Visible = False
        End If


    End Sub
    Private Sub Switch_Spec_and_Page(ByVal AA As String, ByVal BB As String, ByVal Target_Page As String)
        Dim xx_index_path As String
        Dim temp_f_path As String

        '------------------
        '   機種切換
        '------------------
        ' If Command() <> "" Then
        'xx_input = Command().Split(" ")





        '------------------
        '   保留現在的 Spec & Page
        '------------------
        LB_LastSpec.Text = Spec_Name & "," & TB_Page_Total.Text & "," & Get_Page_Now()

        'End If


        '-----------------------------------
        '   顯示標題 From Title
        '-----------------------------------
        Spec_Name = AA
        TB_Page_Total.Text = BB

        Name_Selected = AA
        Page_Selected = BB

        Me.Text = "SPCE CHK   " & Spec_Name & "   ( " & TB_Page_Total.Text & " )   " & Title_String


        '------------------
        '   檢查路徑
        '------------------
        temp_f_path = Get_Source_DIR() & "\Spec\" & Spec_Name

        If Directory.Exists(temp_f_path) = False Then
            MsgBox("找不到該路徑:" & temp_f_path)
            'Return
        End If



        Total_Pages = Val(TB_Page_Total.Text)
        TrackBar1.Maximum = Total_Pages

        Page_Dn()

        '---------------------
        '   建立工作資料夾
        '---------------------

        check_and_Create_DIR_ALL()



        '---------------------
        '   讀取索引檔
        '---------------------
        'xx_index_path = Get_Source_DIR & "\Spec\" & Spec_Name & "\index\index.txt"
        xx_index_path = Get_Source_DIR() & "\Spec\" & Spec_Name & "\Set0\index\index.txt"   '20210526

        If File.Exists(xx_index_path) = False Then
            ' MsgBox("找不到該index檔:" & xx_index_path)
            CMB_Index.Items.Clear()
            Label11.Visible = True

            'Return
        Else
            Spec_index_path = xx_index_path
            Read_Page_Index_file(Spec_index_path, True)
            Label11.Visible = False
        End If

        Set_Page_Now(Target_Page)
        'TrackBar1.Value = Val(CC)
        Update_pic()
    End Sub


    Private Sub VV20200819t3ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VV20200819t3ToolStripMenuItem.Click
        Switch_Spec("ECU_VV_20200819_t3", "120")
    End Sub

    Private Sub HP20200901v1r5ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HP20200901v1r5ToolStripMenuItem.Click
        Switch_Spec("ECU_HP_20200901_v1r5", "218")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ' Switch_Ori
        Dim aa_array() As String
        'Dim aaaa As String
        LB_LastSpec.Text = CMB_Spec_List.SelectedItem & ",1"

        aa_array = LB_LastSpec.Text.Split(",")
        Switch_Spec_and_Page(aa_array(0), aa_array(1), aa_array(2))
        Form_Spec_CHK_Ctl.Change_WorkArea_SetX()


        Set_Page_Now(Val(aa_array(2)))
        TrackBar1.Value = Val(aa_array(2))
        Update_pic()
    End Sub
    Public Sub Switch_Ori()
        Dim aa_array() As String
        Dim BBB As String

        Dim zz As String
        zz = CMB_Spec_List.SelectedItem

        If zz <> "" Then
            aa_array = zz.Split(",")

            Name_Default = aa_array(0)
            Page_Default = aa_array(1)

            Switch_Spec(aa_array(0), aa_array(1))
        Else
            MsgBox("請先選擇仕樣")
        End If

        '----------------
        ' 建立LOG資料夾
        '----------------

        BBB = Get_filePath_x_Append() & "\LOG"

        If Not IO.Directory.Exists(BBB) Then
            Debug.Print("Create_LOG_DIRECTORY:" & BBB)
            '如不存在，建立資料夾
            IO.Directory.CreateDirectory(BBB)
        End If


        Form_Spec_CHK_Ctl.Change_WorkArea_SetX()
    End Sub
    Public Sub CHK_DIR_and_Create(ByRef BBB As String)
        If Not IO.Directory.Exists(BBB) Then
            Debug.Print("Create_LOG_DIRECTORY:" & BBB)
            '如不存在，建立資料夾
            IO.Directory.CreateDirectory(BBB)
        End If
    End Sub

    Private Sub NewIDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewIDToolStripMenuItem.Click
        Dim temp_path As String
        Dim temp_para As String

        temp_path = Get_Source_DIR() & "\APP\"
        temp_para = "NewFileID.exe"

        Call Shell(temp_path & temp_para)
    End Sub

    Private Sub ESB34V13ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ESB34V13ToolStripMenuItem.Click
        Switch_Spec("ESB34_V1_3", "104")
    End Sub

    Private Sub VV20200804t2ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VV20200804t2ToolStripMenuItem.Click
        Switch_Spec("ECU_VV_20200804_t2", "119")
    End Sub

    Private Sub UN20200902t6ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UN20200902t6ToolStripMenuItem.Click
        Switch_Spec("ECU_UN_20200902_t6", "134")
    End Sub

    Private Sub CO20200804T10ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CO20200804T10ToolStripMenuItem.Click
        Switch_Spec("ECU_CO_20200804_t10", "112")
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Set_Default_Type()
    End Sub
    Public Function Get_SpecCHK_PathFile() As String
        Return Application.StartupPath & "\" & SpecCHK_PathFile 'Spec_Path.txt"
    End Function
    Private Sub Set_Default_Type()
        Dim filename1 As String
        Dim filenum As Integer

        'Get_Spec_path()
        filename1 = Get_SpecCHK_PathFile() 'Spec_Path.txt"

        'If (File.Exists(filename1)) Then
        '    ' MsgBox("找到該路徑:" & filename1)
        '    Get_Spec_path(filename1)

        'Else
        '    'Create_New_file()
        'End If

        ' filename1 = TextString



        filenum = FreeFile()
        FileOpen(filenum, filename1, OpenMode.Output)

        ' For i = 0 To QQ.Items.Count - 1

        'content = QQ.Items.Item(i)

        PrintLine(filenum, Main_Path)   '一般先從 D槽
        PrintLine(filenum, Sub_Path)    '其次_從 C槽

        PrintLine(filenum, LA_DIR)

        'PrintLine(filenum, Name_Default)
        'PrintLine(filenum, Page_Default)
        PrintLine(filenum, Name_Selected)
        PrintLine(filenum, Page_Selected)

        'Next i

        FileClose(filenum)


    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Form_Spec_CHK_Ctl.Show()
    End Sub

    Private Sub 自動掃描ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 自動掃描ToolStripMenuItem.Click
        Scan_All_Page()
    End Sub

    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click
        TextBox1.Text = Me.Location.X & "," & Me.Location.Y


    End Sub

    Private Sub Button10_Click_1(sender As Object, e As EventArgs) Handles Button10.Click
        TextBox2.Text = Form_Spec_CHK_Ctl.Location.X & "," & Form_Spec_CHK_Ctl.Location.Y
    End Sub

    Private Sub Button11_Click_1(sender As Object, e As EventArgs) Handles Button11.Click
        TextBox3.Text = Form_Spec_CHK_Top.Location.X & "," & Form_Spec_CHK_Top.Location.Y
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        TextBox4.Text = Form_Spec_All_Pc.Location.X & "," & Form_Spec_All_Pc.Location.Y
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click

        Save_Location(ComboBox3.SelectedIndex)
    End Sub
    Private Sub Save_Location(ByVal tType As Integer)

        Dim filename1 As String
        Dim filenum As Integer

        If tType = 0 Then
            filename1 = "Location.txt"
        Else
            filename1 = "Location2.txt"
        End If

        filenum = FreeFile()
        FileOpen(filenum, filename1, OpenMode.Output)




        PrintLine(filenum, TextBox1.Text)
        PrintLine(filenum, TextBox2.Text)
        PrintLine(filenum, TextBox3.Text)
        PrintLine(filenum, TextBox4.Text)


        FileClose(filenum)
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Location_Read()
    End Sub
    Private Sub Location_Read()
        Dim xfilename As String

        If ComboBox3.SelectedIndex = 0 Then
            xfilename = "Location.txt"
        Else
            xfilename = "Location2.txt"
            If (File.Exists(xfilename)) Then

            Else
                xfilename = "Location.txt"
            End If
        End If


        Dim fs As FileStream = New FileStream(xfilename, FileMode.Open)
        Using sr As StreamReader = New StreamReader(fs, _
                                                    System.Text.Encoding.Default)


            'Parameter_Line = sr.ReadLine() & vbCrLf                                     '讀取一行資料 略過標題

            TextBox1.Text = sr.ReadLine()
            TextBox2.Text = sr.ReadLine()
            TextBox3.Text = sr.ReadLine()
            TextBox4.Text = sr.ReadLine()




            sr.Close()                                                          '關閉檔案




        End Using
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        'Dim xy As Point

        'Dim xx_array() As String
        'xx_array = TextBox1.Text.Split(",")
        'xy.X = xx_array(0)
        'xy.Y = xx_array(1)
        'Me.Location = xy



        set_location(Me, TextBox1.Text)
    End Sub
    Private Sub set_location(ByRef AA As Form, ByVal BB As String)
        Dim xy As Point

        Dim xx_array() As String
        xx_array = BB.Split(",")
        xy.X = xx_array(0)
        xy.Y = xx_array(1)
        AA.Location = xy
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        Dim xy As Point

        Dim xx_array() As String
        xx_array = TextBox2.Text.Split(",")
        xy.X = xx_array(0)
        xy.Y = xx_array(1)
        Form_Spec_CHK_Ctl.Location = xy
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Dim xy As Point

        Dim xx_array() As String
        xx_array = TextBox3.Text.Split(",")
        xy.X = xx_array(0)
        xy.Y = xx_array(1)
        Form_Spec_CHK_Top.Location = xy
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Dim xy As Point

        Dim xx_array() As String
        xx_array = TextBox4.Text.Split(",")
        xy.X = xx_array(0)
        xy.Y = xx_array(1)
        Form_Spec_All_Pc.Location = xy

    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        TextBox1.Text = Me.Location.X & "," & Me.Location.Y
        TextBox2.Text = Form_Spec_CHK_Ctl.Location.X & "," & Form_Spec_CHK_Ctl.Location.Y
        TextBox3.Text = Form_Spec_CHK_Top.Location.X & "," & Form_Spec_CHK_Top.Location.Y
        TextBox4.Text = Form_Spec_All_Pc.Location.X & "," & Form_Spec_All_Pc.Location.Y

        Save_Location(ComboBox3.SelectedIndex)

    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        Local_Read()
    End Sub
    Public Sub Local_Read()
        '----------------------
        '   讀取儲存位置
        '----------------------

        Location_Read()


        Form_Spec_CHK_Ctl.Show()
        Form_Spec_CHK_Top.Show()
        Form_Spec_CHK_Bottom.Show()

        Update_pic()
        CKB_Create_Spic.Checked = False

        Form_Spec_All_Pc.Show()


        set_location(Me, TextBox1.Text)
        set_location(Form_Spec_CHK_Ctl, TextBox2.Text)
        set_location(Form_Spec_CHK_Top, TextBox3.Text)
        set_location(Form_Spec_All_Pc, TextBox4.Text)


        Set_Form_On_Top()


        Update_pic()
    End Sub
    Public Sub Set_Form_On_Top()
        Me.TopMost = True
        Form_Spec_CHK_Ctl.TopMost = True
        Form_Spec_CHK_Bottom.TopMost = True
        Form_Spec_CHK_Top.TopMost = True
        Form_Spec_All_Pc.TopMost = True
        Me.TopMost = False
        Form_Spec_CHK_Ctl.TopMost = False
        Form_Spec_All_Pc.TopMost = False
        Form_Spec_CHK_Bottom.TopMost = False
        Form_Spec_CHK_Top.TopMost = False
    End Sub


    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        TextBox5.Text = Now.ToString("yyyyMMdd_HHmm_")
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        ' Dim filename1 As String

        Dim Report_Path_with_Date_Create As String
        Dim Report_Path_with_Date_Pic_Create As String
        Dim Report_Path_with_Date_CSV_Create As String

        Dim Report_LOG_path As String
        Dim Report_LOG_HTML_path As String
        Dim Use_log As Boolean = False

        If ReportPath = "" Then
            If (Get_ExportPath_x_Append() = "") Then
                If Get_Export_Path_x() = "" Then
                    Set_Export_Path_x(Get_Export_DIR() & "\Spec\" & Spec_Name)
                End If
                Set_ExportPath_x_Append(Get_Export_Path_x() & "\" & SetX_Str)
            End If
            ReportPath = Get_ExportPath_x_Append() & "\Report"
        End If


        If Directory.Exists(ReportPath) = False Then
            ' MsgBox("找不到該路徑:" & Final_Path)
            Form_Move.Create_All_Path(ReportPath)
            'IO.Directory.CreateDirectory(ReportPath)

        End If

        '-------------------------
        '   報告存放資料夾
        '-------------------------
        Report_Path_with_Date_Create = ReportPath & "\" & TextBox5.Text
        Report_Path_with_Date_Pic_Create = ReportPath & "\" & TextBox5.Text & "\pic"
        Report_Path_with_Date_CSV_Create = ReportPath & "\" & TextBox5.Text & "\csv"

        If Directory.Exists(Report_Path_with_Date_Create) = False Then
            ' MsgBox("找不到該路徑:" & Final_Path)
            IO.Directory.CreateDirectory(Report_Path_with_Date_Create)

        End If

        '-------------
        '   建立圖片檔資料夾
        '-------------
        If Directory.Exists(Report_Path_with_Date_Pic_Create) = False Then
            ' MsgBox("找不到該路徑:" & Final_Path)
            IO.Directory.CreateDirectory(Report_Path_with_Date_Pic_Create)

        End If

        '-------------
        '   建立圖片檔資料夾
        '-------------
        If Directory.Exists(Report_Path_with_Date_Pic_Create) = False Then
            ' MsgBox("找不到該路徑:" & Final_Path)
            IO.Directory.CreateDirectory(Report_Path_with_Date_Pic_Create)

        End If

        '-------------
        '   建立 CSV檔資料夾
        '-------------
        If Directory.Exists(Report_Path_with_Date_CSV_Create) = False Then
            ' MsgBox("找不到該路徑:" & Final_Path)
            IO.Directory.CreateDirectory(Report_Path_with_Date_CSV_Create)

        End If


        'If (File.Exists(Report_Path_with_Date_Create & "\index.html")) Then

        'Else

        'End If



        If (File.Exists(FullName_ALL & "\Spec_CHK_ALL__Now.Png")) Then

        Else
            MsgBox("None Spec_CHK_ALL__Now.Png")
            Exit Sub
        End If

        If (File.Exists(FullName_ALL & "\Status_All_A.Png")) Then

        Else
            MsgBox("None Status_All_A.Png")
            Exit Sub
        End If

        If (File.Exists(FullName_ALL & "\Status_All_B.Png")) Then

        Else
            MsgBox("None Status_All_B.Png")
            Exit Sub
        End If

        If (File.Exists(FullName_ALL & "\Status_All_C.Png")) Then

        Else
            MsgBox("None Status_All_C.Png")
            Exit Sub
        End If

        If (File.Exists(FullName_ALL & "\Status_All_D.Png")) Then

        Else
            MsgBox("None Status_All_D.Png")
            Exit Sub
        End If

        File.Copy(Get_Source_DIR() & "\HTML\index.html", Report_Path_with_Date_Create & "\index.html")
        File.Copy(FullName_ALL & "\Spec_CHK_ALL__Now.Png", Report_Path_with_Date_Create & "\Spec_CHK_ALL__Now.Png")
        File.Copy(FullName_ALL & "\Status_All_A.Png", Report_Path_with_Date_Create & "\Status_All_A.Png")
        File.Copy(FullName_ALL & "\Status_All_B.Png", Report_Path_with_Date_Create & "\Status_All_B.Png")
        File.Copy(FullName_ALL & "\Status_All_C.Png", Report_Path_with_Date_Create & "\Status_All_C.Png")
        File.Copy(FullName_ALL & "\Status_All_D.Png", Report_Path_with_Date_Create & "\Status_All_D.Png")

        ' File.Copy(Get_filePath_x_Append & "\All\Spec_CHK_ALL__Now.Png", Report_Path_with_Date_Create & "\Spec_CHK_ALL__Now.Png")



        '-------------
        '   處理LOG檔
        '-------------
        If CKB_compare.Checked = False Then
            Report_LOG_path = Get_filePath_x_Append() & "\LOG\log.txt"
        Else
            Report_LOG_path = Get_filePath_x_Append() & "\LOG\log_Compare.txt"
        End If

        If (File.Exists(Report_LOG_path)) Then
            File.Copy(Report_LOG_path, Report_Path_with_Date_Create & "\log.txt")
            wait_ms(1000)

            Trans_to_Html(Report_Path_with_Date_Create & "\log.txt", 5, "", "0")
            Report_LOG_HTML_path = (Report_Path_with_Date_Create & "\log.txt").Replace(".txt", ".html")
            Use_log = True
        End If

        '-------------
        '   複製圖片
        '-------------
        My.Computer.FileSystem.CopyDirectory(Get_CHK_PicPath, Report_Path_with_Date_Pic_Create)

        '-------------
        '   複製工作檔
        '-------------
        My.Computer.FileSystem.CopyDirectory(FullName_item_Path, Report_Path_with_Date_CSV_Create)

        Dim temp_str As String

        '-------------
        '   轉換成 html
        '-------------
        For i = 1 To Total_Pages



            'wait_ms(600)  '一定要等待畫面繪製好, 才可以擷取


            temp_str = Report_Path_with_Date_CSV_Create & "\Page_" & i.ToString("000") & ".csv"
            Page_Exist(i) = False

            If File.Exists(temp_str) = True Then

                If FileLen(temp_str) <> 0 Then
                    Trans_to_Html(temp_str, 7, "", "0")
                    Page_Exist(i) = True
                Else
                    Page_Exist(i) = False
                End If

            Else
                Page_Exist(i) = False
            End If



        Next

        '-------------
        '   產出Index
        '-------------


        Dim filenum As Integer
        Dim temp_Indexfile As String
        Dim content As String

        temp_Indexfile = Report_Path_with_Date_Create & "\Index_Page.txt"

        filenum = FreeFile()
        FileOpen(filenum, temp_Indexfile, OpenMode.Output)

        content = "ALL_Status,A.html"
        PrintLine(filenum, content)
        If Use_log = True Then
            content = "log,log.html"
            PrintLine(filenum, content)
        End If

        '------------------------
        '   標記Index 檔是否存在
        '------------------------
        For i = 1 To Total_Pages
            If Page_Exist(i) = True Then
                content = "Page_" & i.ToString("000") & ",csv\Page_" & i.ToString("000") & ".html"
            Else
                'content = "No_Page_" & i.ToString("000") & ",csv\Page_" & i.ToString("000") & ".html"
                content = "Page_" & i.ToString("000") & "_None,csv\Page_" & i.ToString("000") & ".html"
            End If

            PrintLine(filenum, content)
        Next

        FileClose(filenum)
        wait_ms(1000)

        '--------------------
        '   轉換 Index_page
        '--------------------
        Trans_to_Html(temp_Indexfile, 4, "", "0")


        '================================
        '   處理 (位置上)分頁
        '================================

        temp_Indexfile = Report_Path_with_Date_Create & "\A.txt"

        filenum = FreeFile()
        FileOpen(filenum, temp_Indexfile, OpenMode.Output)

        Form_Note.Read_FormTitle()
        content = Form_Note.GetTitle() & ","
        PrintLine(filenum, content)

        'For i = 1 To Total_Pages
        content = "ALL_Status,Spec_CHK_ALL__Now.Png"
        PrintLine(filenum, content)
        content = "ALL_Status_A,Status_All_A.Png"
        PrintLine(filenum, content)
        content = "ALL_Status_B,Status_All_B.Png"
        PrintLine(filenum, content)
        content = "ALL_Status_C,Status_All_C.Png"
        PrintLine(filenum, content)
        content = "ALL_Status_C,Status_All_D.Png"
        PrintLine(filenum, content)
        ' Next

        FileClose(filenum)
        wait_ms(1000)
        Trans_to_Html(temp_Indexfile, 2, "", "0")


        '================================
        '   開啟 Report_html
        '================================
        wait_ms(1500)
        Dim filename_html As String

        filename_html = Report_Path_with_Date_Create & "\index.html"

        If (File.Exists(filename_html)) = True Then
            '        ReportFileName
            If (filename_html <> "") Then
                System.Diagnostics.Process.Start(filename_html)
            Else
                MsgBox("No (Report) File Name")
            End If
        Else
            wait_ms(1500)

            If (File.Exists(filename_html)) = True Then
                '        ReportFileName
                If (filename_html <> "") Then
                    System.Diagnostics.Process.Start(filename_html)
                Else
                    MsgBox("No (Report) File Name")
                End If
            Else

                MsgBox("找不到該檔案:" & filename_html)
            End If
        End If

        'MsgBox("Done.")

    End Sub
    '=======================================
    '   1.檔名
    '   2.轉換類別
    '   3.附加測試結果
    '--------------------------
    '   類別
    '   (0) OK/NG
    '   (1) ++/--
    '   (2) PNG
    '   (3) HYBER LINK
    '   (4) HYBER LINK(right frame)
    '   (5) 表格(一段)
    '   (6) 表格(Script)
    '   (7) 表格()
    '   (8)
    '   (9)
    '
    '=======================================
    Public Sub Trans_to_Html(TXT_file As String, ByVal t_Type As Integer, ByVal Append_str As String, ByVal Font_str As String)

        Dim temp_path As String
        Dim temp_para As String
        Dim A0 As String = "BUS_ECU_1234"

        temp_path = Get_Source_DIR() & "\APP\"
        temp_para = "VB_HTML.exe " & A0 & " " & TXT_file & " " & t_Type & " " & Append_str & " " & Font_str

        Call Shell(temp_path & temp_para)
    End Sub
    Private Sub PathToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PathToolStripMenuItem.Click
        Form_Path.Show()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CMB_Spec_List.SelectedIndexChanged

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CMB_Compare.SelectedIndexChanged

        Get_Compare_Type()
    End Sub
    Public Sub Get_Compare_Type()
        Dim temp_xxx() As String
        Dim temp_op As String
        'temp_op = CMB_Compare.SelectedItem
        temp_op = CMB_Compare.Text
        temp_xxx = temp_op.Split(",")
        Spec_Compare_Name = temp_xxx(0)
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        'Dim Spec_Diff As New Form_Spec_All_Pc
        If spec_diff_Opened = False Then

            Spec_Diff.Set_ALLPic_Type(True)
            Spec_Diff.Show()
            Spec_Diff.Text = Spec_Diff.Text & " - Compare Different With PreVersion"
            spec_diff_Opened = True

        Else
            MsgBox("請關閉程式重新開啟, 才能再使用")
        End If

    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        Manual_Switch_Size()



    End Sub
    Public Sub Manual_Switch_Size()
        If TextBox6.Text <> "" Then
            Form_Spec_CHK_Top.XChange_Ratio_and_Resize(Val(TextBox6.Text))
        Else
            Form_Spec_CHK_Top.Drag_Cal_Ratio_and_Resize()
        End If


        Form_Spec_CHK_Top.Switch_Compare_Mode()

        Page_Now()
        Debug.Print("NFS_1")
        Update_for_Sync()
    End Sub

    Private Sub LB_LastSpec_Click(sender As Object, e As EventArgs) Handles LB_LastSpec.Click
        '-----------------------
        '   切換回 前一個轉跳過來的仕樣
        '-----------------------
        Switch_temp_Spec()

    End Sub
    Public Sub Switch_temp_Spec()
        Dim aa_array() As String
        'Dim aaaa As String
        aa_array = LB_LastSpec.Text.Split(",")

        Switch_Spec_and_Page(aa_array(0), aa_array(1), aa_array(2))

        Form_Spec_CHK_Ctl.Change_WorkArea_SetX()

        '—————————————————
        '為了顯示部份的同步
        '—————————————————

        Set_Page_Now(Val(aa_array(2)))
        TrackBar1.Value = Val(aa_array(2))
        Update_pic()
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        TextBox6.Text = "0.5"
        Manual_Switch_Size()
    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        TextBox6.Text = "1"
        Manual_Switch_Size()
    End Sub

    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click
        Dim x_str As String
        Dim x_str_array() As String

        x_str = LB_LastSpec.Text
        x_str_array = x_str.Split(",")

        If UBound(x_str_array) >= 2 Then
            Form_Spec_CHK_Ctl.TB_Title1.Text = "XLINK"
            Form_Spec_CHK_Ctl.TB_subTitle1.Text = x_str_array(0)
            Form_Spec_CHK_Ctl.TB_Title2.Text = x_str_array(1)
            Form_Spec_CHK_Ctl.TB_subTitle2.Text = x_str_array(2)

            Form_Spec_CHK_Ctl.CMB_Font_Color.SelectedIndex = 6 '(黃綠色)
            Form_Spec_CHK_Ctl.CMB_BG_Color.SelectedIndex = 0 '(無色)
        Else
            MsgBox("請先選擇目標")
        End If

    End Sub

    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        '==============================
        '   產生原始圖的XAppend
        '==============================
        Form_Spec_CHK_Ctl.TB_Title1.Text = "XAppend"

        If CKB_Relative_Path.Checked = True Then
            '--------------------
            '   使用相對路徑
            '--------------------
            Form_Spec_CHK_Ctl.TB_Note.Text = TB_FName.Text.Replace(Get_Source_DIR, "..")
        Else
            Form_Spec_CHK_Ctl.TB_Note.Text = TB_FName.Text
        End If

        Form_Spec_CHK_Ctl.CMB_Font_Color.SelectedIndex = 10 '(淺藍色)
        Form_Spec_CHK_Ctl.CMB_BG_Color.SelectedIndex = 0 '(無色)
    End Sub

    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click
        '==============================
        '   產生工作圖的XAppend
        '==============================
        Dim temp_x_path As String
        Dim temp_x_path_array() As String
        Dim temp_target As String
        Dim temp_target_no As Integer
        Form_Spec_CHK_Ctl.TB_Title1.Text = "XAppend"

        If CKB_Relative_Path.Checked = True Then
            temp_x_path = Get_work_DIR()
            temp_x_path_array = temp_x_path.Split("\")
            temp_target_no = UBound(temp_x_path_array)
            temp_target = temp_x_path_array(temp_target_no)

            'Form_Spec_CHK_Ctl.TB_Note.Text = TB_CHK_PicName.Text.Replace(temp_x_path_array(0) & "\" & temp_x_path_array(1), "..\..")
            Form_Spec_CHK_Ctl.TB_Note.Text = TB_CHK_PicName.Text.Replace(Get_work_DIR, "..\..\" & temp_target)
        Else
            Form_Spec_CHK_Ctl.TB_Note.Text = TB_CHK_PicName.Text
        End If

        Form_Spec_CHK_Ctl.CMB_Font_Color.SelectedIndex = 10 '(淺藍色)
        Form_Spec_CHK_Ctl.CMB_BG_Color.SelectedIndex = 0 '(無色)
    End Sub


    Private Sub CKB_compare_CheckedChanged(sender As Object, e As EventArgs) Handles CKB_compare.CheckedChanged
        Dim xx_open As Boolean = False

        If Form_Spec_CHK_Top.Created = True Then
            xx_open = True
            Form_Spec_CHK_Top.Close()
            Form_Spec_CHK_Bottom.Close()
            Form_Spec_CHK_Top.Show()
            Form_Spec_CHK_Bottom.Show()

            Form_Spec_CHK_Bottom.Location = Form_Spec_CHK_Top.Location

            Form_Spec_CHK_Bottom.TopMost = True
            Form_Spec_CHK_Top.TopMost = True
            Form_Spec_CHK_Bottom.TopMost = False
            Form_Spec_CHK_Top.TopMost = False

            xx_open = False
        End If


        Form_Spec_CHK_Ctl.Change_WorkArea_SetX()

        If CKB_compare.Checked = True Then
            ComboBox3.SelectedIndex = 1
        Else
            ComboBox3.SelectedIndex = 0
        End If

        If CKB_Auto.Checked = True Then
            Local_Read()
        End If



    End Sub

    Private Sub MaskCopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MaskCopyToolStripMenuItem.Click
        Form_MaskCopy.Show()
    End Sub

    Private Sub StatusAnalysisToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StatusAnalysisToolStripMenuItem.Click
        Form_status.Show()
    End Sub

    Private Sub CKB_USE_WorkPIC_CheckedChanged(sender As Object, e As EventArgs) Handles CKB_USE_WorkPIC.CheckedChanged
        Page_Now()
        Debug.Print("NFS_2")
        Update_for_Sync()
    End Sub

    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Button35.Click
        '------------------
        '   手動按下讀取
        '------------------
        Read_Page_Index_file(Spec_index_path, True)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Form_status.Open_X_file(Spec_index_path)
    End Sub

    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Button36.Click



        '----------------------
        '   儲存位置 移到最前
        '----------------------

        'Location_Read()

        TextBox1.Text = "10,10"
        TextBox2.Text = "50,50"
        TextBox3.Text = "100,100"
        TextBox4.Text = "150,150"



        Form_Spec_CHK_Ctl.Show()
        Form_Spec_CHK_Top.Show()
        Form_Spec_CHK_Bottom.Show()

        Update_pic()
        CKB_Create_Spic.Checked = False

        Form_Spec_All_Pc.Show()


        set_location(Me, TextBox1.Text)
        set_location(Form_Spec_CHK_Ctl, TextBox2.Text)
        set_location(Form_Spec_CHK_Top, TextBox3.Text)
        set_location(Form_Spec_All_Pc, TextBox4.Text)


        Set_Form_On_Top()


        Update_pic()
    End Sub

    Private Sub MoveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MoveToolStripMenuItem.Click
        Form_Move.Show()
    End Sub
End Class