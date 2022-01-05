Imports System.IO '新增此命名空間

Public Class Form_Spec_All_Pc

    Public Const Default_Title As String = "Form_Spec_All_Pc"
    Public FullName_item_Small_PIC As String

    Public PicType As Boolean = False

    Public Pic_Created As Boolean = False


    Public PB_Array(500) As PictureBox
    Public Sub Set_ALLPic_Type(ByVal AAA As Boolean)
        PicType = AAA
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Load_XS_pic()

    End Sub
    Public Sub Load_XS_pic()
        If Pic_Created = False Then
            'Setup_PageX2(Form_Spec_All_Pc, PB_Array, PB_1, 1, Form_Spec.Total_Pages, 16)
            Setup_PageX2(GroupBox1, PB_Array, PB_1, 1, Form_Spec.Total_Pages, 16)
            Pic_Created = True
        End If
    End Sub
    Public Sub Get_Pic_Name(ByVal dNO As Integer)
        'Dim FullName_item_CSV As String
        ' Dim Get_file_Path_x As String
        Dim xpage As Integer

        xpage = dNO

        If PicType = False Then
            'Debug.Print("FullName_Small_Pat：h" & Form_Spec.Get_Full_PathSmall_Name)
            FullName_item_Small_PIC = Form_Spec.Get_Full_PathSmall_Name & "\Page_" & xpage.ToString("000") & "_s.png"
        Else
            FullName_item_Small_PIC = Form_Spec.Get_Full_PathSmall_Name & "\compare\Page_" & xpage.ToString("000") & "_s.png"
        End If



    End Sub

    Public Sub ReSize_PB_Array(ByRef KK As GroupBox, ByRef AA() As PictureBox, ByRef A1 As PictureBox, ByVal dStart As Integer, ByVal dEnd As Integer, ByVal LineNo As Integer)
        Dim a1x As String
        Dim tempx As Integer
        Dim tempq As Integer
        Dim TB_width As Integer
        Dim TB_Height As Integer
        Dim plus_Height As Integer
        Dim plus_width As Integer

        Try


            AA(dStart) = A1                  '建立CH陣列
            Get_Pic_Name(dStart)

            If (File.Exists(FullName_item_Small_PIC)) Then
                A1.Load(FullName_item_Small_PIC)
            End If

            A1.SizeMode = PictureBoxSizeMode.StretchImage

            a1x = A1.Name

            tempx = InStr(a1x, "_")
            a1x = Mid(a1x, 1, tempx)

            TB_width = A1.Width
            TB_Height = A1.Height

            plus_width = TB_width * (LineNo) + 20

            For i = dStart + 1 To dEnd
                'AA(i) = New PictureBox
                tempq = (i - 1) \ LineNo

                'Debug.Print(tempq)

                AA(i).Location = New System.Drawing.Point(A1.Location.X + (TB_width * ((i - 1) Mod LineNo)), A1.Location.Y + (TB_Height * tempq))

                plus_Height = TB_Height * (tempq + 1) + 40

                '-----------
                '高度調整
                '-----------
                If KK.Height < plus_Height Then
                    KK.Height = plus_Height
                Else
                    KK.Height = plus_Height
                End If

                '-----------
                '宽度調整
                '-----------
                If KK.Width < plus_width Then
                    KK.Width = plus_width
                Else
                    KK.Width = plus_width
                End If

                AA(i).Name = a1x & i
                AA(i).Text = i
                AA(i).Width = A1.Width
                AA(i).Height = A1.Height
                AA(i).BorderStyle = BorderStyle.FixedSingle
                AA(i).SizeMode = PictureBoxSizeMode.StretchImage
                Get_Pic_Name(i)

                If AA(i).Created = True Then
                    If (File.Exists(FullName_item_Small_PIC)) Then
                        AA(i).Load(FullName_item_Small_PIC)
                    End If

                End If

                ' KK.Controls.Add(AA(i))


            Next

        Catch ex As Exception

        Finally

        End Try


    End Sub
    Public Sub Setup_PageX2(ByRef KK As GroupBox, ByRef AA() As PictureBox, ByRef A1 As PictureBox, ByVal dStart As Integer, ByVal dEnd As Integer, ByVal LineNo As Integer)
        Dim a1x As String
        Dim tempx As Integer
        Dim tempq As Integer
        Dim TB_width As Integer
        Dim TB_Height As Integer
        Dim plus_Height As Integer


        AA(dStart) = A1                  '建立CH陣列
        Get_Pic_Name(dStart)

        If (File.Exists(FullName_item_Small_PIC)) Then

            A1.Load(FullName_item_Small_PIC)
        End If

        A1.SizeMode = PictureBoxSizeMode.StretchImage

        a1x = A1.Name

        tempx = InStr(a1x, "_")
        a1x = Mid(a1x, 1, tempx)

        TB_width = A1.Width
        TB_Height = A1.Height

        For i = dStart + 1 To dEnd
            AA(i) = New PictureBox
            tempq = (i - 1) \ LineNo

            'Debug.Print(tempq)

            AA(i).Location = New System.Drawing.Point(A1.Location.X + (TB_width * ((i - 1) Mod LineNo)), A1.Location.Y + (TB_Height * tempq))

            plus_Height = TB_Height * (tempq + 1) + 30

            If KK.Height < plus_Height Then
                KK.Height = plus_Height
            Else
                KK.Height = plus_Height
            End If

            AA(i).Name = a1x & i
            AA(i).Text = i
            AA(i).Width = A1.Width
            AA(i).Height = A1.Height
            AA(i).BorderStyle = BorderStyle.FixedSingle
            AA(i).SizeMode = PictureBoxSizeMode.StretchImage
            Get_Pic_Name(i)

            If (File.Exists(FullName_item_Small_PIC)) Then
                AA(i).Load(FullName_item_Small_PIC)
            End If

            AddHandler AA(i).Click, AddressOf PICs_Click

            KK.Controls.Add(AA(i))

            'AA(i).Dispose()
        Next
        'A1.Dispose()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Resize_Init()
    End Sub
    Public Sub Resize_Init()
        Dim line_no As Integer

        PB_1.Width = Val(TextBox3.Text)
        PB_1.Height = Val(TextBox4.Text)
        line_no = Val(TextBox1.Text)

        ReSize_PB_Array(GroupBox1, PB_Array, PB_1, 1, Form_Spec.Total_Pages, line_no)

        Set_Arror(Val(Label4.Text))

        Me.Height = GroupBox1.Location.Y + GroupBox1.Height + 25
    End Sub
    '===========================
    '   [ SP00 ] 
    '===========================

    Private Sub PICs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim To_Page As Integer

        Label4.Text = sender.text

        To_Page = Val(sender.text)

        Debug.Print("Change_Page_From_PICs_Click")
        Form_Spec.TrackBar1.Value = To_Page

        'Form_Spec.Change_Page(To_Page)

        Set_Arror(To_Page)

    End Sub

    Private Sub Form_Spec_All_Pc_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = "p" Then
            '------------------
            '   最上層顯示
            '------------------
            Form_Spec.Set_Form_On_Top()
        End If

    End Sub
    Private Sub Form_Spec_All_Pc_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Text = Default_Title
        'If Form_Spec.CKB_compare.Checked = True Then
        '    Set_ALLPic_Type(True)
        'End If

        Load_XS_pic()

        Load_Setup()

        Resize_Init()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim temp As String

        '最近狀態
        temp = Form_Spec.FullName_ALL & "\Spec_CHK_ALL__Now.Png"
        Capture_SavePNG_Ori(temp)

        '保留備查用狀態
        temp = Form_Spec.FullName_ALL & "\Spec_CHK_ALL_" & ComboBox1.Text & "_" & Now.ToString("yyyyMMdd_HHmmss") & ".Png"
        TB_CaptureName.Text = temp

        Capture_SavePNG_Ori(temp)
    End Sub
    Public Sub Capture_SavePNG_Ori(ByVal zz As String)
        Dim xx As Integer
        Dim yy As Integer
        Dim d1 As Integer
        Dim d2 As Integer

        'Label4.Text = TabControl2.Left
        ' Label5.Text = TabControl2.Top

        d1 = 7
        d2 = 28

        xx = GroupBox1.Width
        yy = GroupBox1.Height

        'Dim myImage As New Bitmap(1024, 768)
        Dim Screen_Img As New Bitmap(xx, yy)
        Dim g = Graphics.FromImage(Screen_Img)
        'Dim Small_bmp As New Bitmap(xx, yy)
        ' g.CopyFromScreen(New Point(Me.Left + 38, Me.Top + 45), New Point(0, 0), ScrnSize) '以自己視窗為始點修正偏移

        '==============================================
        '擷取螢幕


        g.CopyFromScreen(New Point(Me.Location.X + GroupBox1.Left + d1, Me.Location.Y + GroupBox1.Top + d2), New Point(0, 0), New Size(xx, yy))


        'g.CopyFromScreen(New Point(TabControl2.Left, TabControl2.Top), New Point(TabControl2.Right, TabControl2.Bottom), New Size(xx, yy))
        ' g.CopyFromScreen(New Point(TabControl2.Left, TabControl2.Top), New Point(TabControl2.Right, TabControl2.Bottom), New Size(xx, yy))

        '==============================================
        '改大小

        'Dim ScrnPB As New PictureBox
        'ScrnPB.Width = 720
        'ScrnPB.Height = 960

        'ScrnPB.Image = Screen_Img

        ' img => bmp
        'Dim gfx As Graphics = Graphics.FromImage(Small_bmp)
        ' gfx.DrawImage(ScrnPB.Image, 0, 0, xx, yy)

        'gfx.Dispose()


        '==============================================

        Screen_Img.Save(zz, System.Drawing.Imaging.ImageFormat.Png)

        Screen_Img.Dispose()
        ' Small_bmp.Dispose()
        'ScrnPB.Dispose()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim Full_target_name As String

        ' Full_target_name = ComboBox1.Text & "_" & TB_CaptureName.Text
        Full_target_name = TB_CaptureName.Text

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

    Private Sub Form_Spec_All_Pc_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        Label6.Text = e.X & "," & e.Y

    End Sub

    Private Sub PB_1_Click(sender As Object, e As EventArgs) Handles PB_1.Click
        Form_Spec.Change_Page(1)
        Set_Arror(1)
    End Sub


    Public Sub Set_Arror(ByVal Addr As Integer)
        Dim xx As Point
        Try
            If PB_Array(Addr).Created Then
                xx = PB_Array(Addr).Location
                xx.X += 15
                xx.Y += 50

                Label4.Text = Addr

                If CheckBox1.Checked = True Then
                    PicBox_Arrow.Location = xx
                    PicBox_Arrow.Visible = True
                Else
                    PicBox_Arrow.Visible = False
                End If
            End If
        Catch ex As Exception

        Finally

        End Try

    End Sub
    Public Sub Use_temp_Small_Img(ByVal DD As Integer)
        Try
            If PB_Array(DD).Created Then
                PB_Array(DD).Load(Form_Spec.Get_Source_DIR & "\Spec\TempIMG.png")
            End If
        Catch ex As Exception

        Finally

        End Try


    End Sub

    Public Sub Use_Ori_Small_Img(ByVal DD As Integer)
        Try
            Get_Pic_Name(DD)
            PB_Array(DD).Load(FullName_item_Small_PIC)
        Catch ex As Exception

        Finally

        End Try


    End Sub



    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            PicBox_Arrow.Visible = True
        Else
            PicBox_Arrow.Visible = False
        End If

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If CheckBox1.Checked = True Then
            If PicBox_Arrow.Visible = False Then
                PicBox_Arrow.Visible = True
            Else
                PicBox_Arrow.Visible = False
            End If
        Else
            PicBox_Arrow.Visible = False

        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim filename1 As String
        Dim content As String
        Dim filenum As Integer

        filename1 = Form_Spec.Get_file_Path_x & "\Set0\Index\AllPic_setup.txt"

        filenum = FreeFile()
        FileOpen(filenum, filename1, OpenMode.Output)


        content = TextBox3.Text & "," & TextBox4.Text & "," & TextBox1.Text

        PrintLine(filenum, content)


        FileClose(filenum)
    End Sub
    Private Sub Load_Setup()
        Dim Parameter_Line As String = ""
        Dim filename1 As String
        'Dim content As String
        ' Dim filenum As Integer
        Dim xx_array() As String

        filename1 = Form_Spec.Get_file_Path_x & "\Set0\Index\AllPic_setup.txt"


        If (File.Exists(filename1)) Then





            '開啟證件
            Dim fs As FileStream = New FileStream(filename1, FileMode.Open)
            Using sr As StreamReader = New StreamReader(fs, _
                                                        System.Text.Encoding.Default)


                '讀取一行資料 略過標題



                'While (sr.EndOfStream <> True)



                '送出相關設定
                Parameter_Line = sr.ReadLine()                                    '讀取一行資料

                xx_array = Parameter_Line.Split(",")


                TextBox3.Text = xx_array(0)
                TextBox4.Text = xx_array(1)
                TextBox1.Text = xx_array(2)

                'End While







                sr.Close()                                                          '關閉檔案




            End Using

        End If

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Width = (Val(TextBox3.Text) * Val(TextBox1.Text)) + 60
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox3.Text = "30"
        TextBox4.Text = "40"
        Resize_Init()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TextBox3.Text = "60"
        TextBox4.Text = "80"
        Resize_Init()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        TextBox3.Text = "120"
        TextBox4.Text = "80"
        Resize_Init()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        TextBox3.Text = "240"
        TextBox4.Text = "160"
        Resize_Init()
    End Sub
End Class