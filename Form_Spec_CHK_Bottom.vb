Public Class Form_Spec_CHK_Bottom



    Public PB2_ori_x As Integer
    Public PB2_ori_y As Integer

    Private Sub Form_Spec_CHK_Bottom_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        If Form_Spec_CHK_Top.focus_Flag = False Then
            Form_Spec_CHK_Top.focus_Flag = True
            Form_Spec_CHK_Top.Set_to_TOP()
        End If
    End Sub



















    Private Sub Form_Spec_CHK_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Form_Spec_CHK2.Show()
        'Dim xy_P As Point



        Form_Spec_CHK_Top.S_W_Ratio = 1
        Form_Spec_CHK_Top.S_H_Ratio = 1

        PB2_ori_x = PicBox_Comp.Location.X
        PB2_ori_y = PicBox_Comp.Location.Y


        If Form_Spec.CKB_compare.Checked = False Then
            '919, 1050
            Me.Width = Form_Spec_CHK_Top.get_Form_Width_Single


            PicBox_Comp.Visible = False
        Else

            '1900,1050
            Me.Width = Form_Spec_CHK_Top.get_Form_Width_Double


            PicBox_Comp.Visible = True
            PicBox_Comp.Size = PicBox_Main.Size

            Create_Compare_WorkDIR()


        End If

        'Form_Spec_CHK2.Overlap_windows()



    End Sub
    Public Sub Create_Compare_WorkDIR()
        Dim chk_Path_x As String

        chk_Path_x = Form_Spec.Get_file_Path_x & "\DATA_CSV\compare"
        CHK_DIR_Create(chk_Path_x)

        chk_Path_x = Form_Spec.Get_file_Path_x & "\All\compare"
        CHK_DIR_Create(chk_Path_x)

        chk_Path_x = Form_Spec.Get_file_Path_x & "\sp\compare"
        CHK_DIR_Create(chk_Path_x)

        chk_Path_x = Form_Spec.Get_file_Path_x & "\Normal\compare"
        CHK_DIR_Create(chk_Path_x)

        Form_Spec_CHK_Ctl.Load_Compare_Offset_File()
    End Sub
    Private Sub CHK_DIR_Create(ByVal AAA As String)
        If Not IO.Directory.Exists(AAA) Then

            '如不存在，建立資料夾
            IO.Directory.CreateDirectory(AAA)
            Debug.Print("CHK_DIR_Create:" & AAA)
        End If
    End Sub


    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PicBox_Comp.Click

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PicBox_Main.Click

    End Sub
End Class