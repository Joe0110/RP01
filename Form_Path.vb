Public Class Form_Path

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Load_Path()


    End Sub
    Public Sub Load_Path()
        ListBox1.Items.Clear()

        Path_Display_add("Word Dir", Form_Spec.Get_Source_DIR)
        Path_Display_add("Get_Compare_Path", Form_Spec.Get_Compare_Path)
        Path_Display_add("Get_file_Path_x", Form_Spec.Get_file_Path_x)
        Path_Display_add(" ", " ")

        Path_Display_add("Pattern_Type", Form_Spec.Pattern_Type)
        Path_Display_add(" ", " ")
        Path_Display_add("FullName_Index_Path", Form_Spec.FullName_Index_Path)
        Path_Display_add("FullName_item_Path", Form_Spec.FullName_item_Path)
        Path_Display_add("FullName_item_CSV", Form_Spec.FullName_item_CSV)
        Path_Display_add("Full_Target_Path", Form_MaskCopy.Full_Target_Path)
        'Full_Target_Path()
        Path_Display_add(" ", " ")
        ' Path_Display_add("FullName_ALL", Form_Spec.FullName_ALL)

        Path_Display_add("Get_Full_PathSmall_Name", Form_Spec.Get_Full_PathSmall_Name)
        Path_Display_add("Get_Full_Small_Name", Form_Spec.Get_Full_Small_Name)
        Path_Display_add(" ", " ")

        Path_Display_add("Get_CHK_PicPath", Form_Spec.Get_CHK_PicPath)
        Path_Display_add("Get_CHK_PicName", Form_Spec.Get_CHK_PicName)
        Path_Display_add(" ", " ")

        'Path_Display_add("Pattern_Type", Form_Spec.Pattern_Type)

        Path_Display_add("FullName_ALL", Form_Spec.FullName_ALL)



        Path_Display_add("ReportPath", Form_Spec.ReportPath)

        Path_Display_add(" ", " ")

        Path_Display_add("Spec_Name", Form_Spec.Spec_Name)
        Path_Display_add("Name_Default", Form_Spec.Name_Default)
        Path_Display_add("Page_Default", Form_Spec.Page_Default)
        'FullName_Index_Path

    End Sub
    Public Sub Path_Display_add(ByVal tName As String, ByVal tPath As String)
        'Dim xxx As String
        Dim x_len As Integer
        x_len = Len(tName)

        If x_len < 20 Then
            For i = 0 To (20 - x_len)
                tName = tName & " "
            Next
        End If

        ' xxx = Format(tName, "############")
        ListBox1.Items.Add(tName & tPath)
    End Sub

    Private Sub Form_Path_Load(sender As Object, e As EventArgs) Handles Me.Load
        Load_Path()
    End Sub
End Class