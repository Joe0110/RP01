<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Move
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TB_Target = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TB_Source = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.TB_Target2 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TB_BasePath_New = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TB_WorkPath_New = New System.Windows.Forms.TextBox()
        Me.TB_ExportPath_New = New System.Windows.Forms.TextBox()
        Me.TB_ExportPath = New System.Windows.Forms.TextBox()
        Me.TB_WorkPath = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.TB_BasePath = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.LB_Old_Flag = New System.Windows.Forms.Label()
        Me.LB_New_Flag = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.LB_PathFile_Version = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.LB_Setup_Filename = New System.Windows.Forms.Label()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TB_Target
        '
        Me.TB_Target.Location = New System.Drawing.Point(106, 323)
        Me.TB_Target.Name = "TB_Target"
        Me.TB_Target.Size = New System.Drawing.Size(182, 22)
        Me.TB_Target.TabIndex = 0
        Me.TB_Target.Text = "C:\Joe\ECU_TEST_work"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(311, 318)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(71, 27)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "CHK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(45, 326)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 12)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "New_Path:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(45, 400)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 12)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Select_Src:"
        '
        'TB_Source
        '
        Me.TB_Source.Location = New System.Drawing.Point(106, 431)
        Me.TB_Source.Name = "TB_Source"
        Me.TB_Source.Size = New System.Drawing.Size(182, 22)
        Me.TB_Source.TabIndex = 4
        Me.TB_Source.Text = "C:\Joe\ECU_TEST_work"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(294, 431)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(71, 27)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Select"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(311, 351)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(95, 34)
        Me.Button3.TabIndex = 6
        Me.Button3.Text = "Action"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'TB_Target2
        '
        Me.TB_Target2.Location = New System.Drawing.Point(106, 351)
        Me.TB_Target2.Name = "TB_Target2"
        Me.TB_Target2.Size = New System.Drawing.Size(182, 22)
        Me.TB_Target2.TabIndex = 7
        Me.TB_Target2.Text = "C:\Joe\ECU_TEST"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(39, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 12)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "舊版:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(605, 72)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(32, 12)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "新版:"
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("PMingLiU", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Button4.Location = New System.Drawing.Point(500, 66)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(71, 27)
        Me.Button4.TabIndex = 10
        Me.Button4.Text = "--->"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Font = New System.Drawing.Font("PMingLiU", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Button5.Location = New System.Drawing.Point(500, 105)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(71, 27)
        Me.Button5.TabIndex = 11
        Me.Button5.Text = "<---"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(756, 120)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(44, 12)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "設定檔:"
        '
        'TB_BasePath_New
        '
        Me.TB_BasePath_New.Location = New System.Drawing.Point(806, 117)
        Me.TB_BasePath_New.Name = "TB_BasePath_New"
        Me.TB_BasePath_New.Size = New System.Drawing.Size(182, 22)
        Me.TB_BasePath_New.TabIndex = 13
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(739, 120)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(11, 12)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "_"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(744, 185)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(56, 12)
        Me.Label7.TabIndex = 15
        Me.Label7.Text = "工作路徑:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(744, 218)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(56, 12)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "匯出路徑:"
        '
        'TB_WorkPath_New
        '
        Me.TB_WorkPath_New.Location = New System.Drawing.Point(806, 182)
        Me.TB_WorkPath_New.Name = "TB_WorkPath_New"
        Me.TB_WorkPath_New.Size = New System.Drawing.Size(182, 22)
        Me.TB_WorkPath_New.TabIndex = 17
        '
        'TB_ExportPath_New
        '
        Me.TB_ExportPath_New.Location = New System.Drawing.Point(806, 215)
        Me.TB_ExportPath_New.Name = "TB_ExportPath_New"
        Me.TB_ExportPath_New.Size = New System.Drawing.Size(182, 22)
        Me.TB_ExportPath_New.TabIndex = 18
        '
        'TB_ExportPath
        '
        Me.TB_ExportPath.Location = New System.Drawing.Point(106, 215)
        Me.TB_ExportPath.Name = "TB_ExportPath"
        Me.TB_ExportPath.Size = New System.Drawing.Size(182, 22)
        Me.TB_ExportPath.TabIndex = 25
        '
        'TB_WorkPath
        '
        Me.TB_WorkPath.Location = New System.Drawing.Point(106, 182)
        Me.TB_WorkPath.Name = "TB_WorkPath"
        Me.TB_WorkPath.Size = New System.Drawing.Size(182, 22)
        Me.TB_WorkPath.TabIndex = 24
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(44, 218)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(56, 12)
        Me.Label9.TabIndex = 23
        Me.Label9.Text = "匯出路徑:"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(44, 185)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(56, 12)
        Me.Label10.TabIndex = 22
        Me.Label10.Text = "工作路徑:"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(39, 120)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(11, 12)
        Me.Label11.TabIndex = 21
        Me.Label11.Text = "_"
        '
        'TB_BasePath
        '
        Me.TB_BasePath.Location = New System.Drawing.Point(106, 117)
        Me.TB_BasePath.Name = "TB_BasePath"
        Me.TB_BasePath.Size = New System.Drawing.Size(182, 22)
        Me.TB_BasePath.TabIndex = 20
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(56, 120)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(44, 12)
        Me.Label12.TabIndex = 19
        Me.Label12.Text = "設定檔:"
        '
        'LB_Old_Flag
        '
        Me.LB_Old_Flag.AutoSize = True
        Me.LB_Old_Flag.Location = New System.Drawing.Point(20, 72)
        Me.LB_Old_Flag.Name = "LB_Old_Flag"
        Me.LB_Old_Flag.Size = New System.Drawing.Size(11, 12)
        Me.LB_Old_Flag.TabIndex = 26
        Me.LB_Old_Flag.Text = "_"
        '
        'LB_New_Flag
        '
        Me.LB_New_Flag.AutoSize = True
        Me.LB_New_Flag.Location = New System.Drawing.Point(588, 72)
        Me.LB_New_Flag.Name = "LB_New_Flag"
        Me.LB_New_Flag.Size = New System.Drawing.Size(11, 12)
        Me.LB_New_Flag.TabIndex = 27
        Me.LB_New_Flag.Text = "_"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(229, 31)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(105, 12)
        Me.Label15.TabIndex = 28
        Me.Label15.Text = "[ Spec_Path.txt ] Ver:"
        '
        'LB_PathFile_Version
        '
        Me.LB_PathFile_Version.AutoSize = True
        Me.LB_PathFile_Version.Location = New System.Drawing.Point(340, 31)
        Me.LB_PathFile_Version.Name = "LB_PathFile_Version"
        Me.LB_PathFile_Version.Size = New System.Drawing.Size(23, 12)
        Me.LB_PathFile_Version.TabIndex = 29
        Me.LB_PathFile_Version.Text = "___"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(18, 9)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(32, 12)
        Me.Label16.TabIndex = 30
        Me.Label16.Text = "檔名:"
        '
        'LB_Setup_Filename
        '
        Me.LB_Setup_Filename.AutoSize = True
        Me.LB_Setup_Filename.Location = New System.Drawing.Point(56, 9)
        Me.LB_Setup_Filename.Name = "LB_Setup_Filename"
        Me.LB_Setup_Filename.Size = New System.Drawing.Size(23, 12)
        Me.LB_Setup_Filename.TabIndex = 31
        Me.LB_Setup_Filename.Text = "___"
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(20, 28)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(59, 26)
        Me.Button6.TabIndex = 32
        Me.Button6.Text = "開啟"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Form_Move
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1163, 528)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.LB_Setup_Filename)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.LB_PathFile_Version)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.LB_New_Flag)
        Me.Controls.Add(Me.LB_Old_Flag)
        Me.Controls.Add(Me.TB_ExportPath)
        Me.Controls.Add(Me.TB_WorkPath)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.TB_BasePath)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.TB_ExportPath_New)
        Me.Controls.Add(Me.TB_WorkPath_New)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TB_BasePath_New)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TB_Target2)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.TB_Source)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TB_Target)
        Me.Name = "Form_Move"
        Me.Text = "Form_Move"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TB_Target As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TB_Source As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents TB_Target2 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TB_BasePath_New As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TB_WorkPath_New As System.Windows.Forms.TextBox
    Friend WithEvents TB_ExportPath_New As System.Windows.Forms.TextBox
    Friend WithEvents TB_ExportPath As System.Windows.Forms.TextBox
    Friend WithEvents TB_WorkPath As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TB_BasePath As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents LB_Old_Flag As System.Windows.Forms.Label
    Friend WithEvents LB_New_Flag As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents LB_PathFile_Version As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents LB_Setup_Filename As System.Windows.Forms.Label
    Friend WithEvents Button6 As System.Windows.Forms.Button
End Class
