<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Spec_CHK_Bottom
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form_Spec_CHK_Bottom))
        Me.PicBox_Main = New System.Windows.Forms.PictureBox()
        Me.PicBox_Comp = New System.Windows.Forms.PictureBox()
        CType(Me.PicBox_Main, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBox_Comp, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicBox_Main
        '
        Me.PicBox_Main.Image = CType(resources.GetObject("PicBox_Main.Image"), System.Drawing.Image)
        Me.PicBox_Main.Location = New System.Drawing.Point(29, 21)
        Me.PicBox_Main.Name = "PicBox_Main"
        Me.PicBox_Main.Size = New System.Drawing.Size(720, 960)
        Me.PicBox_Main.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicBox_Main.TabIndex = 0
        Me.PicBox_Main.TabStop = False
        '
        'PicBox_Comp
        '
        Me.PicBox_Comp.Image = CType(resources.GetObject("PicBox_Comp.Image"), System.Drawing.Image)
        Me.PicBox_Comp.Location = New System.Drawing.Point(765, 21)
        Me.PicBox_Comp.Name = "PicBox_Comp"
        Me.PicBox_Comp.Size = New System.Drawing.Size(118, 148)
        Me.PicBox_Comp.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicBox_Comp.TabIndex = 22
        Me.PicBox_Comp.TabStop = False
        Me.PicBox_Comp.Visible = False
        '
        'Form_Spec_CHK_Bottom
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(903, 1012)
        Me.Controls.Add(Me.PicBox_Main)
        Me.Controls.Add(Me.PicBox_Comp)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form_Spec_CHK_Bottom"
        Me.Text = "Form_Spec_CHK_Bottom"
        CType(Me.PicBox_Main, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBox_Comp, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PicBox_Main As System.Windows.Forms.PictureBox
    Friend WithEvents PicBox_Comp As System.Windows.Forms.PictureBox
End Class
