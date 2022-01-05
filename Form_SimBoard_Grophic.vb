Public Class Form_SimBoard_Grophic

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        If (e.X <= 0) Then
            Return
        End If
        Form_Main.Label46.Text = e.X & " " & e.Y

        'Form_Main.Label47.Text = Form_Main.Y_.bufs(e.X)
        Form_Main.Label48.Text = Form_Main.FAN_.bufs(e.X)

        Form_Main.Label59.Text = Form_Main.SPH1_.bufs(e.X)
        Form_Main.Label60.Text = Form_Main.SPL1_.bufs(e.X)
        Form_Main.Label61.Text = Form_Main.SPH2_.bufs(e.X)
        Form_Main.Label62.Text = Form_Main.SPL2_.bufs(e.X)

        'If (Form1.ES34_PWM1.buf(e.X) = 0) Then
        'Label4.Text = Str(0) & " 轉速:0"
        'Else
        Label6.Text = Str(Form_Main.BLCB1_.bufs(e.X)) & "%"
        Label8.Text = Str(Form_Main.BLCB2_.bufs(e.X)) & "%"
        'End If

        If (Form_Main.LB_ES34_tp.Text = "S") Then

            Label4.Text = Form_Main.ES34_SPI1_FB.bufs(e.X)
            Label5.Text = Form_Main.ES34_SPI2_FB.bufs(e.X)

        Else

            If (Form_Main.ECC1_.bufs(e.X)) Then
                'Form_Main.Label53.Text = "c%:" & Str(Form_Main.ECC1_.bufs(e.X)) & Str((((Form_Main.ECC1_.bufs(e.X) - 20) / 80) * 5000) + 2000)
                'Form_Main.Label53.Text = "c%:" & (RPM_to_Ratio(Form_Main.ECC1_.bufs(e.X))).ToString("##") ' & "," & Str(Form_Main.ECC1_.bufs(e.X))    '[20170413]
                Form_Main.Label53.Text = "c%:" & (((Form_Main.ECC1_.bufs(e.X) - 800) / 97.5) + 10).ToString("00") & "," & Str(Form_Main.ECC1_.bufs(e.X))    '[20170413]
                '(((Form_Main.ECC1_.bufs(e.X) - 800) / 97.5) + 10).ToString("00")
            Else
                Form_Main.Label53.Text = "c%:0, (0)"
            End If

            If (Form_Main.ECC2_.bufs(e.X)) Then
                'Form_Main.Label54.Text = "c%:" & Str(Form_Main.ECC2_.bufs(e.X)) & Str((((Form_Main.ECC2_.bufs(e.X) - 20) / 80) * 5000) + 2000)
                'Form_Main.Label54.Text = "c%:" & (RPM_to_Ratio(Form_Main.ECC2_.bufs(e.X))).ToString("##") ' & "," & Str(Form_Main.ECC2_.bufs(e.X))    '[20170413]
                Form_Main.Label54.Text = "c%:" & (((Form_Main.ECC2_.bufs(e.X) - 800) / 97.5) + 10).ToString("00") & "," & Str(Form_Main.ECC2_.bufs(e.X))    '[20170413]
            Else
                Form_Main.Label54.Text = "c%:0, (0)"
            End If


            If (Form_Main.ES34_PWM1.bufs(e.X) = 0) Then
                Label4.Text = Str(0) & "% 轉速:0"
            Else
                Label4.Text = Str(Form_Main.ES34_PWM1.bufs(e.X)) & "% 轉速:" & Str(800 + (Form_Main.ES34_PWM1.bufs(e.X) - 10) * 97.5)
            End If


            If (Form_Main.ES34_PWM2.bufs(e.X) = 0) Then
                Label5.Text = Str(0) & "% 轉速:0"
            Else
                Label5.Text = Str(Form_Main.ES34_PWM2.bufs(e.X)) & "% 轉速:" & Str(800 + (Form_Main.ES34_PWM2.bufs(e.X) - 10) * 97.5)
            End If

        End If




        Form_Main.eXX = e.X
        Form_Main.eYY = e.Y
    End Sub
    Private Function Ratio_to_RPM(ByVal PWMr As Integer)
        Return (((PWMr - 20) / 80) * 5000) + 2000
    End Function
    Private Function RPM_to_Ratio(ByVal PWMr As Integer)
        Return (((PWMr - 800) / 97.5)) + 10
    End Function

    Private Sub Form_SimBoard_Grophic_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub


End Class