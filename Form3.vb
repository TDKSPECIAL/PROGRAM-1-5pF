Public Class Form3
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        'C1005
        SELSHEETNAME = "C1005STD"
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            RadioButton3.Checked = False
            RadioButton4.Checked = False
            RadioButton5.Checked = False
        End If
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles Me.Load
        'ラジオボタン初期化
        RadioButton1.Checked = False 'C0603STD
        RadioButton2.Checked = True  'C1005STDを初期値設定
        RadioButton3.Checked = False 'C1608STD
        RadioButton4.Checked = False 'C2012STD
        RadioButton5.Checked = False 'C3216STD
        'SELSHEETNAME = ""  '初期値　“”　空欄を下記に修正
        '初期値として"C1005STD”’2023年8月23日初期値追加とラジオボタン初期設定チェック（C1005STD)
        SELSHEETNAME = "C1005STD"  '初期値　“C1005STD”

        'MsgBox("スタンダードチップ作成の" & vbCrLf & vbCrLf &
        '       "チップサイズを選択してください。”)
        MsgBox("作成するQAスタンダードの" & vbCrLf & vbCrLf &
               "チップサイズを選択してください。”)
        ' MsgBox("測定開始するチップサイズでの " & vbCrLf & vbCrLf &
        '        "OPEN/SHORT補正を必ず測定前に" & vbCrLf & vbCrLf &
        '        "実施してください。")


        '2023年9月1日ツールチップテキスト追加
        ToolTip1.ShowAlways = True

        ToolTip1.SetToolTip(Button1, "どれか1つのチップサイズを選択してから" & vbCrLf & vbCrLf &
                                                " 選択実行ボタンを押してください")
        ToolTip1.SetToolTip(RadioButton1, "C0603を選択時チェックします。 ")
        ToolTip1.SetToolTip(RadioButton2, "C1005を選択時チェックします。 ")
        ToolTip1.SetToolTip(RadioButton3, "C1608を選択時チェックします。 ")
        ToolTip1.SetToolTip(RadioButton4, "C2012を選択時チェックします。 ")
        ToolTip1.SetToolTip(RadioButton5, "C3216を選択時チェックします。 ")




    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        'C0603
        If RadioButton1.Checked = True Then
            SELSHEETNAME = "C0603STD"
            RadioButton2.Checked = False
            RadioButton3.Checked = False
            RadioButton4.Checked = False
            RadioButton5.Checked = False
        End If

    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        'C1608
        If RadioButton3.Checked = True Then
            SELSHEETNAME = "C1608STD"
            RadioButton1.Checked = False
            RadioButton2.Checked = False
            RadioButton4.Checked = False
            RadioButton5.Checked = False
        End If
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        'C2012
        If RadioButton4.Checked = True Then
            SELSHEETNAME = "C2012STD"
            RadioButton1.Checked = False
            RadioButton2.Checked = False
            RadioButton3.Checked = False
            RadioButton5.Checked = False
        End If
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        'C3216
        If RadioButton5.Checked = True Then
            SELSHEETNAME = "C3216STD"
            RadioButton1.Checked = False
            RadioButton2.Checked = False
            RadioButton3.Checked = False
            RadioButton4.Checked = False
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '選択ラジオボタンのチップMSGではあえて表示しない
        '  MsgBox("SELECT " & SELSHEETNAME & " SHEET!!")
        Me.Hide() 'FORM3 CLOSE

        ' MsgBox("測定開始前にOPEN/SHORT補正を必ず実施してください。” & vbCrLf & vbCrLf &
        '       "OPEN/SHORT補正実施後、エクセルシートを開き" & vbCrLf & vbCrLf &
        '      "測定を開始して下さい。")
        MsgBox("測定開始前に" & Mid(SELSHEETNAME, 1, 5) & "のOPEN/SHORT補正を" & vbCrLf & vbCrLf &
               "必ず実施してください。” & vbCrLf & vbCrLf &
               "OPEN/SHORT補正実施後、エクセルシートを開き" & vbCrLf & vbCrLf &
               "測定を開始して下さい。")
        Form1.Show()

    End Sub


End Class