Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form4
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Form4上の”実施済”クリック
        calact = 1 'open/short実施済みフラグ設定

        'フォームを一度実施したフラグ設定
        'done4 = 1

        Me.Dispose()
        'Me.Hide()  'Form4をHide
        Form2.Show()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Form4上のButton2（”未実施”）クリック
        calact = 0 '計測器コンディション確認未実施フラグ設定

        Me.Hide()

        MsgBox("選択したチップサイズ容量" & Mid$(SELSHEETNAME, 1, 5) & "のQA-STD容量値" & HI_CRead & "での" & vbCrLf &
               "計測器確認を実施してください。" & vbCrLf & vbCrLf &
               "このままにして4278Aのリモート表示部のLDLボタンを押してコンディション確認してください。" & vbCrLf & vbCrLf &
               "計器器校正後作製した時のC値との差分が0.02pF以下を確認後、OKボタン を押して測定開始します。")
        '計測器コンディション確認実施OK後に確認未実施フラグ設定を1にする（計測器コンディションok）
        calact = 1

        'MsgBox("作成するQAスタンダード容量での計器コンディション確認を" & vbCrLf & vbCrLf &
        '       "実施してください。" & vbCrLf & vbCrLf &
        '       "R=0.05以下である事！")

        'Form4 実施確認フォームをHIDEにする
       　' Me.Hide()


        ' Call Form1.Button2_Click(sender, e)

        Exit Sub

        'Form2.Show()



    End Sub

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Label1.Text = ""
        calact = 0 '初期値は　open/short未実施
        'Me.Label1.Text = "選択チップ " & Mid$(SELSHEETNAME, 1, 5) & " のOPEN/SHORT"
        Me.Label1.Text = "選択したチップサイズ容量" & Mid(SELSHEETNAME, 1, 5) & "の " & HI_CRead & "の"

        Me.Label2.Text = "QA-STD容量値での計測器確認”
        Me.Label3.Text = "作製当時C値との差分が0.02pF以下”

        Me.Button2.Enabled = True
        Me.Button2.Visible = True

        'ツールチップにてボタン２の使用無を説明する
        ToolTip1.ShowAlways = True

        ' ToolTip1.SetToolTip(Button1, "このボタンを押してから測定器のリモート表示LEDにあるLDLボタンを " & vbCrLf &
        '                               "押して計測器コンディション確認します。" & vbCrLf & vbCrLf &
        '                               "確認時の差分C値が0.02pF以下を確認後、継続して" & vbCrLf &
        '                               "測定ボタンでの測定を実施します。")

        'ToolTip1.SetToolTip(Button2, "「未だ確認してません」ボタンは使用禁止です。 " & vbCrLf &
        '                             "ボタンを押さずに計測器コンディション確認可能です。")

    End Sub


End Class