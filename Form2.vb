Imports Ivi.Visa.Interop
Imports Microsoft.Office.Interop
Imports System.Diagnostics.Eventing.Reader
Imports System.Resources

Public Class Form2
    Dim i As Integer

    Dim myPos As Integer '4278ACdとDのデータ値区切り","位置
    Dim measdata As String 'GPIB測定データ
    'Public GPIBAD As String
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        ' Me.Hide() 'Form2をHide 'これを入れるとForm2を抜けてしまう

        '2023年8月29日追加のForm4フォームの処理でオープンショート確認済みの場合は
        '2回目からは表示なし

        richi = xlsApplication.ActiveCell.Column '列
        cichi = xlsApplication.ActiveCell.Row    '行
        TextBox4.Text = richi.ToString  'FORM2上テキストに表示
        TextBox5.Text = cichi.ToString　'FORM2上テキストに表示

        HI_CRead = "" '入れ子初期化 値　"1pF","50pF","300pF" e.t.cを設定する
        HI_CSet = ""　'入れ子初期化

        AUTORSET = RTrim(richi)

        Select Case AUTORSET　'列の位置を数値化
            Case 2　　　　　　　　'2列目（Bの位置）
                HI_CRead = "1pF" '1pFにセットする
                'LENGTH = 3
            Case 4
                HI_CRead = "2pF"
                'LENGTH = 3
            Case 6
                HI_CRead = "5pF"

                'LENGTH = 3
            Case 8
                HI_CRead = "10pF"

                'LENGTH = 4
            Case 10
                HI_CRead = "15pF"

                'LENGTH = 4
            Case 12
                HI_CRead = "50pF"

                'LENGTH = 4
            Case 14
                HI_CRead = "100pF"

                'LENGTH = 5
            Case 16
                HI_CRead = "150pF"

                'LENGTH = 5
            Case 18
                HI_CRead = "200pF"

                'LENGTH = 5
            Case 20
                HI_CRead = "250pF"

                'LENGTH = 5
            Case 22
                HI_CRead = "300pF"  '1000pFも300pFレンジ

                'LENGTH = 5
            Case Else
                MsgBox("Something ERROR Happen!")
        End Select








        OPSDISP = RTrim(cichi)  '13～17までの内13の時Form4を表示　14から17は非表示
        Debug.WriteLine(cichi)
        Select Case OPSDISP
            Case 13
                ' done4 = 0
                calact = 0
            Case 14
                'done4 = 1
                calact = 1
            Case 15
                'done4 = 1
                calact = 1
            Case 16
                'done4 = 1
                calact = 1
            Case 17
                'done4 = 1
                calact = 1
            Case Else
        End Select

        'If done4 = 0 Then
        If calact = 0 Then
            Form4.ShowDialog()
            'If calact = 0 Then
            'MRComObject(xlsWorkSheets)
            'MRComObject(xlsWorkSheets)
            'xlsWorkbook.Close()
            'MRComObject(xlsWorkbook)
            'MRComObject(xlsWorkbooks)
            'xlsApplication.Quit()
            'MRComObject(xlsApplication)


            'MsgBox(”一旦測定用フォームとエクセルを全て閉じます。" & vbCrLf & vbCrLf &
            '              "対象チップサイズの対象容量で4278Aコンディションを確認して下さい。" & vbCrLf & vbCrLf &
            '              "最初からやり直してください。")


            'End

            'Exit Sub
            'End If

            'done4 = 1
            If calact = 1 Then
                'none todo and go to next routine
            End If
        End If

        '****************
        '******************************************
        'Form4　ボタン押下に対応した処理
        'calact = 0 時は計測器コンディション確認を実施する
        'calact = 1時は計測器コンディション確認実施済みOKで次のルーチン実施
        '**********************************************************
        '     If calact = 0 Then
        '        MRComObject(xlsWorkSheets)
        '        MRComObject(xlsWorkSheets)
        '        xlsWorkbook.Close()
        '        MRComObject(xlsWorkbook)
        '        MRComObject(xlsWorkbooks)
        '        xlsApplication.Quit()
        '        MRComObject(xlsApplication)

        '        MsgBox(”一旦測定用フォームとエクセルを全て閉じます。" & vbCrLf & vbCrLf &
        '              "対象チップサイズのOPEN/SHORT補正を実施して下さい。" & vbCrLf & vbCrLf &
        '              "最初からやり直してください。")
        '        End

        '        Exit Sub
        '     End If
        '     If calact = 1 Then
        '        none todo and go to next routine
        '      End If

        MsgBox("■" & HI_CRead & "測定用のC値の書き込み部セルをクリックしましたか？" & vbCrLf &
               "■測定端子にチップコンデンサをセットしましたか？" & vbCrLf &
               "OKで測定します。")





        Dim ioMgr As Ivi.Visa.Interop.ResourceManager
        Dim instrument As Ivi.Visa.Interop.FormattedIO488
        Dim idn As String

        '***************************************************************
        'High Accuracy Mode Auto Setting routine 
        'ADD 2022NOV17
        '
        'MODEFIED 2023AUG23
        'Dim HI_CRead As String
        'Dim HI_CSet As String
        ' HI_CRead = ""
        ' HI_CSet = ""
        'HI_CRead = xlsApplication.Cells(8, 24).Text '"150pF"

        Dim pfsakujo As Integer
        pfsakujo = Len(HI_CRead)
        Select Case pfsakujo
            Case 3
                HI_CSet = Trim(Mid(HI_CRead, 1, 1)) '"1-5pF:ex) 5pF" -> "1"OR"2"OR"5"
                '文字列で高精細時のﾚﾝｼﾞ"RC=" & HI_CRead & "E-12"
                '次のコマンドを構成する
                'instrument.WriteString("RC=" & HI_CSet & "E-12")
                '中味として
                'instrument.WriteString("RC="& "1" & "E-12"
                '下記が構成された事となる
                'instrument.WriteString("RC=1E-12") '1pF高精細レンジをセットする
                '
            Case 4 '以下4桁の場合は2桁の文字変数（数字）
                HI_CSet = Trim(Mid(HI_CRead, 1, 2)) '"10-50pF:ex)50pF" -> "10"OR"15"OR"50"
                '同様にコマンドを構成
                'instrument.WriteString("RC=" & HI_CSet & "E-12")
                '中味として
                'instrument.WriteString("RC=" & "50" & "E-12")
                '下記が構成された事となる
                'instrument.WriteString("RC=50E-12")　’50pF高精細レンジセットする
            Case 5
                HI_CSet = Trim(Mid(HI_CRead, 1, 3)) '"100-300pF:ex)300pF" ->"100"OR"150"OR"200"OR"250"OR"300"
                '３，４と同じなので省略
                '1000pFは300pFレンジと同じなのでなし
                'Case 6
                'HI_CSet = Trim(Mid(HI_CRead, 1, 4)) '"1000pF" -> "1000"
            Case Else
                MsgBox("EXCEL CELL READING ERROR" & vbCrLf &
                       "Check variable richi Value" & vbCrLf &
                       "at Carsol setting location")
        End Select
        '***************************************************************
        'MsgBox(HI_CSet)


        ' Dim GPIBDAT As String 'for automatic checked Visa address so that no need

        ' Dim GPIBAD As String

        'GPIBDAT = "17"
        ' GPIBDAT = Trim(TextBox1.Text)

        'GPIBAD = "GPIB0::" & GPIBDAT & "::INSTR"  '現状のSUBSTD測定用4278A
        'GPIBAD = "GPIB1::" & GPIBDAT & "::INSTR" 'test
        'GPIBAD = "GPIB2::" & GPIBDAT & "::INSTR" '借用の4278A 
        'GPIBAD = pAliaslfExists 'When Form2 load .check Visa Address & Alias

        'GPIBAD = setVisaAddress

        'ioMgr = New Ivi.Visa.Interop.ResourceManager   'OK
        'ioMgr = New ResourceManager                    'NG
        ioMgr = New Ivi.Visa.Interop.ResourceManager   '2023AUG04 fixed

        'instrument = New Ivi.Visa.Interop.FormattedIO488
        instrument = New FormattedIO488
        instrument.IO = ioMgr.Open(GPIBAD)

        '4278A設定
        '
        instrument.WriteString("MPAR1")
        instrument.WriteString("FREQ2")
        instrument.WriteString("OSC=1.0")

        instrument.WriteString("HIAC1") 'SM-11S96 Line 4278A also  setting ok!
        '***************************************************
        instrument.WriteString("RC=" & HI_CSet & "E-12")
        '***************************************************
        'instrument.WriteString("RB0")

        instrument.WriteString("ITIM3")
        instrument.WriteString("DTIM=0")

        instrument.WriteString("AVE=32")

        instrument.WriteString("TRIG1")
        instrument.WriteString("CABL0")


        instrument.WriteString("DATA?")
        idn = instrument.ReadString()

        'MsgBox(idn)  '
        measdata = Trim(idn) '    
        '                                123456789012345678    9
        'for check                      ":DATA +15.0690E+03" & vbLf
        Dim lmeasdata As Integer
        lmeasdata = Len(measdata)

        '*****************************************************
        'デバッグモード設定用　
        superslim = 0      ' Set 1:7555MultiMeter, Set 0:4278A
        '*****************************************************
        '                                123456789012345678    9
        'for check                      ":DATA +15.0690E+03" & vbLf

        Select Case superslim
            Case 1
                sngC = Mid(measdata, 8, 18) '"15.0690E+03"
                TextBox2.Text = sngC
                sngD = 0.001
                TextBox3.Text = sngD
            Case 0
                myPos = InStr(1, measdata, ",", vbTextCompare)　'データ区切り位置
                sngC = Mid(measdata, 1, myPos - 1) 'Cd値抜き取り
                sngC = sngC * 1000000000000.0# 'pF単位に変換処理
                'sngC = 1.00123 'karisettei
                TextBox2.Text = sngC
                sngD = Mid(measdata, myPos + 1) 'D値抜き取り
                sngD = sngD * 100  'for textbox用にD値を　％にする
                'sngD = 0.12 '0.012% karisettei
                TextBox3.Text = sngD
            Case Else
        End Select
        '最初のsngC(Cp)測定値をカーソル指定セルに入れる処理
        xlsRange = xlsWorkSheet.Cells(cichi, richi)
        xlsRange.Value = sngC
        '
        '通常測定時の基準セルの下にD値を入れる処理
        'xlsRange = xlsWorkSheet.Cells(cichi + 1, richi)　通常カーソル位置➡Cp　カーソル位置下➡D
        '
        '今回は基準セルにCp値を入れその右隣のセルにD値を入れる
        '最初のsngD/100(%)値をカーソルの1つ右隣のセルに入れる処理
        'Sub-Std作成シート様にCp、Dの入れる方向をカーソル位置➡Cp　カーソル位置右セル➡Dに書替え
        '下記の対応となる
        ' カーソル置いたセルが基準  Cells(行,列  ）＝　Cells(cichi,richi  )
        '                       　　Cells(行,列+1）＝  Cells(cichi,richi+1)
        'よって下記の位置にD(%)を入れる
        xlsRange = xlsWorkSheet.Cells(cichi, richi + 1)
        '設定したセルにD（％）をいれる
        ' xlsRange.Value = sngD / 100
        xlsRange.Value = sngD

        'tooltipにて説明に変更
        'MsgBox("測定継続→次のセルクリック→コンデンサ入替→測定” & vbCrLf & vbCrLf & "測定終了→「名前を付けて保存」")

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        Form1.Show()

    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles Me.Load

        ToolTip1.ShowAlways = True

        ToolTip1.SetToolTip(Button1, "作製測定終了時は" & vbCrLf & "  [名前を付けて保存処理]ボタン  " & vbCrLf &
                                     "を押してください。")
        ToolTip1.SetToolTip(Button2, "データを入れるセルをクリックしてから " & vbCrLf &
                                     "測定ボタンを押して下さい。" & vbCrLf &
                                     "最初の測定時には測定器のコンディション確認実施の" & vbCrLf &
                                     "の確認フォームが表示されます。")

        '****************************************************************************
        'Automatically Get Visa Address & Visa Alias if any.
        Dim VisaCount As Integer
        VisaCount = 0

        Dim RM = New Ivi.Visa.Interop.ResourceManager
        VisaAdds = RM.FindRsrc("GPIB?*INSTR")
        GPIBAD = ""
        VisaCount = UBound(VisaAdds)


        For i = 0 To UBound(VisaAdds)
            RM.ParseRsrcEx(VisaAdds(i), plnterfaceType, plnterfaceNumber, pSessionType, pUnaliasedExpandedResourceName, pAliaslfExists)
            '
            '
            'Me.TextBox1.Text = "VISA Address = " & VisaAdds(i)
            'Me.TextBox6.Text = "VISA Alias = " & pAliaslfExists

            'Me.TextBox1.Text = VisaAdds(i)
            'Me.TextBox6.Text = pAliaslfExists

        Next
        Me.TextBox1.Text = VisaAdds(0)
        Me.TextBox6.Text = VisaCount.ToString

        'GPIBのVISAアドレスをグローバル変数GPIBADに設定
        'EX)設定VISAアドレス： "GPIB2::17::INSTR"
        GPIBAD = VisaAdds(0) '最終設定Visa Address
        'GPIBAD = VisaAdds(1) '前の設定Visa Address
        RM = Nothing
    End Sub


End Class