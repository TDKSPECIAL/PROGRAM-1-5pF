Imports System.CodeDom
Imports System.IO
Imports System.Linq.Expressions
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel



Public Class Form1
    'Dim strFileName As String = "C:\ABC\DATA1-5\ANYCSIZE-3M-6M-1-5pF-5DG.xls"
    Dim strFileName As String = "C:\ABC\DATA1-5\QA-SUB-STD-MAKE-5TIMES-MEASUREMENT230825.xlsm"
    Dim addkakuchosi As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Error処理追加　2022年10月25日
        On Error GoTo ErrorHandler


        'Excelアプリケーション起動
        xlsApplication = New Excel.Application
        'ExcelのWorkbooks取得
        xlsWorkbooks = xlsApplication.Workbooks

        'Dim sh As Microsoft.Office.Interop.Excel.Worksheet
        'xlsWorkSheets = xlsApplication.Worksheets   '**************

        'Excel Visible = true:表示,Visible = false:非表示
        xlsApplication.Visible = True
        xlsApplication.DisplayAlerts = False

        '*****************************************************************
        '既存 Excel ファイルを開く
        xlsWorkbook = xlsWorkbooks.Open(strFileName)
        '*****************************************************************
        'Excel の Worksheets 取得

        '*********************************************************************************************
        '指定したエクセルシートで開く処理
        'xlsWorkSheet = xlsWorkbook.Worksheets("C1608STD")  '**********************TEST時必要
        'xlsWorkSheet = xlsWorkbook.Worksheets("C1005STD")  '**********************TEST時必要
        'xlsWorkSheet = xlsWorkbook.Worksheets("C2012STD")  '**********************TEST時必要
        'xlsWorkSheet = xlsWorkbook.Worksheets("C3216STD")  '**********************TEST時必要
        xlsWorkSheet = xlsWorkbook.Worksheets(SELSHEETNAME)  '選択したシートを開く

        'xlsWorkSheet = xlsWorkbook.Worksheets("CAPACITOR-1-5pF")
        'xlsWorkSheet = xlsWorkbook.Worksheets("CAPACITOR-1-5pF②")
        'xlsWorkSheet = xlsWorkbook.Worksheets("CAPACITOR-1-5pF③")
        'xlsWorkSheet = xlsWorkbook.Worksheets("CAPACITOR-1-5pF②")

        xlsWorkSheet.Activate()  '*************************************************TEST時必要
        '*********************************************************************************************
        '

        'xlsWorkSheets = xlsWorkbook.Worksheets '************************************ORIGINAL
        'Excel の Worksheet 取得
        'xlsWorkSheet = CType(xlsWorkSheets.Item(1), Excel.Worksheet) '**************ORIGINAL
        '*********************************************************************************************

        'xlsWorkSheet.Visible = True
        'シート名称
        'xlsWorkSheet.Name = "シート名test"
        'セル選択
        '    xlsRange = xlsWorkSheet.Range("A1")
        'セルに値設定
        '   xlsRange.Value = "TEST123"

        '*******************************
        ' Public CSIZE As String
        ' Public CAPA As String
        ' Public LIMITMONTH As String
        ' Public KOSU As String
        '*******************************
        'CSIZE = xlsWorkSheet.Application.Cells(1, 1).ToString
        '    CSIZE = xlsWorkSheet.Application.Range("V8").Value.ToString  '"C0603"
        '    Debug.WriteLine(CSIZE)
        '    CAPA = xlsWorkSheet.Application.Range("X8").Value.ToString   '"5pF"
        '    Debug.WriteLine(CAPA)
        '    LIMITMONTH = xlsWorkSheet.Application.Range("U8").Value.ToString  '"3M" or "6M"
        '    Debug.WriteLine(LIMITMONTH)
        '原図を修正してから読取対応検討
        'YEAR-MONTH-DAY 用変数　NENGAPPI
        '    R_NENGAPPI = xlsWorkSheet.Application.Range("S6").Value.ToString '"2022/09/13"
        '    NENGAPPI = Mid(R_NENGAPPI, 1, 10)
        '    Debug.WriteLine(NENGAPPI)

        '-------------------------- FILE用年月日再構成---------------------------------
        '"2022/09/22"から　"2022","09","22"を取り出し再構成で　"20220922"とする
        '     Yearda = Mid(NENGAPPI, 1, 4)
        '     monthda = Mid(NENGAPPI, 6, 2)
        '     dayda = Mid(NENGAPPI, 9, 2)

        '    MATOMENENN = Yearda & monthda & dayda '20220922
        '    Debug.WriteLine(MATOMENENN)
        '------------------------------------------------------------------------------

        'IRAIMOTOはダイアログを開いたときに読み込む
        '    IRAIMOTO = xlsWorkSheet.Application.Range("AA8").Value.ToString '依頼元設定
        '    Debug.WriteLine(IRAIMOTO)

        '個数を出すルーチンから
        '    KOSU = xlsWorkSheet.Application.Range("Z8").Value.ToString '5
        '    Debug.WriteLine(KOSU)

        '最終のファイル構成もダイアログを開いたときに読み込み構成する
        'D_FILENAME 拡張子無しのファイル名のみ
        '    D_FILENAME = CSIZE & "-" & CAPA & "-" & LIMITMONTH & "-" & KOSU & "PCS" & "-" & IRAIMOTO & "-" & MATOMENENN
        '    Debug.WriteLine(D_FILENAME)


        Me.Hide()
        '********************************************
        'Form2表示へ移る
        '(エクセルの指定セルに値を入れる為の測定フォーム）

        Form2.ShowDialog()

        MRComObject(xlsRange)

        '//////////////////////////////////////////////////////////////////////////////
        '********************************************
        xlsApplication.DisplayAlerts = False

        '********************************************
        '保存ダイアログを開く（ボタン3のルーチン）
        'Call Button3_Click(sender, e)
        Call D_open()

        'addkakuchosi = objSFD.FileName & ".xlsx"
        addkakuchosi = objSFD.FileName '拡張子付の保存構成ファイル名

        xlsWorkbook.SaveAs(Filename:=addkakuchosi, FileFormat:=Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)

        xlsApplication.DisplayAlerts = True
        '//////////////////////////////////////////////////////////////////////////////

        '保存時の問合せダイアログを非表示に設定
        ' xlsApplication.DisplayAlerts = False
        'ファイルに保存 (Excel 2007～ブック形式)
        'xlsWorkbook.SaveAs(Filename:=strFileName, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbook)
        '保存時の問合せダイアログを表示に戻す
        'xlsApplication.DisplayAlerts = True

        '終了処理
        'xlsWorkSheet の解放
        MRComObject(xlsWorkSheet)
        'xlsWorkSheets の解放
        MRComObject(xlsWorkSheets)
        'xlsWorkbookを閉じる
        xlsWorkbook.Close(False)
        'xlsWorkbook の解放
        MRComObject(xlsWorkbook)
        'xlsWorkbooks の解放
        MRComObject(xlsWorkbooks)
        'Excelを閉じる 
        xlsApplication.Quit()
        'xlsApplication を解放
        MRComObject(xlsApplication)

        'End 処理ボタンコールして終了
        Call Button2_Click(sender, e)

ErrorHandler:
        Select Case Err.Number
            Case 1004
                MsgBox("errorNo.= " & Err.Number & "エラーが発生しました。" & vbCrLf &
                       "エクセルのプラットフォームまで開けましたが、下記の場所に読み込む” & vbCrLf &
                       "エクセル原図ファイルが有りませんでした。" &
                       vbCrLf & vbCrLf &
                       "パソコン階層フォルダ　➡　C:\ABC\DATA1-5" &
                       vbCrLf & vbCrLf &
                       "上記階層フォルダ内に ANYCSIZE-3M-6M-1-5pF-5DG.xls " & vbCrLf & vbCrLf &
                       "エクセルファイル原図の設置保管を確認してください。" & vbCrLf & vbCrLf &
                       "このエラーメッセージボックスを閉じてから、保管処理解決ののち" & vbCrLf &
                       "このままSUB-STD測定プログラムを継続使用できます。" & vbCrLf & vbCrLf &
                       "エクセル原図保管設置OK後、エクセルOPEN測定開始ボタンで" & vbCrLf &
                       "処理再開できます。")

                '***************************************************************
                '開いたところまでのbookをメモリから開放する処理
                MRComObject(xlsWorkbooks)
                'Excelを閉じる 
                xlsApplication.Quit()
                'xlsApplication を解放
                MRComObject(xlsApplication)
                '***************************************************************
                Exit Sub

            Case -2147221487
                MsgBox("erroNo.= " & Err.Number & "エラーが発生しました。" & vbCrLf &
                       "エラー内容:" & Err.Description &
                       vbCrLf & vbCrLf &
                       "アプリ指定のGPIBインターフェースが一致していません。" & vbCrLf &
                       "又はアドレス指定が異なったアドレス状態です。" &
                       vbCrLf & vbCrLf &
                       "Keysight Connection Expert 2022でインターフェースが" & vbCrLf &
                       "接続出来る事を確認して下さい。又、GPIBアドレス設定も17である" & vbCrLf &
                       "事を確認して下さい。" &
                       vbCrLf & vbCrLf &
                       "アプリは一旦終了します。エクセルを閉じる場合は　保存しない " & vbCrLf &
                       "を選んで終了してください。")

                '***************************************************************
                '開いたところまでのbookをメモリから開放する処理
                MRComObject(xlsWorkbooks)
                'Excelを閉じる 
                xlsApplication.Quit()
                'xlsApplication を解放
                MRComObject(xlsApplication)
                '***************************************************************
                Exit Sub

            Case Else
                MsgBox("errorNo.= " & Err.Number & " " & Err.Description)

                '***************************************************************
                '開いたところまでのbookをメモリから開放する処理
                MRComObject(xlsWorkbooks)
                'Excelを閉じる 
                xlsApplication.Quit()
                'xlsApplication を解放
                MRComObject(xlsApplication)
                '***************************************************************
                Exit Sub

        End Select

    End Sub

    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Private -> Public
        End
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        ToolTip1.ShowAlways = True

        ToolTip1.SetToolTip(Label1, "選択実施したチップサイズです。")

        ToolTip1.SetToolTip(Button1, "データ記録用エクセルファイルを開いて" & vbCrLf &
                                     "【QAスタンダードチップ作製】の測定を行ないます。")
        ToolTip1.SetToolTip(Button2, "QA-SUB-STD作製用測定プログラムを終了します。")

        Me.Label1.Text = Mid(SELSHEETNAME, 1, 5)

        done4 = 0 '初期状態開始時のオープンショート確認フォームは最初は表示




    End Sub
End Class
