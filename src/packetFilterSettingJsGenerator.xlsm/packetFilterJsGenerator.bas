Attribute VB_Name = "packetFilterJsGenerator"
Option Explicit


Const ROW_ATTRIBUTE_NAME = 2 ' 行番号：属性名
Const ROW_ENABLE = 3 ' 行番号：有効（ここにYが入っているもののみ、出力コード実装済）
Const ROW_DATA_START = 21 ' 行番号：データ開始

Const COL_INDEX = 1 ' 列番号：出力番号（この列が空になるまで処理する）
Const COL_FLG_OUTPUT = 2 ' 列番号：出力フラグ（この列がYの場合のみ出力する）
Const COL_ATTRIBUTE_START = 3 ' 列番号：属性開始（この列から、出力データに関する内容が記述されているものとみなす）

Const EOL = vbLf

Dim log As New Logger

Sub exec()
    
    ' 実行開始時点のアクティブブックとアクティブシートを取得
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    Dim outputPath As String
    outputPath = generateOutputPath(wb)
    
    Call log.init(log.DESTINATION_DEBUG_PRINT, log.LEVEL_DEBUG)
    log.logInfo ("")
    log.logInfo ("------------------------------------------")
    Call generateScript(ws, outputPath)
    log.logInfo ("------------------------------------------")
    
    
End Sub


' 出力ファイルのパスを生成する。
' 出力元ワークブックと同じパスに、実行日時をファイル名としたものとする。
Private Function generateOutputPath(wb As Workbook) As String
    
    Dim wbPath As String: wbPath = wb.Path
    Dim timestamp As String: timestamp = Format(Now, "yyyymmdd_hhmmss")
    Dim fullpath As String: fullpath = wbPath & "\" & "generated_script_" & timestamp & ".txt"
    generateOutputPath = fullpath
    
End Function


' スクリプトを生成する
' @param ws 対象データのあるワークシート
' @param outputPath 生成するスクリプトのファイルパス
Private Function generateScript(ws As Worksheet, outputPath As String)
    
    Dim logPrefix As String: logPrefix = "generateScript: "
    log.logDebug (logPrefix & "start. outputPath = " & outputPath)
    
    ' スクリプトを格納する変数
    Dim scriptAll As String: scriptAll = ""
    
    ' データ開始行をセット
    Dim rowWk As Long: rowWk = ROW_DATA_START
       
    
    ' INDEX行が空になるまで、行単位の処理を行う
    Do While ws.Cells(rowWk, COL_INDEX).value <> ""
        Dim scriptForRow As String: scriptForRow = generateOneRule(ws, rowWk)
        scriptAll = scriptAll & scriptForRow
        rowWk = rowWk + 1
    Loop
    
        
    ' 生成したスクリプトを出力
    Open outputPath For Append As #1
    Print #1, scriptAll
    Close #1
    ' log.logDebug (logPrefix & EOL & scriptAll)
    
    
    log.logDebug (logPrefix & "end.")
    
End Function


' 1行分のスクリプトを生成する
' @param ws 対象データのあるワークシート
' @param row 生成対象の行番号
' @return string 1行分のスクリプト
Private Function generateOneRule(ws As Worksheet, row As Long) As String
    Dim logPrefix As String: logPrefix = "generateOneRule: "
    log.logDebug (logPrefix & "start. row = " & row)
    
    ' 出力フラグにYが設定されてない場合はなにもしない
    If ws.Cells(row, COL_FLG_OUTPUT) <> "Y" Then
        log.logDebug ("output flag is not set.")
        Exit Function
    End If
    
    
    ' スクリプトを格納する変数
    Dim scriptForRow As String: scriptForRow = ""
    
    ' 属性開始列をセット
    Dim colWk As Long: colWk = COL_ATTRIBUTE_START
    
    ' 属性列が空になるまで、列単位の処理を行う
    Do While ws.Cells(ROW_ATTRIBUTE_NAME, colWk) <> ""
        Dim scriptForColumn As String: scriptForColumn = generateOneAttribute(ws, row, colWk)
        scriptForRow = scriptForRow & scriptForColumn
        colWk = colWk + 1
    Loop
    
    ' フォームsubmitのスクリプトを追加
    scriptForRow = _
    "/* -------------- */" _
    & EOL & scriptForRow _
    & EOL & "/* 送信 */" _
    & EOL & "document.getElementById('Apply_1').click();" _
    & EOL _
    & EOL
    ' & EOL & "var sleep = waitTime => new Promise( resolve => setTimeout(resolve, 10000) );" _

    ' generateOneRule = scriptForRow & EOL
    generateOneRule = Replace(scriptForRow, vbLf, "") & EOL
End Function


' 1項目分のスクリプトを生成する
' @param ws 対象データのあるワークシート
' @param row 対象行
' @param col 対象列
' @return string 1項目分のスクリプト
Private Function generateOneAttribute(ws As Worksheet, row As Long, col As Long) As String
    
    Dim logPrefix As String: logPrefix = "generateOneAttribute: "
    log.logDebug (logPrefix & "start. row = " & row & ", col = " & col)
    
    Dim val As String: val = ws.Cells(row, col).value
    ' 空欄ならなにも生成しない
    If val = "" Then
        log.logDebug (logPrefix & "blank. nothing to generate.")
        Exit Function
    End If
    ' 未実装ならなにも生成しない
    If ws.Cells(ROW_ENABLE, col).value <> "Y" Then
        log.logDebug (logPrefix & "not enabled. nothing to generate.")
        Exit Function
    End If
    
    
    ' 項目名のセルのコメントにあるスクリプトテンプレートを取得する
    Dim scriptTemplate As String: scriptTemplate = ws.Cells(ROW_ATTRIBUTE_NAME, col).comment.Text
    
    ' 値がYのときはテンプレートそのまま、それ以外のときは置換する
    Dim script As String
    If val = "Y" Then
        script = scriptTemplate
    Else
        script = Replace(scriptTemplate, "%s", val)
    End If
    
    ' 項目名をコメントとする
    Dim comment As String: comment = ws.Cells(ROW_ATTRIBUTE_NAME, col).value
    
    ' 全体を組み合わせ
    Dim result As String: result = "/* " & comment & " */" & EOL & script & EOL
    
    log.logDebug (logPrefix & "generated: " & EOL & result)
    generateOneAttribute = result
    
End Function

