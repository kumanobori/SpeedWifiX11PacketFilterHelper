Attribute VB_Name = "packetFilterJsGenerator"
Option Explicit


Const ROW_ATTRIBUTE_NAME = 2 ' �s�ԍ��F������
Const ROW_ENABLE = 3 ' �s�ԍ��F�L���i������Y�������Ă�����̂̂݁A�o�̓R�[�h�����ρj
Const ROW_DATA_START = 21 ' �s�ԍ��F�f�[�^�J�n

Const COL_INDEX = 1 ' ��ԍ��F�o�͔ԍ��i���̗񂪋�ɂȂ�܂ŏ�������j
Const COL_FLG_OUTPUT = 2 ' ��ԍ��F�o�̓t���O�i���̗�Y�̏ꍇ�̂ݏo�͂���j
Const COL_ATTRIBUTE_START = 3 ' ��ԍ��F�����J�n�i���̗񂩂�A�o�̓f�[�^�Ɋւ�����e���L�q����Ă�����̂Ƃ݂Ȃ��j

Const EOL = vbLf

Dim log As New Logger

Sub exec()
    
    ' ���s�J�n���_�̃A�N�e�B�u�u�b�N�ƃA�N�e�B�u�V�[�g���擾
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


' �o�̓t�@�C���̃p�X�𐶐�����B
' �o�͌����[�N�u�b�N�Ɠ����p�X�ɁA���s�������t�@�C�����Ƃ������̂Ƃ���B
Private Function generateOutputPath(wb As Workbook) As String
    
    Dim wbPath As String: wbPath = wb.Path
    Dim timestamp As String: timestamp = Format(Now, "yyyymmdd_hhmmss")
    Dim fullpath As String: fullpath = wbPath & "\" & "generated_script_" & timestamp & ".txt"
    generateOutputPath = fullpath
    
End Function


' �X�N���v�g�𐶐�����
' @param ws �Ώۃf�[�^�̂��郏�[�N�V�[�g
' @param outputPath ��������X�N���v�g�̃t�@�C���p�X
Private Function generateScript(ws As Worksheet, outputPath As String)
    
    Dim logPrefix As String: logPrefix = "generateScript: "
    log.logDebug (logPrefix & "start. outputPath = " & outputPath)
    
    ' �X�N���v�g���i�[����ϐ�
    Dim scriptAll As String: scriptAll = ""
    
    ' �f�[�^�J�n�s���Z�b�g
    Dim rowWk As Long: rowWk = ROW_DATA_START
       
    
    ' INDEX�s����ɂȂ�܂ŁA�s�P�ʂ̏������s��
    Do While ws.Cells(rowWk, COL_INDEX).value <> ""
        Dim scriptForRow As String: scriptForRow = generateOneRule(ws, rowWk)
        scriptAll = scriptAll & scriptForRow
        rowWk = rowWk + 1
    Loop
    
        
    ' ���������X�N���v�g���o��
    Open outputPath For Append As #1
    Print #1, scriptAll
    Close #1
    ' log.logDebug (logPrefix & EOL & scriptAll)
    
    
    log.logDebug (logPrefix & "end.")
    
End Function


' 1�s���̃X�N���v�g�𐶐�����
' @param ws �Ώۃf�[�^�̂��郏�[�N�V�[�g
' @param row �����Ώۂ̍s�ԍ�
' @return string 1�s���̃X�N���v�g
Private Function generateOneRule(ws As Worksheet, row As Long) As String
    Dim logPrefix As String: logPrefix = "generateOneRule: "
    log.logDebug (logPrefix & "start. row = " & row)
    
    ' �o�̓t���O��Y���ݒ肳��ĂȂ��ꍇ�͂Ȃɂ����Ȃ�
    If ws.Cells(row, COL_FLG_OUTPUT) <> "Y" Then
        log.logDebug ("output flag is not set.")
        Exit Function
    End If
    
    
    ' �X�N���v�g���i�[����ϐ�
    Dim scriptForRow As String: scriptForRow = ""
    
    ' �����J�n����Z�b�g
    Dim colWk As Long: colWk = COL_ATTRIBUTE_START
    
    ' �����񂪋�ɂȂ�܂ŁA��P�ʂ̏������s��
    Do While ws.Cells(ROW_ATTRIBUTE_NAME, colWk) <> ""
        Dim scriptForColumn As String: scriptForColumn = generateOneAttribute(ws, row, colWk)
        scriptForRow = scriptForRow & scriptForColumn
        colWk = colWk + 1
    Loop
    
    ' �t�H�[��submit�̃X�N���v�g��ǉ�
    scriptForRow = _
    "/* -------------- */" _
    & EOL & scriptForRow _
    & EOL & "/* ���M */" _
    & EOL & "document.getElementById('Apply_1').click();" _
    & EOL _
    & EOL
    ' & EOL & "var sleep = waitTime => new Promise( resolve => setTimeout(resolve, 10000) );" _

    ' generateOneRule = scriptForRow & EOL
    generateOneRule = Replace(scriptForRow, vbLf, "") & EOL
End Function


' 1���ڕ��̃X�N���v�g�𐶐�����
' @param ws �Ώۃf�[�^�̂��郏�[�N�V�[�g
' @param row �Ώۍs
' @param col �Ώۗ�
' @return string 1���ڕ��̃X�N���v�g
Private Function generateOneAttribute(ws As Worksheet, row As Long, col As Long) As String
    
    Dim logPrefix As String: logPrefix = "generateOneAttribute: "
    log.logDebug (logPrefix & "start. row = " & row & ", col = " & col)
    
    Dim val As String: val = ws.Cells(row, col).value
    ' �󗓂Ȃ�Ȃɂ��������Ȃ�
    If val = "" Then
        log.logDebug (logPrefix & "blank. nothing to generate.")
        Exit Function
    End If
    ' �������Ȃ�Ȃɂ��������Ȃ�
    If ws.Cells(ROW_ENABLE, col).value <> "Y" Then
        log.logDebug (logPrefix & "not enabled. nothing to generate.")
        Exit Function
    End If
    
    
    ' ���ږ��̃Z���̃R�����g�ɂ���X�N���v�g�e���v���[�g���擾����
    Dim scriptTemplate As String: scriptTemplate = ws.Cells(ROW_ATTRIBUTE_NAME, col).comment.Text
    
    ' �l��Y�̂Ƃ��̓e���v���[�g���̂܂܁A����ȊO�̂Ƃ��͒u������
    Dim script As String
    If val = "Y" Then
        script = scriptTemplate
    Else
        script = Replace(scriptTemplate, "%s", val)
    End If
    
    ' ���ږ����R�����g�Ƃ���
    Dim comment As String: comment = ws.Cells(ROW_ATTRIBUTE_NAME, col).value
    
    ' �S�̂�g�ݍ��킹
    Dim result As String: result = "/* " & comment & " */" & EOL & script & EOL
    
    log.logDebug (logPrefix & "generated: " & EOL & result)
    generateOneAttribute = result
    
End Function

