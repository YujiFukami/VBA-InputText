Attribute VB_Name = "ModInputText"
Option Explicit

'InputText          ・・・元場所：FukamiAddins3.ModFile
'InputTextShiftJIS  ・・・元場所：FukamiAddins3.ModFile
'GetRowCountTextFile・・・元場所：FukamiAddins3.ModFile
'InputTextUTF8      ・・・元場所：FukamiAddins3.ModFile
'fncGetCharset      ・・・元場所：FukamiAddins3.ModFile

'------------------------------

'------------------------------


Function InputText(FilePath$, Optional KugiriMoji$ = "")
'テキストファイルを読み込んで配列で返す
'文字コードは自動的に判定して読込形式を変更する
'20210721

'FilePath  ・・・テキストファイルのフルパス
'KugiriMoji・・・テキストファイルを読み込んで区切り文字で区切って配列で出力する場合の区切り文字

    'テキストファイルの存在確認
    If Dir(FilePath, vbDirectory) = "" Then
        MsgBox ("「" & FilePath & "」" & vbLf & _
               "の存在が確認できません。" & vbLf & _
               "処理を終了します。")
        End
    End If
    
    'テキストファイルの文字コードを取得
    Dim strCode
    strCode = fncGetCharset(FilePath)
    If strCode = "UTF-8 BOM" Or strCode = "UTF-8" Then
        strCode = "UTF-8"
    ElseIf strCode = "UTF-16 LE BOM" Or strCode = "UTF-16 BE BOM" Then
        strCode = "UTF-16LE"
    Else
        strCode = Empty
    End If
    
    'テキストファイル読込
    Dim Output
    Dim RowCount&
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim FileNo%, Buffer$
    
    If IsEmpty(strCode) = False Then 'UTF8版の場合※※※※※※※※※※※※※※※※※
   
        Output = InputTextUTF8(FilePath, KugiriMoji)
    
    Else 'Shift-JIS版の場合※※※※※※※※※※※※※※※※※
        
        Output = InputTextShiftJIS(FilePath, KugiriMoji)
     
    End If

    InputText = Output
    
End Function

Private Function InputTextShiftJIS(FilePath$, Optional KugiriMoji$ = "")
'テキストファイルを読み込む ShiftJIS形式専用
'20210721

'FilePath・・・テキストファイルのフルパス
'KugiriMoji・・・テキストファイルを読み込んで区切り文字で区切って配列で出力する場合の区切り文字
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim FileNo%, Buffer$, SplitBuffer
    Dim Output1, Output2
    ' FreeFile値の取得(以降この値で入出力する)
    FileNo = FreeFile
    
    N = GetRowCountTextFile(FilePath)
    ReDim Output1(1 To N)
    ' 指定ファイルをOPEN(入力モード)
    Open FilePath For Input As #FileNo
            
    ' ファイルのEOF(End of File)まで繰り返す
    I = 0
    M = 0
    Do Until EOF(FileNo)
        Line Input #FileNo, Buffer
        I = I + 1
        Output1(I) = Buffer '1次読込み
        
        If KugiriMoji <> "" Then '文字で区切る場合は区切り個数を計算
            '区切り文字による区切り個数の最大値を常に更新していく
            M = WorksheetFunction.Max(M, UBound(Split(Buffer, KugiriMoji)) + 1)
        End If
    Loop
    
    Close #FileNo
    
    '区切り文字の処理
    If KugiriMoji = "" Then
        '区切り文字なし
        Output2 = Output1
    Else
        ReDim Output2(1 To N, 1 To M)
        For I = 1 To N
            Buffer = Output1(I)
            SplitBuffer = Split(Buffer, KugiriMoji)
            For J = 0 To UBound(SplitBuffer)
                Output2(I, J + 1) = SplitBuffer(J)
            Next J
        Next I
    End If
    
    InputTextShiftJIS = Output2

End Function

Private Function GetRowCountTextFile(FilePath$)
'テキストファイル、CSVファイルの行数を取得する
'20210720

    'ファイルの存在確認
    If Dir(FilePath, vbDirectory) = "" Then
        MsgBox ("「" & FilePath & "」がありません" & vbLf & _
                "終了します")
        End
    End If
    
    Dim Output&
    With CreateObject("Scripting.FileSystemObject")
        Output = .OpenTextFile(FilePath, 8).Line
    End With
    
    GetRowCountTextFile = Output
    
End Function

Private Function InputTextUTF8(FilePath$, Optional KugiriMoji$ = "")
'テキストファイルを読み込む UTF8形式専用
'20210721

'FilePath・・・テキストファイルのフルパス
'KugiriMoji・・・テキストファイルを読み込んで区切り文字で区切って配列で出力する場合の区切り文字

    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim Buffer$, SplitBuffer
    Dim Output1, Output2
    
    N = GetRowCountTextFile(FilePath)
    ReDim Output1(1 To N)
    
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Type = 2 ' ファイルのタイプ(1:バイナリ 2:テキスト)
        .Open
        .LineSeparator = 10 '改行コード
        .LoadFromFile (FilePath)
        
        For I = 1 To N
            Buffer = .ReadText(-2)
            Output1(I) = Buffer
            If KugiriMoji <> "" Then '文字で区切る場合は区切り個数を計算
                '区切り文字による区切り個数の最大値を常に更新していく
                M = WorksheetFunction.Max(M, UBound(Split(Buffer, KugiriMoji)) + 1)
            End If
        Next I
        .Close
    End With
    
    '区切り文字の処理
    If KugiriMoji = "" Then
        '区切り文字なし
        Output2 = Output1
    Else
        ReDim Output2(1 To N, 1 To M)
        For I = 1 To N
            Buffer = Output1(I)
            SplitBuffer = Split(Buffer, KugiriMoji)
            For J = 0 To UBound(SplitBuffer)
                Output2(I, J + 1) = SplitBuffer(J)
            Next J
        Next I
    End If
    
    InputTextUTF8 = Output2
    
End Function

Private Function fncGetCharset(FileName As String) As String
'20200909追加
'テキストファイルの文字コードを返す
'参考https://popozure.info/20190515/14201

    Dim I                   As Long
    
    Dim hdlFile             As Long
    Dim lngFileLen          As Long
    
    Dim bytFile()           As Byte
    Dim b1                  As Byte
    Dim b2                  As Byte
    Dim b3                  As Byte
    Dim b4                  As Byte
    
    Dim lngSJIS             As Long
    Dim lngUTF8             As Long
    Dim lngEUC              As Long
    
    On Error Resume Next
    
    'ファイル読み込み
    lngFileLen = FileLen(FileName)
    ReDim bytFile(lngFileLen)
    If (Err.Number <> 0) Then
        Exit Function
    End If
    
    hdlFile = FreeFile()
    Open FileName For Binary As #hdlFile
    Get #hdlFile, , bytFile
    Close #hdlFile
    If (Err.Number <> 0) Then
        Exit Function
    End If
    
    'BOMによる判断
    If (bytFile(0) = &HEF And bytFile(1) = &HBB And bytFile(2) = &HBF) Then
        fncGetCharset = "UTF-8 BOM"
        Exit Function
    ElseIf (bytFile(0) = &HFF And bytFile(1) = &HFE) Then
        fncGetCharset = "UTF-16 LE BOM"
        Exit Function
    ElseIf (bytFile(0) = &HFE And bytFile(1) = &HFF) Then
        fncGetCharset = "UTF-16 BE BOM"
        Exit Function
    End If
    
    'BINARY
    For I = 0 To lngFileLen - 1
        b1 = bytFile(I)
        If (b1 >= &H0 And b1 <= &H8) Or (b1 >= &HA And b1 <= &H9) Or (b1 >= &HB And b1 <= &HC) Or (b1 >= &HE And b1 <= &H19) Or (b1 >= &H1C And b1 <= &H1F) Or (b1 = &H7F) Then
            fncGetCharset = "BINARY"
            Exit Function
        End If
    Next I
           
    'SJIS
    For I = 0 To lngFileLen - 1
        b1 = bytFile(I)
        If (b1 = &H9) Or (b1 = &HA) Or (b1 = &HD) Or (b1 >= &H20 And b1 <= &H7E) Or (b1 >= &HB0 And b1 <= &HDF) Then
            lngSJIS = lngSJIS + 1
        Else
            If (I < lngFileLen - 2) Then
                b2 = bytFile(I + 1)
                If ((b1 >= &H81 And b1 <= &H9F) Or (b1 >= &HE0 And b1 <= &HFC)) And _
                   ((b2 >= &H40 And b2 <= &H7E) Or (b2 >= &H80 And b2 <= &HFC)) Then
                   lngSJIS = lngSJIS + 2
                   I = I + 1
                End If
            End If
        End If
    Next I
           
    'UTF-8
    For I = 0 To lngFileLen - 1
        b1 = bytFile(I)
        If (b1 = &H9) Or (b1 = &HA) Or (b1 = &HD) Or (b1 >= &H20 And b1 <= &H7E) Then
            lngUTF8 = lngUTF8 + 1
        Else
            If (I < lngFileLen - 2) Then
                b2 = bytFile(I + 1)
                If (b1 >= &HC2 And b1 <= &HDF) And (b2 >= &H80 And b2 <= &HBF) Then
                   lngUTF8 = lngUTF8 + 2
                   I = I + 1
                Else
                    If (I < lngFileLen - 3) Then
                        b3 = bytFile(I + 2)
                        If (b1 >= &HE0 And b1 <= &HEF) And (b2 >= &H80 And b2 <= &HBF) And (b3 >= &H80 And b3 <= &HBF) Then
                            lngUTF8 = lngUTF8 + 3
                            I = I + 2
                        Else
                            If (I < lngFileLen - 4) Then
                                b4 = bytFile(I + 3)
                                If (b1 >= &HF0 And b1 <= &HF7) And (b2 >= &H80 And b2 <= &HBF) And (b3 >= &H80 And b3 <= &HBF) And (b4 >= &H80 And b3 <= &HBF) Then
                                    lngUTF8 = lngUTF8 + 4
                                    I = I + 3
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next I

    'EUC-JP
    For I = 0 To lngFileLen - 1
        b1 = bytFile(I)
        If (b1 = &H7) Or (b1 = 10) Or (b1 = 13) Or (b1 >= &H20 And b1 <= &H7E) Then
            lngEUC = lngEUC + 1
        Else
            If (I < lngFileLen - 2) Then
                b2 = bytFile(I + 1)
                If ((b1 >= &HA1 And b1 <= &HFE) And _
                   (b2 >= &HA1 And b2 <= &HFE)) Or _
                   ((b1 = &H8E) And (b2 >= &HA1 And b2 <= &HDF)) Then
                   lngEUC = lngEUC + 2
                   I = I + 1
                End If
            End If
        End If
    Next I
           
    '文字コード出現順位による判断
    If (lngSJIS <= lngUTF8) And (lngEUC <= lngUTF8) Then
        fncGetCharset = "UTF-8"
        Exit Function
    End If
    If (lngUTF8 <= lngSJIS) And (lngEUC <= lngSJIS) Then
        fncGetCharset = "Shift_JIS"
        Exit Function
    End If
    If (lngUTF8 <= lngEUC) And (lngSJIS <= lngEUC) Then
        fncGetCharset = "EUC-JP"
        Exit Function
    End If
    fncGetCharset = ""
    
End Function


