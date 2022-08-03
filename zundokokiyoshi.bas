Option Explicit
Option Base 1

#If Win64 Then
    Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
    (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
    ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
#Else
    Private Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" _
    (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
    ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
#End If


Sub zundokoKiyoshi()
'「ズン」「ドコ」のいずれかをランダムで出力し続けて「ズン」「ズン」「ズン」「ズン」「ドコ」の配列が出たら「キ・ヨ・シ！」って出力
'https://qiita.com/shunsugai@github/items/971a15461de29563bf90

    Dim cs As CastSpell: Set cs = New CastSpell 'おまじないを呼び出す
    
    'Long
    Dim cntZun As Long
    Dim cnt As Long
    Dim cntArr As Long
    Dim cntZundoko As Integer '配列用の乱数
    Dim lngRow As Long: lngRow = 1
    Dim lastRow As Long
    
    'String
    Dim arrZundoko(2) As String
        arrZundoko(1) = "ズン♪"
        arrZundoko(2) = "ドコ"
    Dim strZundoko As String
    'セルC1にズンドコ音楽のパス
    Dim strMusicZundoko As String: strMusicZundoko = Worksheets("ズンドコ").Range("C1").Value
    
    Worksheets("ズンドコ").Activate
    Columns(1).Clear 'A列の値を削除
    
    '---ズンドコここから開始---
    Do
        cntZundoko = Int(2 * Rnd + 1)
        strZundoko = strZundoko + arrZundoko(cntZundoko)
        
        If arrZundoko(cntZundoko) = "ズン♪" And cntZun < 4 Then 'ズンがやってきて、ズンカウンタが4未満だったら
            cntZun = cntZun + 1 'ズンが出てきたのでカウンタをインクリメントする
        ElseIf arrZundoko(cntZundoko) = "ドコ" And cntZun = 4 Then 'ドコがやってきて、ズンカウンタが4だったら
            Cells(lngRow, 1).Value = strZundoko & " キ・ヨ・シ！" 'ズンドコキヨシ整いました
            Exit Do
        Else 'どれにも該当しなかったら
            Cells(lngRow, 1).Value = strZundoko
            strZundoko = ""
            cntZun = 0
            lngRow = lngRow + 1
        End If
    
    Loop
    '---ズンドコここまで---
    
    Columns("A:A").AutoFit
    
    Worksheets("ズンドコヒストリー").Activate
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row + 1
    Cells(lastRow, 1).Value = Now
    Cells(lastRow, 2).Value = lngRow
    
    Cells(1, 6).Value = WorksheetFunction.Average(Range(Cells(1, 2), Cells(lastRow, 2)))
    
    Columns("A:F").AutoFit
    
    Set cs = Nothing
    
    Worksheets("ズンドコ").Activate
    
    Kiyoshi.Caption = lngRow & "回目で整いました！"
    Call PlayMusic(strMusicZundoko)
    Kiyoshi.Show
    
    MsgBox lngRow & "回目で整いました！"
    
    Unload Kiyoshi

End Sub

Function PlayMusic(ByRef file As String)
'参考：https://liclog.net/mcisendstring-function-vba-api/

    Call mciSendString("play " & file, "", 0, 0)

End Function
        
'ユーザフォーム側に仕込んだもの
Private Sub UserForm_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    
    With Kiyoshi
        .Width = 375.6
        .Height = 544.4
    End With
    startupposition = 2  '画面の中央
End Sub
