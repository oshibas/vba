Sub 表示形式()
    Range(Range("B4"), Range("B4").End(xlDown)).NumberFormat = "d日(aaa)"
End Sub

Sub 表示形式を削除()
    Range(Range("B4"), Range("B4").End(xlDown)).NumberFormat = "m月d日"
End Sub

Sub フォントの色の設定()
    Range("B3").CurrentRegion.Font.ColorIndex = Range("F11").Value
End Sub

Sub データクリア()
    Range(Range("B4"), Range("B4").End(xlDown)).Offset(, 2).ClearContents
End Sub

Sub 書式コピー()
    Worksheets("6月元").Range("B3").CurrentRegion.Copy
    Range("B3").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = falsh
End Sub

Sub セル範囲のサイズ変更()
    Dim Gyou As Integer
    Dim Retsu As Integer
    Gyou = Range("F6").Value
    Retu = Range("H6").Value
    Range("B3").Resize(Gyou, Retu).Select
End Sub

Sub 行数取得()
    Dim Nissu As Integer
    Nissu = Range("B3").CurrentRegion.Rows.Count - 1
    MsgBox "カレンダーの日数は " & Nissu & "日です。 "
End Sub

Sub 行列数取得()
    Dim Gyou As Integer
    Dim Retsu As Integer
    
    Gyou = Selection.Rows.Count
    Retsu = Selection.Columns.Count
    
    MsgBox "選択範囲の" & vbCrLf & "行数は " & Gyou & vbCrLf & "列数は " & Retsu & " です。"
End Sub

Sub 列幅自動調整()
    Columns(4).AutoFit
End Sub

Sub 列表示切替()
    Columns(3).Hidden = Not Columns(3).Hidden
End Sub

Sub 並べ替え昇順()
    Range("B3").Sort key1:=Range("C3"), Order1:=xlAscending, Header:=xlYes
End Sub

Sub 並べ替え昇順2()
    Range("B3").Sort key1:=Range("B3"), Header:=xlYes
End Sub

Sub データ抽出()
    Range("B3").CurrentRegion.AdvancedFilter _
        Action:=xlFilterCopy, criteriarange:=Range("J10:J11"), _
        copytorange:=Worksheets("抽出範囲").Range("B3:D3")
        
        Worksheets("抽出範囲").Select
End Sub
Sub データ検索()
    Dim Myrange As Range
    Dim Hakken As Range
    Dim Benti As String
    Set Myrange = Range("B3").CurrentRegion
    Myrange.Offset(1).Interior.ColorIndex = xlNone
    Set Hakken = Myrange.Find(What:=Range("J16").Value, LookIn:=xlValues, lookat:=xlPart, matchbyte:=False)
    
    If Not Hakken Is Nothing Then
        Benti = Hakken.Address
        Do
            Hakken.Interior.Color = vbMagenta
            Set Hakken = Myrange.FindNext(Hakken)
        Loop Until Hakken.Address = Benti
    End If
    Set Myrange = Nothing
    Set Hakken = Nothing
End Sub
