Attribute VB_Name = "Module1"
Sub CreateOrderList()

  Dim wbThis As Workbook
  Set wbThis = ThisWorkbook
  Dim wbInventory As Workbook
  Dim wbTrueInventory As Workbook
  Dim wsOrder As Worksheet
  Dim wsData As Worksheet
  Dim wsTrueData As Worksheet
  Dim wsError As Worksheet
  Dim Fd As FileDialog
  Dim Folderpath As String
  Dim Code As String
  Dim first As String
  Dim Fso As Object
  Dim Folder As Object
  Dim Found As Range
  Dim Result As VbMsgBoxResult
  Dim DataLast As Long
  Dim TrueDataLast As Long
  Dim Errorlast As Long
  Dim cnt As Long
  Dim i As Long
  Dim j As Long
  Dim File As Variant
  Dim DataArr As Variant
  Dim TrueDataArr As Variant
  Dim key As Variant
  Dim dict As Object
  

'フォルダパス取得処理
  Set Fd = Application.FileDialog(msoFileDialogFolderPicker)

If Fd.Show = -1 Then
  Folderpath = Fd.SelectedItems(1)
  
Else
  Debug.Print "フォルダパスが見つかりません。"
  Exit Sub
  
End If


'フォルダパスの存在確認 ～ フォルダ取得
  Set Fso = CreateObject("scripting.filesystemobject")
  
If Fso.folderexists(Folderpath) Then
  Set Folder = Fso.getfolder(Folderpath)
  Debug.Print "フォルダ取得"
  
Else
  Debug.Print "フォルダが見つかりません。"
  Exit Sub

End If


'ファイル取得 ～ 「在庫一覧」シート取得
For Each File In Folder.Files

  If Left(File.Name, 9) = "Inventory" And Right(File.Name, 5) = ".xlsx" Then
    
    On Error Resume Next
    Set wbInventory = Workbooks.Open(File.Path)
    On Error GoTo 0
    
    If wbInventory Is Nothing Then
      GoTo jump
      
    Else
      Debug.Print "ファイル取得。ファイル名:" & File.Name

    End If
    
    
    On Error Resume Next
    Set wsData = Worksheets("在庫一覧")
    On Error GoTo 0
    
    If wsData Is Nothing Then
      GoTo jump
      
    Else
      Debug.Print "「在庫一覧」シート取得"
      Exit For
      
    End If
    
jump:

  End If
  
Next File



'ユーザーへ「発注一覧データ」作成処理、実行開始の可否確認
Result = MsgBox("処理を開始しますか？", vbYesNo + vbQuestion, "確認")

If Result = vbNo Then

  Exit Sub
  
Else

  Debug.Print "処理開始"
  
End If


'「エラー一覧」シート取得
On Error Resume Next
  wbThis.Worksheets("エラー一覧").Delete
On Error GoTo 0
    
  Set wsError = wbThis.Worksheets.Add
  wsError.Name = "エラー一覧"
  Debug.Print "「エラー一覧」シート取得"
  
  wsError.Range("a1").Value = "エラー箇所"
  wsError.Range("b1").Value = "エラー行番号"
  wsError.Range("c1").Value = "エラー理由"
  
  wbThis.Worksheets("エラー一覧").Columns("A:C").ColumnWidth = 13.5
  

'「在庫一覧」シートデータを配列へ格納
  DataLast = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
  DataArr = wsData.Range("a2:g" & DataLast).Value
  Debug.Print "在庫データを配列へ格納完了"


j = 2

For i = 1 To UBound(DataArr, 1)

  If DataArr(i, 1) = "" Then
    wsError.Cells(j, 1).Value = wsData.Range("a1")
    wsError.Cells(j, 2).Value = i + 1 & "行目"
    wsError.Cells(j, 3).Value = "セル空白"
    
    j = j + 1

  End If
  
  If DataArr(i, 2) = "" Then
    wsError.Cells(j, 1).Value = wsData.Range("b1")
    wsError.Cells(j, 2).Value = i + 1 & "行目"
    wsError.Cells(j, 3).Value = "セル空白"
    
    j = j + 1
    
  End If

  If DataArr(i, 3) = "" Then
    wsError.Cells(j, 1).Value = wsData.Range("c1")
    wsError.Cells(j, 2).Value = i + 1 & "行目"
    wsError.Cells(j, 3).Value = "セル空白"
    
    j = j + 1
    
  End If

  If DataArr(i, 4) = "" Then
    wsError.Cells(j, 1).Value = wsData.Range("d1")
    wsError.Cells(j, 2).Value = i + 1 & "行目"
    wsError.Cells(j, 3).Value = "セル空白"
    
    j = j + 1

  End If

  If DataArr(i, 5) = "" Then
    wsError.Cells(j, 1).Value = wsData.Range("e1")
    wsError.Cells(j, 2).Value = i + 1 & "行目"
    wsError.Cells(j, 3).Value = "セル空白"
    
    j = j + 1

  End If

  If DataArr(i, 6) = "" Then
    wsError.Cells(j, 1).Value = wsData.Range("f1")
    wsError.Cells(j, 2).Value = i + 1 & "行目"
    wsError.Cells(j, 3).Value = "セル空白"
    
    j = j + 1

  End If

  If Not IsNumeric(DataArr(i, 6)) Then
    wsError.Cells(j, 1).Value = wsData.Range("f1")
    wsError.Cells(j, 2).Value = i + 1 & "行目"
    wsError.Cells(j, 3).Value = "数値以外の値"
    
    j = j + 1
    
  End If
    
  If IsNumeric(DataArr(i, 6)) = True Then
    
    If DataArr(i, 6) <> Int(DataArr(i, 6)) Then
      wsError.Cells(j, 1).Value = wsData.Range("f1")
      wsError.Cells(j, 2).Value = i + 1 & "行目"
      wsError.Cells(j, 3).Value = "小数点あり"
    
      j = j + 1

    End If
      
    If DataArr(i, 6) < 0 Then
      wsError.Cells(j, 1).Value = wsData.Range("f1")
      wsError.Cells(j, 2).Value = i + 1 & "行目"
      wsError.Cells(j, 3).Value = "マイナス値"
    
      j = j + 1

    End If
  
  End If
  
Next i


  Errorlast = wsError.Cells(wsError.Rows.Count, 1).End(xlUp).Row
  
If Errorlast >= 2 Then
  MsgBox "「エラー一覧」シート出力。異常件数:" & Errorlast - 1 & "件" & vbCrLf _
  & "異常データの修正をしてください。"
  'exit sub
  '本来なら在庫データに異常が1つでもあればここで処理を終了するが、
  '今回は異常データを残す方向で進めるので処理を続行する
  
Else
  MsgBox "異常なし。"

End If


'ユーザーが在庫データの異常を修正したと仮定して、修正後の
'仮のデータを別ファイルから取得
For Each File In Folder.Files

  If Left(File.Name, 14) = "Inventory_True" And _
     Right(File.Name, 5) = ".xlsx" Then
    
    On Error Resume Next
    Set wbTrueInventory = Workbooks.Open(File.Path)
    On Error GoTo 0
    
    If wbTrueInventory Is Nothing Then
      GoTo jump2
      
    Else
      Debug.Print "ファイル取得。ファイル名:" & File.Name

    End If
    
    
    On Error Resume Next
    Set wsTrueData = Worksheets("在庫一覧")
    On Error GoTo 0
    
    If wsTrueData Is Nothing Then
      GoTo jump2
      
    Else
      Debug.Print "「在庫一覧」シート取得"
      Exit For
      
    End If
    
jump2:

  End If
  
Next File

    

'「発注一覧」シート取得
On Error Resume Next
  wbThis.Worksheets("発注一覧").Delete
On Error GoTo 0
    
  Set wsOrder = wbThis.Worksheets.Add
  wsOrder.Name = "発注一覧"
  Debug.Print "「発注一覧」シート取得"
  
  wsOrder.Range("a1").Value = "商品名"
  wsOrder.Range("b1").Value = "発注数"
  wbThis.Worksheets("発注一覧").Columns("a:b").ColumnWidth = 13.5

'修正後の在庫データを配列へ格納
  TrueDataLast = wsTrueData.Cells(wsTrueData.Rows.Count, 1).End(xlUp).Row
  TrueDataArr = wsTrueData.Range("a2:g" & TrueDataLast).Value
  Set dict = CreateObject("scripting.dictionary")
  
  
'商品別在庫数カウント
For i = 1 To UBound(TrueDataArr, 1)

  If Not dict.exists(TrueDataArr(i, 2)) Then
    dict.Add TrueDataArr(i, 2), TrueDataArr(i, 6)
    
  Else
    dict(TrueDataArr(i, 2)) = dict(TrueDataArr(i, 2)) + TrueDataArr(i, 6)
    
  End If

Next i

Debug.Print "商品別在庫数カウント完了"



j = 2
cnt = 0

For Each key In dict.keys
  
  Set Found = wsTrueData.Range("b:b").Find(what:=key, lookat:=xlWhole, LookIn:=xlValues)
  
  If Found Is Nothing Then
    Debug.Print "「在庫一覧」シート、データなし。"
    Exit Sub
    
  End If
  
  cnt = Found.Offset(0, 5).Value - dict(key) '発注数
  
  If cnt <= 0 Then
    wsOrder.Cells(j, 1).Value = key
    wsOrder.Cells(j, 2).Value = 0
  
  j = j + 1
  
  Else
    wsOrder.Cells(j, 1).Value = key
    wsOrder.Cells(j, 2).Value = cnt
  
  j = j + 1
  
  End If
  
Next key


If Not wbInventory Is Nothing Then
  wbInventory.Close savechanges = False
  
End If



End Sub
