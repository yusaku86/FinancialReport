Attribute VB_Name = "Main"
'//**
'*ﾔﾏﾄｺｰﾎﾟﾚｰｼｮﾝ収支報告用表作成
'*作成者:Yusaku Suzuki(2022/05/16)
'**//
Option Explicit
Const STANDARD_COL_THIS_YEAR As Long = 11  '// 当期首月の「実績」の列番号
Const STANDARD_COL_LAST_YEAR As Long = 9   '// 前期首月の「実績」の列番号
Const INTERVAL As Long = 5                 '// 各月の間の列数

'/**
 '* 今期のデータを表に入力
'**/
Public Sub importThisYearData()

    If MsgBox("各シートの[実績]に値を入力します。" & vbLf & "よろしいですか?", vbQuestion + vbYesNo, ThisWorkbook.Name) = vbNo Then
        Exit Sub
    End If

    Call ImportData(STANDARD_COL_THIS_YEAR)

End Sub

'/**
 '* 前期のデータを表に入力
'**/
Public Sub importLastYearData()

    If MsgBox("各シートの[前年度]に値を入力します。" & vbLf & "よろしいですか?", vbQuestion + vbYesNo, ThisWorkbook.Name) = vbNo Then
        Exit Sub
    End If

    Call ImportData(STANDARD_COL_LAST_YEAR)

End Sub
'/**
 '* メインルーチン
 '* Money Oneから出力したCSVデータの各値を表に反映させる
'**/
Sub ImportData(ByVal standardColumn As Long)

    '// 貼り付けた表が正しいものか確認
    If validateFile = False Then: Exit Sub

    '// 1)部門を格納した配列を作成
    Dim arrDiv As Variant: arrDiv = CreateDivArray
    
    '// 2)シート"ワーク"の値をそれぞれのシートのセルに入力
    
    Dim lastRow As Long: lastRow = Sheets("ワーク").Cells(Rows.Count, 2).End(xlUp).Row
    Dim lastColumn As Long: lastColumn = Sheets("ワーク").Cells(2, Columns.Count).End(xlToLeft).Column
    Dim counter As Long
    Dim i As Long
    Dim j As Long: j = 2
    Dim k As Long
    
    Dim dicCode As Dictionary
    
    For i = 0 To UBound(arrDiv)
        
        '// 2-1)コードを格納した配列を作成[コード番号⇒行番号]
        Set dicCode = CreateDicCode(arrDiv(i))
        
        '// 2-2)シート"ワーク"の値を対応するセルに入力
       '/**
        '* ① 部門名(シート"ワーク"のcells(j,1)の値)が変わるまでループ処理を行う
        '* ② 科目ｺｰﾄﾞ(シート"ワーク"のcells(j,2)の値がdicCodeのキーに存在したら③の処理を行う
        '* ③ 期首月から指定期間の最終月までの値を対応するセルに入力
       '**//
        
        Do While Sheets("ワーク").Cells(j, 1).Value = arrDiv(i) '// ①
            
            If dicCode.Exists(Sheets("ワーク").Cells(j, 2).Value) = False Then '// ②
                GoTo Continue
            End If
            
            counter = 0
            
            For k = 6 To lastColumn '// ③
                Sheets(arrDiv(i)).Cells(dicCode(Sheets("ワーク").Cells(j, 2).Value), standardColumn + INTERVAL * counter).Value = Sheets("ワーク").Cells(j, k).Value
                counter = counter + 1
            Next
Continue:
            j = j + 1
        Loop

    Next
    
    Set dicCode = Nothing
    
    MsgBox "入力が完了しました。", Title:=ThisWorkbook.Name
    
End Sub

'/**
 '* 貼り付けた表が適切か確認
 Private Function validateFile() As Boolean

    With Sheets("ワーク").Cells(1, 1)
        If .Value = "部門" _
        And .Offset(, 1).Value = "コード" _
        And .Offset(, 2).Value = "勘定科目" _
        And .Offset(, 3).Value = "期間累計" Then
        
            validateFile = True
            Exit Function
        End If
    End With
    
    MsgBox "シート「ワーク」に貼り付けた表が適切ではありません。", vbExclamation, ThisWorkbook.Name
    
    validateFile = False
        
 End Function
 
 

'//**
'*  部門名を格納した配列を作成
'**//
Private Function CreateDivArray() As Variant

    '// 部門名を格納する配列
    Dim arrDiv() As String
    
    '//配列arrDivに既に値が登録されているか調べる際に使用する
    Dim arrTarget As Variant
    
    '//配列のサイズ
    Dim lastRow As Long: lastRow = Sheets("ワーク").Cells(Rows.Count, 2).End(xlUp).Row
    Dim i As Long
    
    '// 部門名の配列の1つめの項目はシート「ワーク」のA2の値
    ReDim arrDiv(0)
    arrDiv(0) = Sheets("ワーク").Cells(2, 1).Value
            
    '// 部門が配列になければ登録
    For i = 3 To lastRow
        
        arrTarget = Filter(arrDiv, Sheets("ワーク").Cells(i, 1).Value)
        
        If UBound(arrTarget) = -1 Then
            ReDim Preserve arrDiv(UBound(arrDiv) + 1)
            arrDiv(UBound(arrDiv)) = Sheets("ワーク").Cells(i, 1).Value
        End If
    Next

    CreateDivArray = arrDiv
    
End Function

'//**
'*  各シートの科目コードを格納した連想配列を作成
'*
'* @param sheetName as String：配列を作成する際に参照するシート名
'**//
Private Function CreateDicCode(ByVal sheetName As String) As Dictionary
      
    Dim lastRow As Long: lastRow = Sheets(sheetName).Cells(Rows.Count, 2).End(xlUp).Row
    
    '//科目コードを格納する配列[キー：コード番号、値：列番号]
    Dim dicCode As Dictionary: Set dicCode = New Dictionary
    Dim i As Long
    
    For i = 4 To lastRow
        
        '// セルの値が空白、もしくは数字でなければ次のループへ
        If IsNumeric(Sheets(sheetName).Cells(i, 2).Value) = False Or Sheets(sheetName).Cells(i, 2).Value = "" Then
            GoTo Continue
        End If
        
        '// セルの値が配列のキーに存在しなければ配列に追加
        If dicCode.Exists(Sheets(sheetName).Cells(i, 2).Value) = False Then
            dicCode.Add Sheets(sheetName).Cells(i, 2).Value, i
        End If
    
Continue:
    Next

    Set CreateDicCode = dicCode
    Set dicCode = Nothing

End Function

