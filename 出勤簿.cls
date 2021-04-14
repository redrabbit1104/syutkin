VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

Dim A As Date, i As Long, e As Long, d As Long, f As Long, j As Long

A = Cells(3, 4).Value      '最初の日付のシリアル値を取得（例えば12月1日）'
e = DateSerial(Year(A), Month(A) + 1, 0)  '翌月1日の前日を取得'
d = Format(e, "d")    '末日を取得（31日など）'
x = Format(A, "d")    'D3セルの日を取得（3日など)'
f = 5 + d - Day(A)    '末日から１を足したセルの列番号を取得'

If Target.Address = Range("D3").Address Then

For i = 1 To d - x
    Cells(3, 4 + i).Value = A + i
    Cells(4, 4).Value = Format(A, "aaaa")    '最初のD4セルの曜日を指定'
    Cells(4, 4 + i).Value = Format(A + i, "aaaa")  'E4セルから末日のセルまでの曜日を指定'
        
    If i = d - x Then
    
     Application.EnableEvents = False  'eventプロシージャをストップ'
     
     
For j = Range("D3").End(xlToRight).Column To 1 Step -1 'D3セルのデータが入っている右最後のセル番号から1列番号まで-1ずつ逆戻りしながら以下の処理をする'
    With Cells(4, j)
        If _
        .Value = "土曜日" Or _
        .Value = "日曜日" Then
        .EntireColumn.Delete
         End If
    End With
Next j

    End If
 
Next

Dim Rowx As Integer, Columny As Integer
Rowx = Range("D3").End(xlDown).Row
Columny = Range("D3").End(xlToRight).Column

Range("D3", Cells(4, Columny)).BorderAround Weight:=xlMedium, LineStyle:=xlContinuous    'D3セルからColumny4セルまでの外側の線を太さMedium、Continuousに設定'
     

With Range("D3", Cells(4, Columny)).Borders(xlInsideVertical)  'D3セルからColumny4セルまでの内側の垂直線を太さxlThin、Continuousに設定'
     .LineStyle = xlContunuous
     .Weight = xlThin
End With
     
With Range("D3", Cells(4, Columny)).Borders(xlInsideHorizontal)  'D3セルからColumny4セルまでの内側の水平線を太さxlThin、Continuousに設定'
     .LineStyle = xlContunuous
     .Weight = xlThin
End With

With Range("D5", Cells(11, Columny)).Borders(xlEdgeBottom)   'D5セルからColumny11セルまでのBottom Lineの線を太さMedium、Continuousに設定'
     .LineStyle = xlContinuous
     .Weight = xlMedium
End With

With Range("D5", Cells(11, Columny)).Borders(xlInsideVertical)   'D5セルからColumny11セルまでの内側の垂直線を太さxlThin、Continuousに設定'
     .LineStyle = xlContinuous
     .Weight = xlThin
End With

With Range("D5", Cells(11, Columny)).Borders(xlInsideHorizontal)   'D5セルからColumny11セルまでの内側の水平線を太さxlThin、Continuousに設定'
     .LineStyle = xlDot
     .Weight = xlThin
End With

With Range("D12", Cells(17, Columny)).Borders(xlEdgeBottom)   'D12セルからColumny17セルまでのBottom Lineの線を太さMedium、Continuousに設定'
     .LineStyle = xlContinuous
     .Weight = xlMedium
End With

With Range("D12", Cells(17, Columny)).Borders(xlInsideVertical)   'D12セルからColumny17セルまでの内側の垂直線を太さxlThin、Continuousに設定'
     .LineStyle = xlContinuous
     .Weight = xlThin
End With

With Range("D12", Cells(17, Columny)).Borders(xlInsideHorizontal)   'D12セルからColumny17セルまでの内側の水平線を太さxlThin、Continuousに設定'
     .LineStyle = xlDot
     .Weight = xlThin
End With

With Range("D18", Cells(22, Columny)).Borders(xlEdgeBottom)   'D18セルからColumny22セルまでのBottom Lineの線を太さMedium、Continuousに設定'
     .LineStyle = xlContinuous
     .Weight = xlMedium
End With

With Range("D18", Cells(22, Columny)).Borders(xlInsideVertical)   'D18セルからColumny22セルまでの内側の垂直線を太さxlThin、Continuousに設定'
     .LineStyle = xlContinuous
     .Weight = xlThin
End With

With Range("D18", Cells(22, Columny)).Borders(xlInsideHorizontal)   'D18セルからColumny22セルまでの内側の水平線を太さxlThin、Continuousに設定'
     .LineStyle = xlDot
     .Weight = xlThin
End With

With Range("D23", Cells(27, Columny)).Borders(xlEdgeBottom)   'D23セルからColumny27セルまでのBottom Lineの線を太さMedium、Continuousに設定'
     .LineStyle = xlContinuous
     .Weight = xlMedium
End With

With Range("D23", Cells(27, Columny)).Borders(xlInsideVertical)   'D23セルからColumny27セルまでの内側の垂直線を太さxlThin、Continuousに設定'
     .LineStyle = xlContinuous
     .Weight = xlThin
End With

With Range("D23", Cells(27, Columny)).Borders(xlInsideHorizontal)   'D23セルからColumny27セルまでの内側の水平線を太さxlThin、Continuousに設定'
     .LineStyle = xlDot
     .Weight = xlThin
End With

Dim k As Long

For k = Range("D3").End(xlToRight).Column To 1 Step -1 '金曜日が入っているセルの右側の罫線の太さをxlMedium,線の種類をxlContinuousに設定'
    With Cells(4, k)
        If _
        Cells(4, k).Value = "金曜日" Then
        With Range(Cells(3, k), Cells(27, k)).Borders(xlEdgeRight)
         .LineStyle = xlContinuous
         .Weight = xlMedium
         End With
         End If
    End With
Next k

Range("d3", Cells(27, Columny)).Interior.ColorIndex = 0 'セル範囲の色を白に'

Range("d3", Cells(27, Columny)).ColumnWidth = 10   'セル範囲の幅を10に'

Range("d3", Cells(27, Columny)).UnMerge   'セル範囲の結合を解除'

Range("d3", Cells(27, Columny)).Borders(xlDiagonalUp).LineStyle = xlNone   'セル範囲の結合を解除'

End If


    

Application.EnableEvents = True   'eventプロシージャを開始'

End Sub
