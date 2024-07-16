sub main
With ThisComponent.Sheets(0)
 if .getCellRangeByName("B7").value <> 90 then
     .getCellRangeByName("B3").String = ""
           Do While .getCellRangeByName("B3").String = ""
             NumeroEtratto = 1 + int( Rnd() * (90))
             For i = 3 To 12
              For n = 0 To 8
                Numero = .getCellByPosition(i, n).Value
                If NumeroEtratto = Numero Then
                  if .getCellByPosition(i, n).CellBackColor <> RGB(0, 0, 255) Then
                   .getCellRangeByName("B3").Value = NumeroEtratto
                   .getCellByPosition(i, n).CellBackColor = RGB(0, 0, 255)
                   .getCellRangeByName("B7").Value =  .getCellRangeByName("B7").Value + 1
                  End If
                 End if 
              Next n
             Next i
            Loop
 end if
end with 
end sub


Sub Ripulisci
If  ThisComponent.Sheets(0).getCellRangeByName("B7").Value = 90 Then 
  ThisComponent.Sheets(0).getCellRangeByName("D1:M9").CellBackColor = RGB(255, 255, 255)
  ThisComponent.Sheets(0).getCellRangeByName("B7").String = ""
 Else
 MsgBox "Partita in Corso"  
End if  
End sub
