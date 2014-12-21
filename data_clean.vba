'preparation


Sub preparation()

' replace space with ?

'For Each c In Sheets("Origin").Range("J:R")
'        If c.Value = "" Then c.Value = "?"
'    Next
    
    
    
'concatenate_answer

Range("T10").Select
    Sheets("Origin").Select
    Range("T2").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(RC[-10],RC[-9],RC[-8],RC[-7],RC[-6],RC[-5],RC[-4],RC[-3],RC[-2])"
    Range("T2").Select
    Selection.AutoFill Destination:=Range("T2:T" & Worksheets("Origin").UsedRange.Rows.Count)
    Range("T2:T4543").Select
    Selection.Copy
    Range("T2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("V13").Select
    Application.CutCopyMode = False




End Sub



'step1--prepare student list


Sub student_list()

Sheets("Origin").Select

Rows("1:1").Select
Selection.AutoFilter



ActiveSheet.Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets("StudentId").Range("A:A"), Unique:=True

Sheets("StudentId").Rows("1:1").Delete


Sheets("StudentId").Range("A1:A200").Copy Destination:=Sheets("Final").Range("A2")


End Sub


Sub item_list()

Dim rowsnumber As Integer

rowsnumber = Sheets("Final").UsedRange.Rows.Count

'Sheets("Final").Cells(20, 1).Value = rowsnumber

Sheets("ItemId").Range("A1").Copy Destination:=Sheets("Final").Range("B2")

Sheets("ItemId").Range("A2").Copy Destination:=Sheets("Final").Range("F2")

Sheets("ItemId").Range("A3").Copy Destination:=Sheets("Final").Range("J2")

Sheets("ItemId").Range("A4").Copy Destination:=Sheets("Final").Range("N2")

Sheets("Final").Select

Sheets("Final").Range("B2").AutoFill Destination:=Range(Cells(2, 2), Cells(rowsnumber, 2)), Type:=xlFillCopy
Sheets("Final").Range("F2").AutoFill Destination:=Range(Cells(2, 6), Cells(rowsnumber, 6)), Type:=xlFillCopy
Sheets("Final").Range("J2").AutoFill Destination:=Range(Cells(2, 10), Cells(rowsnumber, 10)), Type:=xlFillCopy
Sheets("Final").Range("N2").AutoFill Destination:=Range(Cells(2, 14), Cells(rowsnumber, 14)), Type:=xlFillCopy

End Sub




Sub Run_This()

Sheets("StudentId").Cells.Clear
Sheets("Final").Rows(2 & ":" & Sheets("Final").Rows.Count).Cells.Clear

'Sheets("Origin").Rows("1:1").Select
'Selection.AutoFilter
'Selection.AutoFilter



If Worksheets("Origin").AutoFilterMode = True Then
   Sheets("Origin").Rows("1:1").Select
   Selection.AutoFilter
   Selection.AutoFilter

Else
   Sheets("Origin").Select
   Selection.AutoFilter
   
End If



Call student_list
Call item_list


Sheets("StudentId").Select

For Each Cell In Sheets("StudentId").Range("a1:a50")

   

   If Cell.Value <> "" Then
   
   Sheets("Origin").Select

   Rows("1:1").Select
   Selection.AutoFilter
   Selection.AutoFilter

   ActiveSheet.Range("$A$1:$R$5000").AutoFilter Field:=1, Criteria1:=Cell.Value
   
   Sheets("ItemId").Select
   
       For Each cell_2 In Sheets("ItemId").Range("a1:a50")
        
          If cell_2.Value <> "" Then
       
          Sheets("Origin").Select
     
          ActiveSheet.Range("$A$1:$R$5000").AutoFilter Field:=5, Criteria1:=cell_2.Value
       
          Sheets("Origin").UsedRange.Select
          
          Selection.Copy
       
       
          Sheets("Filter").Select
          Range("A1").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
          Rows("1:1").Delete
          
          Call select_data
          
          Call prepare_final
          

            
          Sheets("Filter").Cells.Clear
        
          Sheets("Data_select").Range("A2:E2").Clear
          
          Else
          Exit For
          
          
            
          End If
          Next cell_2
       
       
           
      







   Else
   Exit Sub
   

End If

Next Cell










End Sub



Sub select_data()

Dim firstresponse As Integer
Dim lastresponse As Integer
Dim responsecount As Integer
Dim firstanswer As String
Dim lastanswer As String
Dim q As Integer
Dim onlyone As String
Dim p As Integer
Dim x As Integer
Dim y As Integer
Dim z As Integer
Dim r As Integer











Sheets("Data_select").Cells(2, 1).Value = Sheets("Filter").Cells(1, 1).Value
Sheets("Data_select").Cells(2, 2).Value = Sheets("Filter").Cells(1, 5).Value


responsecount = Application.WorksheetFunction.CountIf(Sheets("Filter").Range("D:D"), "ResponseChanged")


Select Case responsecount
    
     Case Is < 1
          
         Sheets("Data_select").Cells(2, 3).Value = -9
         Sheets("Data_select").Cells(2, 4).Value = -9
         Sheets("Data_select").Cells(2, 5).Value = 0
        
     Case Is = 1
     
    
     
     q = Application.WorksheetFunction.Match("ResponseChanged", Sheets("Filter").Range("D:D"), 0)
     
     

     
     'onlyone = Sheets("Filter").Cells(q, 10).Value & Sheets("Filter").Cells(q, 11).Value & Sheets("Filter").Cells(q, 12).Value & Sheets("Filter").Cells(q, 13).Value & Sheets("Filter").Cells(q, 14).Value & Sheets("Filter").Cells(q, 15).Value & Sheets("Filter").Cells(q, 16).Value & Sheets("Filter").Cells(q, 17).Value & Sheets("Filter").Cells(q, 18).Value
     onlyone = Sheets("Filter").Cells(q, 20).Value
     
     p = Application.WorksheetFunction.Match(Sheets("Filter").Cells(1, 5), Sheets("ItemId").Range("A:A"), 0)
     
' test only response with key

     If onlyone = Sheets("ItemId").Cells(p, 2).Value Then
                Sheets("Data_select").Cells(2, 3).Value = 1
                Sheets("data_select").Cells(2, 4).Value = ""
    ElseIf onlyone <> Sheets("ItemId").Cells(p, 2).Value Then
                Sheets("Data_select").Cells(2, 3).Value = 0
                Sheets("data_select").Cells(2, 4).Value = ""
    End If
    
    
' get score for change

    Sheets("Data_select").Cells(2, 5).Value = 0


'          If Sheets("Data_select").Cells(2, 3).Value = Sheets("Data_select").Cells(2, 4).Value Then
'                   Sheets("Data_select").Cells(2, 5).Value = 0
'           ElseIf Sheets("Data_select").Cells(2, 3).Value <> Sheets("Data_select").Cells(2, 4).Value Then
'                   Sheets("Data_select").Cells(2, 5).Value = 1
'           End If
'
     
     
     
         'Sheets("Data_select").Cells(2, 3).Value = Sheets("Filter").Cells(2, 9).Value
         'Sheets("Data_select").Cells(2, 4).Value = Sheets("Filter").Cells(secondlast, 9).Value
     
     Case Is > 1
     
     
         x = Sheets("Filter").Columns("D:D").Find("ResponseChanged", SearchDirection:=xlNext, LookIn:=xlValues, LookAt:=xlWhole).Row
         y = Sheets("Filter").Columns("D:D").Find("ResponseChanged", SearchDirection:=xlPrevious, LookIn:=xlValues, LookAt:=xlWhole).Row
         
    
         'Sheets("Data_select").Cells(5, 1).Value = x
         'Sheets("Data_select").Cells(6, 1).Value = y
         

         
         'firstanswer = Sheets("Filter").Cells(x, 10).Value & Sheets("Filter").Cells(x, 11).Value & Sheets("Filter").Cells(x, 12).Value & Sheets("Filter").Cells(x, 13).Value & Sheets("Filter").Cells(x, 14).Value & Sheets("Filter").Cells(x, 15).Value & Sheets("Filter").Cells(x, 16).Value & Sheets("Filter").Cells(x, 17).Value & Sheets("Filter").Cells(x, 18).Value
         'lastanswer = Sheets("Filter").Cells(y, 10).Value & Sheets("Filter").Cells(y, 11).Value & Sheets("Filter").Cells(y, 12).Value & Sheets("Filter").Cells(y, 13).Value & Sheets("Filter").Cells(y, 14).Value & Sheets("Filter").Cells(y, 15).Value & Sheets("Filter").Cells(y, 16).Value & Sheets("Filter").Cells(y, 17).Value & Sheets("Filter").Cells(y, 18).Value

          firstanswer = Sheets("Filter").Cells(x, 20).Value
          lastanswer = Sheets("Filter").Cells(y, 20).Value
          



         'Sheets("Data_select").Cells(7, 1).Value = "firstanswer"
         'Sheets("Data_select").Cells(8, 1).Value = lastanswer
         
          z = Application.WorksheetFunction.Match(Sheets("Filter").Cells(1, 5), Sheets("ItemId").Range("A:A"), 0)
          


' test first response with key

      If firstanswer = Sheets("ItemId").Cells(z, 2).Value Then
                Sheets("Data_select").Cells(2, 3).Value = 1

      ElseIf firstanswer <> Sheets("ItemId").Cells(z, 2).Value Then
                Sheets("Data_select").Cells(2, 3).Value = 0

      End If
         


'' test last response with key

      If lastanswer = Sheets("ItemId").Cells(z, 2).Value Then

                Sheets("data_select").Cells(2, 4).Value = 1
      ElseIf lastanswer <> Sheets("ItemId").Cells(z, 2).Value Then

                Sheets("data_select").Cells(2, 4).Value = 0
      End If
         
         
'' get score for change




        
        
        Sheets("Filter").UsedRange.AutoFilter Field:=4, Criteria1:="ResponseChanged"

        
        For Each u In Range(Cells(2, 20), Cells(Sheets("Filter").AutoFilter.Range.SpecialCells(xlCellTypeVisible).Rows.Count, 20))
        
        
        
        'Sheets("Data_select").Cells(20, 1).Value = Sheets("Filter").AutoFilter.Range.SpecialCells(xlCellTypeVisible).Rows.Count
        
        
        
            If u.Value <> Cells(u.Rows.Count + 1, 20).Value Then
                  Sheets("Data_select").Cells(2, 5).Value = 1
                    Exit For
                  
                  
                  
                  
                             
                  
            Else
                  Sheets("Data_select").Cells(2, 5).Value = 0
                  
        
        
         
            End If
        
        Next u
        
        Sheets("Filter").AutoFilterMode = False
        
        



           
           
    
End Select
    
    



End Sub

Sub prepare_final()

Dim l As Integer
Dim r As Integer




l = Application.WorksheetFunction.Match(Sheets("Data_select").Cells(2, 1).Value, Sheets("Final").Range("A1:A100"), 0)

r = Application.WorksheetFunction.Match(Sheets("Data_select").Cells(2, 2).Value, Sheets("Final").Rows("2:2"), 0)



Sheets("Data_select").Cells(2, 3).Copy Destination:=Sheets("Final").Cells(l, r + 1)

Sheets("Data_select").Cells(2, 4).Copy Destination:=Sheets("Final").Cells(l, r + 2)

Sheets("Data_select").Cells(2, 5).Copy Destination:=Sheets("Final").Cells(l, r + 3)


'rowsused = Sheets("Final").Columns("A").Find("", Cells(Rows.Count, "A")).Count

'Sheets("Final").Cells(5, 1) = rowsused


'Sheets("Final").Cells(2, 3).Value = Sheets("Filter").Cells(2, 9).Value







End Sub


