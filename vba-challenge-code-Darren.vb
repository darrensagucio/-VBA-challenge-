Sub finaltickerhw():

For Each ws In Worksheets

'Declaring Variables
Dim tickername As String

Dim tickervolumetotal As Double
tickervolumetotal = 0

Dim tickerlocation As Integer
tickerlocation = 2

Dim opennum As Double
Dim closednum As Double
opennum = 0
closednum = 0

Dim yearchangenum As Double
yearchangenum = 0

Dim percentchange As Double
percentchange = 0

Dim greatestpercent As Double
greatestpercent = 0

Dim lowestpercent As Double
lowestpercent = 0

Dim checkgreatesttotalvolume As Variant
checkgreatesttotalvolume = 0
Dim greatesttotalvolume As Variant
greatesttotalvolume = 0

Dim greatestpercentname As String
Dim lowestpercentname As String
Dim volumename As String

'Printing Table Names
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change "
ws.Cells(1, 12).Value = "Total Stock Volume"

ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume "

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'Calulation For Column I
    tickername = ws.Cells(i, 1).Value
    
    'Calulation For Column L
    tickervolumetotal = tickervolumetotal + ws.Cells(i, 7).Value
    
    'Calulation For Column K
    closednum = ws.Cells(i, 6).Value
    yearchangenum = closednum - opennum
        If opennum <> 0 Then
        percentchange = (yearchangenum * 100) / opennum
        Else
        percentchange = 0
        End If

            'Adding Data To Desired Columns
            ws.Range("I" & tickerlocation).Value = tickername
            ws.Range("J" & tickerlocation).Value = yearchangenum
            
            If yearchangenum > 0 Then
            ws.Range("J" & tickerlocation).Interior.ColorIndex = 4
            Else
            ws.Range("J" & tickerlocation).Interior.ColorIndex = 3
            End If
        
            ws.Range("K" & tickerlocation).Value = percentchange & " % "
            ws.Range("L" & tickerlocation).Value = tickervolumetotal
     
    
    'Onto Next Stock on Worksheet
    tickerlocation = tickerlocation + 1
    
    'Restarting Variables
    tickervolumetotal = 0
    yearchangenum = 0
    opennum = 0
    closednum = 0
    percentchange = 0
    checkgreatestpercent = 0
    greatestpercent = 0
    Else
        
        'Obtaining First Open Price For Each Stock
        If opennum = 0 Then
        opennum = ws.Cells(i, 3).Value
        End If
    
        'Adding Total Volume
         tickervolumetotal = tickervolumetotal + ws.Cells(i, 7).Value
            
    End If

Next i
         
         lastrowtotalpercent = ws.Cells(Rows.Count, 11).End(xlUp).Row
         
       'Calulations For Finding Greatest Increase And Decrease Percentages
       For i = 2 To lastrowtotalpercent
         If ws.Cells(i, 11).Value > greatestpercent Then
         greatestpercent = ws.Cells(i, 11).Value
         End If
         
         If ws.Cells(i, 11).Value = greatestpercent Then
         greatestpercentname = ws.Cells(i, 9).Value
         End If
         
         If ws.Cells(i, 11).Value < lowestpercent Then
         lowestpercent = ws.Cells(i, 11).Value
         End If
         
         If ws.Cells(i, 11).Value = lowestpercent Then
         lowestpercentname = ws.Cells(i, 9).Value
         End If
       Next i
         
         'Printing Greatest Increase Percentage And Name
         greatestpercent = greatestpercent * 100
         ws.Cells(2, 16).Value = greatestpercent & "%"
         ws.Cells(2, 15).Value = greatestpercentname
         
         ''Printing Greatest Dercrease Percentage And Name
         lowestpercent = lowestpercent * 100
         ws.Cells(3, 16).Value = lowestpercent & "%"
         ws.Cells(3, 15).Value = lowestpercentname
         
        '----------------------------------------------------------
        
         lastrowtotalvolume = ws.Cells(Rows.Count, 12).End(xlUp).Row
           
        'Calulations for Greatest Total Volume
       For i = 2 To lastrowtotalvolume
         If ws.Cells(i + 1, 12).Value > ws.Cells(i, 12).Value Then
         checkgreatesttotalvolume = ws.Cells(i + 1, 12).Value
         Else
         checkgreatesttotalvolume = ws.Cells(i, 12).Value
         End If
        
         If checkgreatesttotalvolume > greatesttotalvolume Then
         greatesttotalvolume = checkgreatesttotalvolume
         End If
             
         If ws.Cells(i, 12).Value = greatesttotalvolume Then
         volumename = ws.Cells(i, 9).Value
         End If
       Next i
                
         'Printing Greatest Total Volume And Name
         ws.Cells(4, 16).Value = greatesttotalvolume
         ws.Cells(4, 15).Value = volumename
         
         
         'Restart Variables
         checkgreatesttotalvolume = 0
         greatesttotalvolume = 0
         countvolumeforname = 0
         
     ws.Columns("A:P").AutoFit

 Next ws

End Sub