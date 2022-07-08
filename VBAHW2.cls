VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub Atest()
    ' Create ticker variable
    'Create volume variable
    'Create yealry change variable
    'Create percent change variable
    
    Dim tick As String
    Dim vol As Double
    Dim yrc As Double
    Dim perc As Double
    Dim ws As Worksheet
    Dim i As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim RowCount As Long
    
    
    
    
    
    
    

For Each ws In Worksheets
    
    
    ' Set header rows
    ws.Range("I1").Value = "Ticker"
    ws.Range("L1").Value = "Total Volume"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    
    
    'Set initial values to 0
    ' vol = 0
    'Get last row
    RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'set up for loop
    For i = 2 To RowCount
    ' This is saying if this first condition doesn't match the second
    'condition(initial formula) then we keep finding the values
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'Set up yearly change & percent change formula
            open_price = ws.Cells(i, 3).Value
            close_price = ws.Cells(i, 6).Value
            yrc = close_price - open_price
           ' Debug.Print yrc
            
            If yrc < 0 Then
                ws.Range("J2").Interior.ColorIndex = 4
            
            ElseIf yrc > 0 Then
                ws.Range("J2").Interior.ColorIndex = 3
                
          
            
            Else
                ws.Range("J2").Interior.ColorIndex = 0
            End If
            
            
                
             If open_price <> 0 Then
             
                perc = yrc / open_price
            
            Else
                perc = 0
            
        End If
        
        'Prints the ticker symbol, j is an integer
            ws.Range("I" & 2 + J).Value = ws.Cells(i, 1).Value
        'Print total value
            ws.Range("L" & 2 + J).Value = vol
            vol = 0
        'Print yearly change
            ws.Range("J" & 2 + J).Value = yrc
            yrc = 0
            
        'Print percent change
            ws.Range("K" & 2 + J).Value = perc
            perc = 0
            
        'Moving to next row J
            J = J + 1
        Else
             vol = vol + ws.Cells(i, 7).Value
             
             
        End If
        
    Next i
                J = 0
               
                
                
                
Next ws
    
End Sub
