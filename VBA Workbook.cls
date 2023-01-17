VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stocks_Loop():

Dim Ws As Worksheet
Dim Summ_Header As Boolean


    Summ_Header = False
         'Set Header flag
 

    
        
        ' Loop through all ws in active workbook.

        For Each Ws In Worksheets

         

            ' Set holding for ticker

            Dim Ticker_Name As String

            Ticker_Name = " "

             

          

            Dim Total_Ticker_Volume As Double

            Total_Ticker_Volume = 0

 

            ' 1st variabe for tpt name

             

            Dim Open_Price As Double

            Open_Price = 0

            Dim Close_Price As Double

            Close_Price = 0

            Dim Delta_Price As Double

            Delta_Price = 0

            Dim Delta_Percent As Double

            Delta_Percent = 0

                        'New variables

 

                ' Now comes the real work lol

 


              

            '  The Code needs to know where each ticker is =>


            Dim Summy_TableRow As Long
            
                Summy_TableRow = 2

             

            

            Dim L As Long
            

            Dim i As Long

 

            ' The start of the row count to the end

             

           
           L = Ws.Cells(Rows.Count, 1).End(x1Up).Row
           

  

            

 

 

            If Summ_Header Then

                ' Set Titles for Summary Table for worksheet

                Ws.Range("I1").Value = "Ticker"

                Ws.Range("J1").Value = "Yearly Change"

                Ws.Range("K1").Value = "Percent Change"

                Ws.Range("L1").Value = "Total Stock Volume"

         

 

            End If

           ' -------------------------------------------------------------------------------------------------------------------------------------------

            ' Set initial value of Open Price for the first Ticker and let it loop

             

            Open_Price = Ws.Cells(2, 3).Value

             

         

            For i = 2 To L

             
             
             'loop from start to finish

 

   

                ' Check if we are still within the same ticker ran, if not - write results to summary table

                If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then

                 

    

                    Ticker_Name = Ws.Cells(i, 1).Value

                     

                    ' Calculate Delta_Price and Delta_Percent

                    Close_Price = Ws.Cells(i, 6).Value

                    Delta_Price = Close_Price - Open_Price

                    

                    If Open_Price <> 0 Then

                                    '    divide by 0




                        Delta_Percent = (Delta_Price / Open_Price) * 100

                                        

 

                     

                    Total_Ticker_Volume = Total_Ticker_Volume + Ws.Cells(i, 7).Value

                   

                                                      '       +to TTV

                     

                    ' Result the Ticker Name  Column I

                    Ws.Range("I" & Summy_TableRow).Value = Ticker_Name

 

                    ' Print the Ticker Name in the Summary Table, Column J

                    Ws.Range("J" & Summy_TableRow).Value = Delta_Price

                     

 

                    End If

                     

                     ' Print the Ticker Name in K

                    Ws.Range("K" & Summy_TableRow).Value = (CStr(Delta_Percent) & "%")

                    ' Print the Ticker Name L

                    Ws.Range("L" & Summy_TableRow).Value = Total_Ticker_Volume

                     

                    ' Add 1 to the summary table row count

                    Summy_TableRow = Summy_TableRow + 1
                    
                    
                     ' ---------------------------------------------------------------
                     
                    

                    ' New Ticker #'s => Reset!
                    
                    Delta_Price = 0

                    Open_Price = Ws.Cells(i + 1, 3).Value

                   

                End If
                
                
    Next Ws
                
            
    
End Sub
