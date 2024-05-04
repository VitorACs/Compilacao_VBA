Attribute VB_Name = "Módulo1"

Sub compilar_plataforma()

'Declanrando as variaveis
Dim meses As Range
Dim lin, ult_lin, i As Integer
Dim MDP1(0 To 1000) As Double
Dim MDP2(0 To 1000) As Double
Dim MDP3(0 To 1000) As Double
Dim ODP1(0 To 1000) As Double
Dim ODP2(0 To 1000) As Double
Dim ODP3(0 To 1000) As Double
Dim ODP4(0 To 1000) As Double


'setando as ranges
Set meses = Sheets("Base").Range("F1:F12")

aba = 2
For Each mes In meses
    im1 = 0
    im2 = 0
    im3 = 0
    io1 = 0
    io2 = 0
    io3 = 0
    io4 = 0
    
    ult_lin = Range("A2").End(xlDown).Row
    For lin = 2 To ult_lin
        
        'Para janeiro
        
        If Cells(lin, 1).Value = mes And Cells(lin, 3).Value = "MDP1" Then
        
            MDP1(0 + im1) = Cells(lin, 4).Value
            im1 = im1 + 1
            
        ElseIf Cells(lin, 1).Value = mes And Cells(lin, 3).Value = "MDP2" Then

            MDP2(0 + im2) = Cells(lin, 4).Value
            im2 = im2 + 1
        
        ElseIf Cells(lin, 1).Value = mes And Cells(lin, 3).Value = "MDP3" Then

            MDP3(0 + im3) = Cells(lin, 4).Value
            im3 = im3 + 1
            
        ElseIf Cells(lin, 1).Value = mes And Cells(lin, 3).Value = "ODP1" Then

            ODP1(0 + io1) = Cells(lin, 4).Value
            io1 = io1 + 1
        
        ElseIf Cells(lin, 1).Value = mes And Cells(lin, 3).Value = "ODP2" Then

            ODP2(0 + io2) = Cells(lin, 4).Value
            io2 = io2 + 1
        
        ElseIf Cells(lin, 1).Value = mes And Cells(lin, 3).Value = "ODP3" Then

            ODP3(0 + io3) = Cells(lin, 4).Value
            io3 = io3 + 1
            
        ElseIf Cells(lin, 1).Value = mes And Cells(lin, 3).Value = "ODP4" Then

            ODP4(0 + io4) = Cells(lin, 4).Value
            io4 = io4 + 1

        End If
        
    Next lin
    
    'Colar na aba janeiro
    
    Sheets(aba).Range("B2:B1000").Value = Application.Transpose(MDP1)
    Sheets(aba).Range("C2:C1000").Value = Application.Transpose(MDP2)
    Sheets(aba).Range("D2:D1000").Value = Application.Transpose(MDP3)
    Sheets(aba).Range("E2:E1000").Value = Application.Transpose(ODP1)
    Sheets(aba).Range("F2:F1000").Value = Application.Transpose(ODP2)
    Sheets(aba).Range("G2:G1000").Value = Application.Transpose(ODP3)
    Sheets(aba).Range("H2:H1000").Value = Application.Transpose(ODP4)
       
    aba = aba + 1
    
    Erase MDP1, MDP2, MDP3, ODP1, ODP2, ODP3, ODP4

Next mes

MsgBox ("Macro Executada com Sucesso")
End Sub
