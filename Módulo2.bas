Attribute VB_Name = "Módulo2"
Sub Limpar_base()
Attribute Limpar_base.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Limpar_base Macro
'

    Sheets("Janeiro").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("Fevereiro").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("Março").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("Abril").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("Maio").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    Sheets("Junho").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    Sheets("Julho").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    Sheets("Agosto").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    Sheets("Setembro").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    Sheets("Outubro").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    Sheets("Novembro").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    Sheets("Dezembro").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    Sheets("Base").Select
End Sub
