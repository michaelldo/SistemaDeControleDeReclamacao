Attribute VB_Name = "mDatabase"
Option Explicit
Option Private Module

Public Function tbAsc() As ListObject: Set tbAsc = wsDatabase.ListObjects("tbAsc"): End Function

Public Function tbOuvidoria() As ListObject: Set tbOuvidoria = wsDatabase.ListObjects("tbOuvidoria"): End Function

Public Function tbStatus() As ListObject: Set tbStatus = wsDatabase.ListObjects("tbStatus"): End Function

Public Function tbIndices() As ListObject: Set tbIndices = wsDatabase.ListObjects("tbIndices"): End Function

Public Function tbTipo() As ListObject: Set tbTipo = wsDatabase.ListObjects("tbTipo"): End Function

Public Function tbInformante() As ListObject: Set tbInformante = wsDatabase.ListObjects("tbInformante"): End Function

Public Function tbUf() As ListObject: Set tbUf = wsDatabase.ListObjects("tbUf"): End Function

Public Function getAutoNumerateID(lo As ListObject) As Long
    Dim lReturn As Long
    Dim iRow As Integer
    
    iRow = getLinha(tbIndices, "tabela", lo.Name)
    
    If iRow = -1 Then Err.Raise 1, , "tabela não encontrada"
        
    With tbIndices
        lReturn = .ListColumns("ID").DataBodyRange(iRow).value
        .ListColumns("ID").DataBodyRange(iRow).value = lReturn + 1
    End With
    
    getAutoNumerateID = lReturn
    
End Function

