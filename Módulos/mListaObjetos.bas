Attribute VB_Name = "mListaObjetos"
Option Explicit
Option Private Module
'função para identificar o ID da coluna de u listObject
Public Function getColuna(loListObject As ListObject, nomeColuna As String) As Long

    On Error GoTo ErrRaise 'caso ocorra erro
        getColuna = WorksheetFunction.Match(nomeColuna, loListObject.HeaderRowRange, 0)
    Exit Function
    
ErrRaise: 'caso ocorra erro retorna o valor -1
    getColuna = -1
End Function

'função para identificar a linha de um ListObject
Public Function getLinha(loListObject As ListObject, nomeColuna As String, value As Variant) As Long
    
    Dim colunaId As Long
    Dim rngDados As Range
    
    
    On Error GoTo ErrRaise 'se ocorrer algum erro por não encontra a linha
    
    colunaId = getColuna(loListObject, nomeColuna)
    
    Set rngDados = loListObject.ListColumns(colunaId).DataBodyRange
    
    getLinha = WorksheetFunction.Match(value, rngDados, 0)
    
    Exit Function

ErrRaise: 'retorna 1 em caso de erro
    getLinha = -1
        
End Function
