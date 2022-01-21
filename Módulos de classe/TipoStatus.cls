VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tipoStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ID As Long
Private nome As String
Private mlo As ListObject

Public Property Get getID() As Long
    getID = ID
End Property

Public Property Get getNome() As String
    getNome = nome
End Property

Public Property Let letNome(sNome As String)
    nome = sNome
End Property

Public Sub save()

    Dim lr As ListRow
    
    If Me.getID = 0 Then 'Esta é um novo status
        ID = getAutoNumerateID(mlo)
        TiposStatus.Add Me, CStr(Me.getID)
        Set lr = mlo.ListRows.Add
    Else
        Set lr = mlo.ListRows(getLinha(mlo, "ID", Me.getID))
        
        With TiposStatus(CStr(Me.getID))
            .nome = Me.getNome
        End With
        
    End If
    
    lr.Range.value = toListRow
    
End Sub

Public Sub load(getID As Long)
    Dim lRow As Long
    
    lRow = getLinha(mlo, "ID", getID)
    
    'propriedades
    ID = mlo.DataBodyRange(lRow, getColuna(mlo, "ID"))
    nome = mlo.DataBodyRange(lRow, getColuna(mlo, "nome"))
    
End Sub

Public Sub delete()
    If Me.getID <> 0 Then
        TiposStatus.Remove CStr(Me.getID)
        mlo.ListRows(getLinha(mlo, "ID", Me.getID)).delete
    End If
End Sub

Private Function toListRow() As Variant()
    
    Dim vaReturn(0 To 0, 0 To 1) As Variant
    
    vaReturn(0, getColuna(mlo, "ID") - 1) = Me.getID
    vaReturn(0, getColuna(mlo, "nome") - 1) = Me.getNome
    
    toListRow = vaReturn

End Function

Private Sub class_initialize()
    Set mlo = tbStatus
End Sub
