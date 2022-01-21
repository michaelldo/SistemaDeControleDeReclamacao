VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "asc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    
Private ID As Long
Private data As String
Private hora As String
Private nome As String
Private cpf As String
Private ASC As String
Private motivo As String
Private mlo As ListObject

Public Property Get getID() As Long
    getID = ID
End Property

Public Property Get getData() As String
    getData = data
End Property

Public Property Get getHora() As String
    getHora = hora
End Property

Public Property Get getNome() As String
    getNome = nome
End Property

Public Property Let letNome(sNome As String)
    nome = sNome
End Property

Public Property Get getCpf() As String
    getCpf = cpf
End Property

Public Property Let letCpf(sCpf As String)
    cpf = sCpf
End Property

Public Property Get getAsc() As String
    getAsc = ASC
End Property

Public Property Let letAsc(sAsc As String)
    ASC = sAsc
End Property

Public Property Get getMotivo() As String
    getMotivo = motivo
End Property

Public Property Let letMotivo(sMotivo As String)
    motivo = sMotivo
End Property

Private Sub class_initialize()
    Set mlo = tbAsc
End Sub

Public Sub save()
     
     Dim lr As ListRow
     
     If Me.getID = 0 Then 'Caso seja novo Asc
        ID = getAutoNumerateID(mlo)
        data = Date
        hora = Time
        Ascs.Add Me, CStr(Me.getID)
        Set lr = mlo.ListRows.Add
    Else
        Set lr = mlo.ListRows(getLinha(mlo, "ID", Me.getID))
        With Ascs(CStr(Me.getID))
            .letNome = Me.getNome
            .letCpf = Me.getCpf
            .letAsc = Me.getAsc
            .letMotivo = Me.getMotivo
        End With
     End If
     
     lr.Range.value = toListRow
     
End Sub

Private Function toListRow() As Variant()

    Dim vaReturn(0 To 0, 0 To 6) As Variant
    
    vaReturn(0, getColuna(mlo, "ID") - 1) = Me.getID
    vaReturn(0, getColuna(mlo, "data") - 1) = Me.getData
    vaReturn(0, getColuna(mlo, "hora") - 1) = Me.getHora
    vaReturn(0, getColuna(mlo, "nome") - 1) = Me.getNome
    vaReturn(0, getColuna(mlo, "cpf") - 1) = Me.getCpf
    vaReturn(0, getColuna(mlo, "asc") - 1) = Me.getAsc
    vaReturn(0, getColuna(mlo, "motivo") - 1) = Me.getMotivo
    
    toListRow = vaReturn
    
End Function

Public Sub load(getID As Long)
    Dim lRow As Long
    
    lRow = getLinha(mlo, "ID", getID)
    
    'propriedades
    ID = mlo.DataBodyRange(lRow, getColuna(mlo, "ID"))
    data = mlo.DataBodyRange(lRow, getColuna(mlo, "data"))
    hora = mlo.DataBodyRange(lRow, getColuna(mlo, "hora"))
    nome = mlo.DataBodyRange(lRow, getColuna(mlo, "nome"))
    cpf = mlo.DataBodyRange(lRow, getColuna(mlo, "cpf"))
    ASC = mlo.DataBodyRange(lRow, getColuna(mlo, "asc"))
    motivo = mlo.DataBodyRange(lRow, getColuna(mlo, "motivo"))
    
End Sub

Public Sub delete()
    If Me.getID <> 0 Then
        Ascs.Remove CStr(Me.getID)
        mlo.ListRows(getLinha(mlo, "ID", Me.getID)).delete
    End If
End Sub

