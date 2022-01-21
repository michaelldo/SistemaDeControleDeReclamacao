VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formAsc 
   Caption         =   "ASC"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9270
   OleObjectBlob   =   "formAsc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formAsc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCadastrar_Click()
    Dim fAsc As formEditarAsc
    Dim ID As Long
    
    Set fAsc = formEditarAsc
    
    Me.Hide
    
    fAsc.Init
    
    If Not fAsc.getCancelado Then CarregarAsc
    
    Unload fAsc
    
    Me.Show
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnEditar_Click()
    
    Dim fAsc As formEditarAsc
    Dim ID As Long
    
    If lbAsc.ListIndex < 0 Then
        MsgBox "Selecione uma categoria para editar", vbCritical
        Exit Sub
    End If
    
    With lbAsc
        ID = .List(.ListIndex, 0)
    End With
    
    Set fAsc = formEditarAsc
    
    Me.Hide
    
    fAsc.Init Ascs(CStr(ID))
    
    If Not fAsc.getCancelado Then CarregarAsc
    
    Unload fAsc
    
    Me.Show
    
End Sub

Private Sub btnExcluir_Click()
    Dim bQuestao As VbMsgBoxResult
    Dim sNome As String
    Dim sData As String
    Dim sHora As String
    Dim sCpf As String
    Dim sAsc As String
    Dim sMotivo As String
    Dim ID As Long
    
    If lbAsc.ListIndex < 0 Then
        MsgBox "Selecione um ASC para excluir", vbCritical
        Exit Sub
    End If
    
    With lbAsc
        ID = .List(.ListIndex, 0)
        sData = .List(.ListIndex, 1)
        sHora = .List(.ListIndex, 2)
        sNome = .List(.ListIndex, 3)
        sCpf = .List(.ListIndex, 4)
        sAsc = .List(.ListIndex, 5)
        sMotivo = .List(.ListIndex, 6)
     End With
     
     bQuestao = MsgBox("Tem certeza que deseja cancelar o ASC '" & sAsc & "'?", vbQuestion + vbYesNo)
     
     If bQuestao = vbNo Then Exit Sub
     
     Set oAsc = New ASC
     
        oAsc.load ID
        oAsc.delete
        
        CarregarAsc
     
End Sub

Private Sub UserForm_Initialize()
    CarregarAsc
End Sub

Private Sub CarregarAsc()
    Dim o As ASC
    
    lbAsc.Clear
    
    For Each o In Ascs
        With lbAsc
            .AddItem
            .List(.ListCount - 1, 0) = o.getID
            .List(.ListCount - 1, 1) = o.getData
            .List(.ListCount - 1, 2) = o.getHora
            .List(.ListCount - 1, 3) = o.getNome
            .List(.ListCount - 1, 4) = o.getCpf
            .List(.ListCount - 1, 5) = o.getAsc
            .List(.ListCount - 1, 6) = o.getMotivo
        End With
    Next o
    
End Sub
