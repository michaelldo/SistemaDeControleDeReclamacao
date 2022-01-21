VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formEditarOuvidoria 
   Caption         =   "OUVIDORIA"
   ClientHeight    =   9720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6540
   OleObjectBlob   =   "formEditarOuvidoria.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formEditarOuvidoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moOuvidoria As OUVIDORIA
Private cancelado As Boolean

Public Property Get getCancelado() As Boolean
    getCancelado = cancelado
End Property

Public Sub Init(Optional o As OUVIDORIA)
    cancelado = False
    Set moOuvidoria = o
    CarregarTipoStatus
    CarregarFormulario
    CarregarDados
    Me.Show
End Sub

Private Sub CarregarTipoStatus()

    Dim o As tipoStatus
    
    With cbStatus
        .Clear

        For Each o In TiposStatus
            .AddItem
            .List(.ListCount - 1, 0) = o.getID
            .List(.ListCount - 1, 1) = o.getNome
        Next o
        
    End With
    
    cbStatus.ListIndex = 1

End Sub

Private Sub CarregarFormulario()

    Dim i As Informante
    Dim t As Tipo
    Dim u As Uf
    
    
    With cbInformante
        .Clear

        For Each i In Informantes
            .AddItem
            .List(.ListCount - 1, 0) = i.getID
            .List(.ListCount - 1, 1) = i.getNome
        Next i
        
    End With
    
    With cbUf
        .Clear

        For Each u In Ufs
            .AddItem
            .List(.ListCount - 1, 0) = u.getID
            .List(.ListCount - 1, 1) = u.getNome
        Next u
        
    End With
    
    With cbTipo1
        .Clear

        For Each t In Tipos
            .AddItem
            .List(.ListCount - 1, 0) = t.getID
            .List(.ListCount - 1, 1) = t.getNome
        Next t
        
    End With
    
    With cbTipo2
        .Clear

        For Each t In Tipos
            .AddItem
            .List(.ListCount - 1, 0) = t.getID
            .List(.ListCount - 1, 1) = t.getNome
        Next t
        
    End With
    
    With cbTipo3
        .Clear

        For Each t In Tipos
            .AddItem
            .List(.ListCount - 1, 0) = t.getID
            .List(.ListCount - 1, 1) = t.getNome
        Next t
        
    End With

End Sub

Private Sub CarregarDados()
    If moOuvidoria Is Nothing Then
        Set moOuvidoria = New OUVIDORIA
        lbData.Caption = Date
        lbHora.Caption = Time
        txtNome.Text = ""
        txtCpf.Text = ""
        txtEmail.Text = ""
        cbInformante.value = ""
        txtCep.Text = ""
        txtEndreco.Text = ""
        txtNumero.Text = ""
        txtComplemento.Text = ""
        cbUf.value = ""
        txtBairro.Text = ""
        txtCidade.Text = ""
        txtTelefone1.Text = ""
        cbTipo1.value = ""
        txtTelefone2.Text = ""
        cbTipo2.value = ""
        txtTelefone3.Text = ""
        cbTipo3.value = ""
        txtAlmope.Text = ""
        txtProtocolo.Text = ""
        txtOuvidoria.Text = ""
        cbStatus.value = ""
    Else
        lbID.Caption = moOuvidoria.getID
        lbData.Caption = moOuvidoria.getData
        lbHora.Caption = moOuvidoria.getHora
        txtNome.Text = moOuvidoria.getNome
        txtCpf.Text = moOuvidoria.getCpf
        txtEmail.Text = moOuvidoria.getEmail
        cbInformante.Text = moOuvidoria.getInformante.getNome
        txtCep.Text = moOuvidoria.getCep
        txtEndreco.Text = moOuvidoria.getEndereco
        txtNumero.Text = moOuvidoria.getNumero
        txtComplemento.Text = moOuvidoria.getComplemento
        cbUf.Text = moOuvidoria.getUf.getNome
        txtBairro.Text = moOuvidoria.getBairro
        txtCidade.Text = moOuvidoria.getCidade
        txtTelefone1.Text = moOuvidoria.getTelefone1
        cbTipo1.Text = moOuvidoria.getTipo1.getNome
        txtTelefone2.Text = moOuvidoria.getTelefone2
        cbTipo2.Text = moOuvidoria.getTipo2.getNome
        txtTelefone3.Text = moOuvidoria.getTelefone3
        cbTipo3.Text = moOuvidoria.getTipo3.getNome
        txtAlmope.Text = moOuvidoria.getAlmope
        txtProtocolo.Text = moOuvidoria.getProtocolo
        txtOuvidoria.Text = moOuvidoria.getOuvidoria
        cbStatus.Text = moOuvidoria.getStatus.getNome
    End If
End Sub

Private Sub btnCancelar_Click()
    cancelado = True
    Me.Hide
End Sub

Private Sub btnSalvar_Click()
    If Validar Then
        moOuvidoria.letNome = txtNome.Text
        moOuvidoria.letCpf = txtCpf.Text
        moOuvidoria.letEmail = txtEmail.Text
        Set moOuvidoria.letInformante = Informantes(CStr(cbInformante.List(cbInformante.ListIndex, 0)))
        moOuvidoria.letCep = txtCep.Text
        moOuvidoria.letEndereco = txtEndreco.Text
        moOuvidoria.letNumero = txtNumero.Text
        moOuvidoria.letComplemento = txtComplemento.Text
        Set moOuvidoria.letUf = Ufs(CStr(cbUf.List(cbUf.ListIndex, 0)))
        moOuvidoria.letBairro = txtBairro.Text
        moOuvidoria.letCidade = txtCidade.Text
        moOuvidoria.letTelefone1 = txtTelefone1.Text
        Set moOuvidoria.letTipo1 = Tipos(CStr(cbTipo1.List(cbTipo1.ListIndex, 0)))
        moOuvidoria.letTelefone2 = txtTelefone2.Text
        Set moOuvidoria.letTipo2 = Tipos(CStr(cbTipo2.List(cbTipo2.ListIndex, 0)))
        moOuvidoria.letTelefone3 = txtTelefone3.Text
        Set moOuvidoria.letTipo3 = Tipos(CStr(cbTipo3.List(cbTipo3.ListIndex, 0)))
        moOuvidoria.letAlmope = txtAlmope.Text
        moOuvidoria.letProtocolo = txtProtocolo.Text
        moOuvidoria.letOuvidoria = txtOuvidoria.Text
        Set moOuvidoria.letStatus = TiposStatus(CStr(cbStatus.List(cbStatus.ListIndex, 0)))
        moOuvidoria.save
        Me.Hide
    End If
End Sub

Private Function Validar()

    If Len(txtCpf.Text) < 11 Then
        MsgBox "Informe o cpf corretamente", vbCritical
        txtCpf.SetFocus
        Exit Function
    End If

    Validar = True

End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = 1
    btnCancelar_Click
End Sub
