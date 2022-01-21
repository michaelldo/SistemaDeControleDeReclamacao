VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formEditarAsc 
   Caption         =   "Editar Asc"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5235
   OleObjectBlob   =   "formEditarAsc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formEditarAsc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moAsc As ASC
Private cancelado As Boolean

Public Property Get getCancelado() As Boolean
    getCancelado = cancelado
End Property

Public Sub Init(Optional o As ASC)
    cancelado = False
    Set moAsc = o
    CarregarDados
    Me.Show
End Sub

Private Sub CarregarDados()
    If moAsc Is Nothing Then
        Set moAsc = New ASC
        lbData.Caption = Date
        lbHora.Caption = Time
        txtNome.Text = ""
        txtCpf.Text = ""
        txtAsc.Text = ""
        txtMotivo.Text = ""
    Else
        lbID.Caption = moAsc.getID
        lbData.Caption = moAsc.getData
        lbHora.Caption = moAsc.getHora
        txtNome.Text = moAsc.getNome
        txtCpf.Text = moAsc.getCpf
        txtAsc.Text = moAsc.getAsc
        txtMotivo.Text = moAsc.getMotivo
    End If
End Sub

Private Sub btnCancelar_Click()
    cancelado = True
    Me.Hide
End Sub

Private Sub btnSalvar_Click()
    If Validar Then
        moAsc.letNome = txtNome.Text
        moAsc.letCpf = txtCpf.Text
        moAsc.letAsc = txtAsc.Text
        moAsc.letMotivo = txtMotivo.Text
        moAsc.save
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
