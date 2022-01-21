VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formOuvidoria 
   Caption         =   "OUVIDORIA"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11520
   OleObjectBlob   =   "formOuvidoria.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formOuvidoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private enableEvents As Boolean

Public Property Get getEnableEvents() As Boolean
    getEnableEvents = enableEvents
End Property

Public Property Let letEnableEvents(sEnableEvents As Boolean)
    enableEvents = sEnableEvents
End Property

Private Sub btnCadastrar_Click()
    Dim fOuvidoria As formEditarOuvidoria
    Dim ID As Long
    
    Set fOuvidoria = formEditarOuvidoria
    
    Me.Hide
    
    fOuvidoria.Init
    
    If Not fOuvidoria.getCancelado Then CarregarOuvidoria
    
    Unload fOuvidoria
    
    Me.Show
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub


Private Sub btnEditar_Click()
    Dim fOuvidoria As formEditarOuvidoria
    Dim ID As Long
    
    If listOuvidoria.ListIndex < 0 Then
        MsgBox "Selecione uma Ouvidoria para editar", vbCritical
        Exit Sub
    End If
    
    With listOuvidoria
        ID = .List(.ListIndex, 0)
    End With
    
    Set fOuvidoria = formEditarOuvidoria
    
    Me.Hide
    
    fOuvidoria.Init Ouvidorias(CStr(ID))
    
    If Not fOuvidoria.getCancelado Then CarregarOuvidoria
    
    Unload fOuvidoria
    
    Me.Show
    
End Sub

Private Sub btnExcluir_Click()
    Dim bQuestao As VbMsgBoxResult
    Dim ID As Long
    Dim sData As String
    Dim sHora As String
    Dim sNome As String
    Dim sCpf As String
    Dim sEmail As String
    Dim sInformante As Informante
    Dim sCep As String
    Dim sEndereco As String
    Dim sNumero As String
    Dim sComplemento As String
    Dim sBairro As String
    Dim sCidade As String
    Dim sUf As Uf
    Dim sTelefone1 As String
    Dim sTipo1 As Tipo
    Dim sTelefone2 As String
    Dim sTipo2 As Tipo
    Dim sTelefone3 As String
    Dim sTipo3 As Tipo
    Dim sAlmope As Long
    Dim sProtocolo As Long
    Dim sOuvidoria As String
    Dim sStatus As tipoStatus
    
    If listOuvidoria.ListIndex < 0 Then
        MsgBox "Selecione uma Ouvidoria para excluir", vbCritical
        Exit Sub
    End If
    
    With listOuvidoria
       ID = .List(.ListIndex, 0)
       sData = .List(.ListIndex, 1)
       sHora = .List(.ListIndex, 2)
       sNome = .List(.ListIndex, 3)
       sCpf = .List(.ListIndex, 4)
       sEmail = .List(.ListIndex, 5)
       sInformante = .List(.ListIndex, 6)
       sCep = .List(.ListIndex, 7)
       sEndereco = .List(.ListIndex, 8)
       sNumero = .List(.ListIndex, 9)
       sComplemento = .List(.ListIndex, 10)
       sBairro = .List(.ListIndex, 11)
       sCidade = .List(.ListIndex, 12)
       sUf = .List(.ListIndex, 13)
       sTelefone1 = .List(.ListIndex, 14)
       sTipo1 = .List(.ListIndex, 15)
       sTelefone2 = .List(.ListIndex, 16)
       sTipo2 = .List(.ListIndex, 17)
       sTelefone3 = .List(.ListIndex, 18)
       sTipo3 = .List(.ListIndex, 19)
       sAlmope = .List(.ListIndex, 20)
       sProtocolo = .List(.ListIndex, 21)
       sOuvidoria = .List(.ListIndex, 22)
       sStatus = .List(.ListIndex, 23)
    End With
    
    bQuestao = MsgBox("Tem certeza que deseja cancelar a Ouvidoria '" & sProtocolo & "'?", vbQuestion + vbYesNo)
    
    If bQuestao = vbNo Then Exit Sub
    
    Set oOuvidoria = New OUVIDORIA
    
        oOuvidoria.load ID
        oOuvidoria.delete
        
        CarregarOuvidoria
    
End Sub

Private Sub cbStatus_Change()
    If Me.getEnableEvents Then CarregarOuvidoria
End Sub

Private Sub UserForm_Initialize()
    Me.letEnableEvents = False
    CarregarTipoStatus
    CarregarDatas
    CarregarOuvidoria
    Me.letEnableEvents = True
End Sub
Private Sub CarregarDatas()
    
    Dim o As OUVIDORIA
    Dim x As OUVIDORIA
    
    With cmbData1
        .Clear
        .AddItem
        .List(.ListCount - 1, 0) = 0
        .List(.ListCount - 1, 1) = "-Selecione-"
        For Each o In Ouvidorias
            .AddItem
            .List(.ListCount - 1, 0) = o.getID
            .List(.ListCount - 1, 1) = o.getData
        Next o
        
    End With
    
    With cmbData2
        .Clear
        .AddItem
        .List(.ListCount - 1, 0) = 0
        .List(.ListCount - 1, 1) = "-Selecione-"
        For Each x In Ouvidorias
            .AddItem
            .List(.ListCount - 1, 0) = x.getID
            .List(.ListCount - 1, 1) = x.getData
        Next x
        
    End With
    
    cmbData1.ListIndex = 0
    cmbData2.ListIndex = 0
End Sub

Private Sub CarregarTipoStatus()

    Dim o As tipoStatus
    
    With cbStatus
        .Clear
        .AddItem
        .List(.ListCount - 1, 0) = 0
        .List(.ListCount - 1, 1) = "Todos os Status"
        For Each o In TiposStatus
            .AddItem
            .List(.ListCount - 1, 0) = o.getID
            .List(.ListCount - 1, 1) = o.getNome
        Next o
        
    End With
    
    cbStatus.ListIndex = 0

End Sub


Private Sub CarregarOuvidoria()
    Dim o As OUVIDORIA
    Dim tipoStatusID As Long
    Dim informanteID As Long
    Dim tipoID As Long
    Dim ufID As Long
    
     
    tipoStatusID = cbStatus.List(cbStatus.ListIndex, 0)
    
        With listOuvidoria
            .Clear
            For Each o In Ouvidorias
                If tipoStatusID = 0 Or tipoStatusID = o.getStatus.getID Then
                    .AddItem
                    .List(.ListCount - 1, 0) = o.getID
                    .List(.ListCount - 1, 1) = o.getData
                    '.List(.ListCount - 1, 2) = o.getHora
                    .List(.ListCount - 1, 2) = o.getNome
                    .List(.ListCount - 1, 3) = o.getCpf
                    '.List(.ListCount - 1, 5) = o.getEmail
                    .List(.ListCount - 1, 4) = o.getInformante.getNome
                    '.List(.ListCount - 1, 7) = o.getCep
                    '.List(.ListCount - 1, 8) = o.getEndereco
                    '.List(.ListCount - 1, 9) = o.getNumero
                    '.List(.ListCount - 1, 10) = o.getComplemento
                    '.List(.ListCount - 1, 11) = o.getUf.getNome
                    '.List(.ListCount - 1, 6) = o.getBairro
                    .List(.ListCount - 1, 5) = o.getCidade
                    '.List(.ListCount - 1, 14) = o.getTelefone1
                    '.List(.ListCount - 1, 15) = o.getTipo1.getNome
                    '.List(.ListCount - 1, 16) = o.getTelefone2
                    '.List(.ListCount - 1, 17) = o.getTipo2.getNome
                    '.List(.ListCount - 1, 18) = o.getTelefone3
                    ''.List(.ListCount - 1, 19) = o.getTipo3.getNome
                    '.List(.ListCount - 1, 20) = o.getAlmope
                    .List(.ListCount - 1, 6) = o.getProtocolo
                    .List(.ListCount - 1, 7) = o.getStatus.getNome
                    .List(.ListCount - 1, 8) = o.getOuvidoria
                End If
            Next o
        End With
End Sub
