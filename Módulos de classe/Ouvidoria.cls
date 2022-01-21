VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ouvidoria"
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
Private email As String
Private Informante As Informante
Private cep As String
Private endereco As String
Private numero As String
Private complemento As String
Private bairro As String
Private cidade As String
Private Uf As Uf
Private telefone1 As String
Private tipo1 As Tipo
Private telefone2 As String
Private tipo2 As Tipo
Private telefone3 As String
Private tipo3 As Tipo
Private almope As Long
Private protocolo As Long
Private OUVIDORIA As String
Private status As tipoStatus
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

Public Property Get getEmail() As String
    getEmail = email
End Property

Public Property Let letEmail(sEmail As String)
    email = sEmail
End Property

Public Property Get getInformante() As Informante
    Set getInformante = Informante
End Property

Public Property Set letInformante(sInformante As Informante)
    Set Informante = sInformante
End Property

Public Property Get getCep() As String
    getCep = cep
End Property

Public Property Let letCep(sCep As String)
    cep = sCep
End Property
    
Public Property Get getEndereco() As String
    getEndereco = endereco
End Property

Public Property Let letEndereco(sEndereco As String)
    endereco = sEndereco
End Property

Public Property Get getNumero() As String
    getNumero = numero
End Property

Public Property Let letNumero(sNumero As String)
    numero = sNumero
End Property

Public Property Get getComplemento() As String
    getComplemento = complemento
End Property

Public Property Let letComplemento(sComplemento As String)
    complemento = sComplemento
End Property

Public Property Get getBairro() As String
    getBairro = bairro
End Property

Public Property Let letBairro(sBairro As String)
    bairro = sBairro
End Property

Public Property Get getCidade() As String
    getCidade = cidade
End Property

Public Property Let letCidade(sCidade As String)
    cidade = sCidade
End Property

Public Property Get getUf() As Uf
    Set getUf = Uf
End Property

Public Property Set letUf(sUf As Uf)
    Set Uf = sUf
End Property

Public Property Get getTelefone1() As String
    getTelefone1 = telefone1
End Property

Public Property Let letTelefone1(sTelefone1 As String)
    telefone1 = sTelefone1
End Property

Public Property Get getTipo1() As Tipo
    Set getTipo1 = tipo1
End Property

Public Property Set letTipo1(sTipo1 As Tipo)
    Set tipo1 = sTipo1
End Property

Public Property Get getTelefone2() As String
    getTelefone2 = telefone2
End Property

Public Property Let letTelefone2(sTelefone2 As String)
    telefone2 = sTelefone2
End Property

Public Property Get getTipo2() As Tipo
    Set getTipo2 = tipo2
End Property

Public Property Set letTipo2(sTipo2 As Tipo)
    Set tipo2 = sTipo2
End Property

Public Property Get getTelefone3() As String
    getTelefone3 = telefone3
End Property

Public Property Let letTelefone3(sTelefone3 As String)
     telefone3 = sTelefone3
End Property

Public Property Get getTipo3() As Tipo
    Set getTipo3 = tipo3
End Property

Public Property Set letTipo3(sTipo3 As Tipo)
    Set tipo3 = sTipo3
End Property

Public Property Get getAlmope() As Long
    getAlmope = almope
End Property

Public Property Let letAlmope(sAlmope As Long)
    almope = sAlmope
End Property

Public Property Get getProtocolo() As Long
    getProtocolo = protocolo
End Property

Public Property Let letProtocolo(sProtocolo As Long)
    protocolo = sProtocolo
End Property

Public Property Get getOuvidoria() As String
    getOuvidoria = OUVIDORIA
End Property

Public Property Let letOuvidoria(sOuvidoria As String)
    OUVIDORIA = sOuvidoria
End Property

Public Property Get getStatus() As tipoStatus
    Set getStatus = status
End Property

Public Property Set letStatus(sStatus As tipoStatus)
    Set status = sStatus
End Property

Public Sub save()
    Dim lr As ListRow
    
    If Me.getID = 0 Then
        ID = getAutoNumerateID(mlo)
        data = Date
        hora = Time
        Ouvidorias.Add Me, CStr(Me.getID)
        Set lr = mlo.ListRows.Add
    Else
        Set lr = mlo.ListRows(getLinha(mlo, "ID", Me.getID))
        
        With Ouvidorias(CStr(Me.getID))
            .letNome = Me.getNome
            .letCpf = Me.getCpf
            .letEmail = Me.getEmail
             Set .letInformante = Me.getInformante
            .letCep = Me.getCep
            .letEndereco = Me.getEndereco
            .letNumero = Me.getNumero
            .letComplemento = Me.getComplemento
            .letBairro = Me.getBairro
            .letCidade = Me.getCidade
            Set .letUf = Me.getUf
            .letTelefone1 = Me.getTelefone1
            Set .letTipo1 = Me.getTipo1
            .letTelefone2 = Me.getTelefone2
            Set .letTipo2 = Me.getTipo2
            .letTelefone3 = Me.getTelefone3
            Set .letTipo3 = Me.getTipo3
            .letAlmope = Me.getAlmope
            .letProtocolo = Me.getProtocolo
            .letOuvidoria = Me.getOuvidoria
             Set .letStatus = Me.getStatus
        End With
    End If
    
    lr.Range.value = toListRow
    
End Sub

Private Function toListRow() As Variant()

    Dim vaReturn(0 To 0, 0 To 23) As Variant
    
    vaReturn(0, getColuna(mlo, "ID") - 1) = Me.getID
    vaReturn(0, getColuna(mlo, "data") - 1) = Me.getData
    vaReturn(0, getColuna(mlo, "hora") - 1) = Me.getHora
    vaReturn(0, getColuna(mlo, "nome") - 1) = Me.getNome
    vaReturn(0, getColuna(mlo, "cpf") - 1) = Me.getCpf
    vaReturn(0, getColuna(mlo, "email") - 1) = Me.getEmail
    vaReturn(0, getColuna(mlo, "informante") - 1) = Me.getInformante.getID
    vaReturn(0, getColuna(mlo, "cep") - 1) = Me.getCep
    vaReturn(0, getColuna(mlo, "endereco") - 1) = Me.getEndereco
    vaReturn(0, getColuna(mlo, "numero") - 1) = Me.getNumero
    vaReturn(0, getColuna(mlo, "complemento") - 1) = Me.getComplemento
    vaReturn(0, getColuna(mlo, "bairro") - 1) = Me.getBairro
    vaReturn(0, getColuna(mlo, "cidade") - 1) = Me.getCidade
    vaReturn(0, getColuna(mlo, "uf") - 1) = Me.getUf.getID
    vaReturn(0, getColuna(mlo, "telefone1") - 1) = Me.getTelefone1
    vaReturn(0, getColuna(mlo, "tipo1") - 1) = Me.getTipo1.getID
    vaReturn(0, getColuna(mlo, "telefone2") - 1) = Me.getTelefone2
    vaReturn(0, getColuna(mlo, "tipo2") - 1) = Me.getTipo2.getID
    vaReturn(0, getColuna(mlo, "telefone3") - 1) = Me.getTelefone3
    vaReturn(0, getColuna(mlo, "tipo3") - 1) = Me.getTipo3.getID
    vaReturn(0, getColuna(mlo, "almope") - 1) = Me.getAlmope
    vaReturn(0, getColuna(mlo, "protocolo") - 1) = Me.getProtocolo
    vaReturn(0, getColuna(mlo, "ouvidoria") - 1) = Me.getOuvidoria
    vaReturn(0, getColuna(mlo, "status") - 1) = Me.getStatus.getID
    
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
    email = mlo.DataBodyRange(lRow, getColuna(mlo, "email"))
    Set Informante = Informantes(CStr(mlo.DataBodyRange(lRow, getColuna(mlo, "informante"))))
    cep = mlo.DataBodyRange(lRow, getColuna(mlo, "cep"))
    endereco = mlo.DataBodyRange(lRow, getColuna(mlo, "endereco"))
    numero = mlo.DataBodyRange(lRow, getColuna(mlo, "numero"))
    complemento = mlo.DataBodyRange(lRow, getColuna(mlo, "complemento"))
    bairro = mlo.DataBodyRange(lRow, getColuna(mlo, "bairro"))
    cidade = mlo.DataBodyRange(lRow, getColuna(mlo, "cidade"))
    Set Uf = Ufs(CStr(mlo.DataBodyRange(lRow, getColuna(mlo, "uf"))))
    telefone1 = mlo.DataBodyRange(lRow, getColuna(mlo, "telefone1"))
    Set tipo1 = Tipos(CStr(mlo.DataBodyRange(lRow, getColuna(mlo, "tipo1"))))
    telefone2 = mlo.DataBodyRange(lRow, getColuna(mlo, "telefone2"))
    Set tipo2 = Tipos(CStr(mlo.DataBodyRange(lRow, getColuna(mlo, "tipo2"))))
    telefone3 = mlo.DataBodyRange(lRow, getColuna(mlo, "telefone3"))
    Set tipo3 = Tipos(CStr(mlo.DataBodyRange(lRow, getColuna(mlo, "tipo3"))))
    almope = mlo.DataBodyRange(lRow, getColuna(mlo, "almope"))
    protocolo = mlo.DataBodyRange(lRow, getColuna(mlo, "protocolo"))
    OUVIDORIA = mlo.DataBodyRange(lRow, getColuna(mlo, "ouvidoria"))
    Set status = TiposStatus(CStr(mlo.DataBodyRange(lRow, getColuna(mlo, "status"))))
    
End Sub

Public Sub delete()
    If Me.getID <> 0 Then
        Ouvidorias.Remove CStr(Me.getID)
        mlo.ListRows(getLinha(mlo, "ID", Me.getID)).delete
    End If
End Sub

Private Sub class_initialize()
    Set mlo = tbOuvidoria
End Sub
