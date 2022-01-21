Attribute VB_Name = "mglobals"
Option Explicit

Public gColTiposStatus As Collection
Public gColAscs As Collection
Public gColOuvidorias As Collection
Public gColUfs As Collection
Public gColInformantes As Collection
Public gColTipos As Collection

Public Function TiposStatus() As Collection
    
    Dim oTipoStatus As tipoStatus
    Dim lo As ListObject
    Dim lr As ListRow
    
    
   If gColTiposStatus Is Nothing Then
   
        Set lo = tbStatus
        Set gColTiposStatus = New Collection
    
        For Each lr In lo.ListRows
            Set oTipoStatus = New tipoStatus
            oTipoStatus.load lr.Range(getColuna(lo, "ID"))
            gColTiposStatus.Add oTipoStatus, CStr(oTipoStatus.getID)
        Next lr
   End If
   
   Set TiposStatus = gColTiposStatus
   
End Function

Public Function Ascs() As Collection
    
    Dim oAsc As ASC
    Dim lo As ListObject
    Dim lr As ListRow
    
    
   If gColAscs Is Nothing Then
   
        Set lo = tbAsc
        Set gColAscs = New Collection
    
        For Each lr In lo.ListRows
            Set oAsc = New ASC
            oAsc.load lr.Range(getColuna(lo, "ID"))
            gColAscs.Add oAsc, CStr(oAsc.getID)
        Next lr
   End If
   
   Set Ascs = gColAscs
   
End Function

Public Function Ouvidorias() As Collection
    
    Dim oOuvidoria As OUVIDORIA
    Dim lo As ListObject
    Dim lr As ListRow
    
    
   If gColOuvidorias Is Nothing Then
   
        Set lo = tbOuvidoria
        Set gColOuvidorias = New Collection
    
        For Each lr In lo.ListRows
            Set oOuvidoria = New OUVIDORIA
            oOuvidoria.load lr.Range(getColuna(lo, "ID"))
            gColOuvidorias.Add oOuvidoria, CStr(oOuvidoria.getID)
        Next lr
   End If
   
   Set Ouvidorias = gColOuvidorias
   
End Function

Public Function Informantes() As Collection
    
    Dim oInformante As Informante
    Dim lo As ListObject
    Dim lr As ListRow
    
    
   If gColInformantes Is Nothing Then
   
        Set lo = tbInformante
        
        Set gColInformantes = New Collection
    
        For Each lr In lo.ListRows
            Set oInformante = New Informante
            oInformante.load lr.Range(getColuna(lo, "ID"))
            gColInformantes.Add oInformante, CStr(oInformante.getID)
        Next lr
   End If
   
   Set Informantes = gColInformantes
   
End Function

Public Function Tipos() As Collection
    
    Dim oTipo As Tipo
    Dim lo As ListObject
    Dim lr As ListRow
    
    
   If gColTipos Is Nothing Then
   
        Set lo = tbTipo
        Set gColTipos = New Collection
    
        For Each lr In lo.ListRows
            Set oTipo = New Tipo
            oTipo.load lr.Range(getColuna(lo, "ID"))
            gColTipos.Add oTipo, CStr(oTipo.getID)
        Next lr
   End If
   
   Set Tipos = gColTipos
   
End Function

Public Function Ufs() As Collection
    
    Dim oUf As Uf
    Dim lo As ListObject
    Dim lr As ListRow
    
    
   If gColUfs Is Nothing Then
   
        Set lo = tbUf
        Set gColUfs = New Collection
    
        For Each lr In lo.ListRows
            Set oUf = New Uf
            oUf.load lr.Range(getColuna(lo, "ID"))
            gColUfs.Add oUf, CStr(oUf.getID)
        Next lr
   End If
   
   Set Ufs = gColUfs
   
End Function
