Attribute VB_Name = "mTeste"
'Option Explicit
'---------------------------------------------------------------------------------------------
'    Sub teste()
'
'        Dim lo As ListObject
'        Dim coluna As Long
'
'        Set lo = wsDatabase.ListObjects("tbOuvidoria")
'
'        Debug.Print getColuna(lo, "tipo3")
'
'    End Sub
'---------------------------------------------------------------------------------------------
'    Sub testeLista()
'
'        Debug.Print getLinha(tbStatus, "nome", "Reversão")
'
'    End Sub
'
'    Sub teste_auto()
'
'        Debug.Print getAutoNumerateID(tbStatus)
'
'    End Sub
'---------------------------------------------------------------------------------------------
'     Sub testeClasse()
'
'        Dim o As tipoStatus
'        Set o = New tipoStatus
'
'            o.load 4
'
'        Debug.Print o.getNome & " " & o.getID
'
'    End Sub
'---------------------------------------------------------------------------------------------
'    Sub testeDelete()
'
'         Dim o As tipoStatus
'         Set o = New tipoStatus
'
'        o.load 5
'        o.delete
'
'    End Sub
'---------------------------------------------------------------------------------------------
'    Sub testeCol()
'
'        Dim o As tipoStatus
'
'        Set o = TiposStatus("4")
'        Debug.Print TiposStatus.Count
'
'    End Sub
'---------------------------------------------------------------------------------------------
'    Sub testeCol2()
'
'        Dim o As tipoStatus
'
'        Debug.Print "Antes: " & TiposStatus.Count
'
'        Set o = New tipoStatus
'
'        o.letNome = "teste"
'        o.save
'
'        Debug.Print "Depois: " & TiposStatus.Count
'
'    End Sub
'---------------------------------------------------------------------------------------------
'    Sub testarload()
'        Dim o As Uf
'        Set o = New Uf
'
'        Debug.Print Ufs.Count
'
'        'o.letNome = "testando"
       'o.save

        'Debug.Print "Depois:"; TiposStatus.Count

'    End Sub

Sub teste1()
    Dim o As Informante
    Set o = New Informante
    
    Debug.Print Informantes.Count
End Sub
