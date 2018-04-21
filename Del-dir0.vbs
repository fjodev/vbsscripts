' ********************************************************************
' * Ficheiro:          Del-dir0.vbs                                  *
' * Função:            Verifica o tamanho de todas sub-directoria    *
' *                    e acaso tenha tamanho 0 paga-as               *
' * Sintase:           Del-dir0.vbs drive:caminho\nome_da_directoria *
' * Autor:             Fernando Oliveira                             *
' * Data:              2008-03-11                                    *
' ********************************************************************

' ************************************************
' ** Variaveis, constantes e outras declarações **
' ************************************************

On Error Resume Next

If Wscript.Arguments.Count > 0 Then
    strDirBase = Wscript.Arguments.Item(0)
Else
    Wscript.Echo "Del-dir0.vbs: deve indicar uma directoria"
    Wscript.Quit
End If

eNaoSistema = true
intTamanho = 0
dim strDirChk

intListDir = 0
intDelDir = 0

strLinha = "------------------------------------------"

strComputer = "."

' *********************************************************
' ** Primerio Verifica se a directoria de base existe    **
' *********************************************************

Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(strDirBase) Then
  
    Set objFolder = objFSO.GetFolder(strDirBase)
    If objFolder.IsRootFolder then
       Wscript.Echo
       Wscript.Echo "Del-dir0.vbs: directoria invalida."
       Wscript.Echo "Por segurança, este script não pode ser executada na directoria raiz."
       Wscript.Quit
    End If
    Set colSubfolders = objFolder.Subfolders
    
    ' ************************************************
    ' ** Ciclo principal                            **
    ' ************************************************

    For Each objSubFolder in colSubfolders
        strDirChk = objSubFolder.Path
        intListDir = intListDir + 1
        intTamanho = objSubFolder.Size
        objA=objSubFolder.Attributes
        If objA AND 2 Then
            eEscondido = true
        End If    
        If objA AND 4 Then
            eNaoSistema = false
        End If
        If intTamanho = 0 AND eNaoSistema Then
            objFSO.DeleteFolder(strDirChk)
            If intDelDir = 0 Then
               Wscript.Echo
               Wscript.Echo "Directorias removidas em " & strDirBase & ": "
               Wscript.Echo strLinha
            End If
            Wscript.Echo strDirChk
            intDelDir = intDelDir + 1
        End If
    Next
    Wscript.Echo
    Wscript.Echo intListDir & " directoria(s) verificada(s)"
    Wscript.Echo intDelDir & " directoria(s) removida(s)"
Else
    Wscript.Echo "Del-dir0.vbs: directoria " & strDirBase & " não existe."
End If
