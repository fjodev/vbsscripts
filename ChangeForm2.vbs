Sub ChangeForm()
   Dim ChangedItems As Integer
   ChangedItems = 0
   
   ' Dados da pasta
   Set Curfolder = Application.ActiveExplorer.CurrentFolder
   Set AllItems = Curfolder.Items
   Numitems = Curfolder.Items.Count

   
   ' Confirma acção
   resposta = MsgBox("Serão alterados " & Numitems & " itens na pasta " _
    & Curfolder & ". Deseja continuar?", _
    vbYesNo, "Mudar para " & NewForm)
   If resposta = vbYes Then
       ' Definição do novo formulário
       usernewform = InputBox("Novo formulário", "Alteração de formulário")
       If usernewform = "" Then
           End
       End If
   
       NewForm = "IPM." & LTrim(usernewform)

      ' Loop through all of the items in the folder
       For I = 1 To Numitems
           Set CurItem = AllItems.Item(I)
           ' Test to see if the Message Class needs to be changed
           If CurItem.MessageClass <> NewForm Then
               CurItem.MessageClass = NewForm
               CurItem.Save
               ChangedItems = ChangedItems + 1
           End If
       Next

   End If
   MsgBox "Alterados " & ChangedItems & " items."

End Sub
