Sub ChangeForm()
   Dim ChangedItems As Integer
   ChangedItems = 0
   
   ' Dados da pasta
   Set Curfolder = Application.ActiveExplorer.CurrentFolder
   Set AllItems = Curfolder.Items
   Numitems = Curfolder.Items.Count

   
   ' Confirma ac��o
   resposta = MsgBox("Ser�o alterados " & Numitems & " itens na pasta " _
    & Curfolder & ". Deseja continuar?", _
    vbYesNo, "Mudar para " & NewForm)
   If resposta = vbYes Then
       ' Defini��o do novo formul�rio
       usernewform = InputBox("Novo formul�rio", "Altera��o de formul�rio")
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
