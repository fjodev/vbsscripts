Set objOL = Wscript.CreateObject("Outlook.Application")
Set olns = objOL.GetNamespace("MAPI")
Set olFolder = objOL.ActiveExplorer.CurrentFolder
CurFolder = olFolder ' get the name of the folder
Set AllItems = olFolder.Items ' get collection of iteems in folder
NumItems = AllItems.Count

' Continue = WshShell.Popup("Currently Selected Folder is: " & vbLF & CurFolder & vbLF & "No Items: " & NumItems & vbLF & "Do you want to continue?", , "Change Message Class",4)

' Ask for new message class i.e. IPM.Contact.MyContacts

txtNew = InputBox("Enter new message class", "New Message Class", "IPM.")

' This code was taken from the MS Word document that can do the same
' thing. But WSH is a much smaller overhead to do the same thing!

'
I = 0
If NumItems <> 0 And txtNew <> "" And Right(txtNew, 1) <> "." Then
  For Each Itm In AllItems
    I = I + 1
    lblCurrent = Itm.MessageClass
' This is where we should display the updated count
    If Itm.MessageClass <> txtNew Then
      Itm.MessageClass = txtNew
      Itm.Save
    End If
  Next
Wscript.Echo "Finished processing message classes."
Else
Wscript.Echo "Cannot process request."
End If

