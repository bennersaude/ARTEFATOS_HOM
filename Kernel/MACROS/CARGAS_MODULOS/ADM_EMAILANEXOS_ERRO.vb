'HASH: 18AD8FD111162E7FCFAF2A6116422904
Public Sub BOTAOLIBERARTODOS_OnClick() 
Dim Frm As CSVirtualForm 
Set Frm = NewVirtualForm 
Frm.TableName = "Z_VIRTUAL_REENVIOEMAIL" 
Frm.Caption = "Liberar e-mails com erros de envio 
If Frm.Show = 0 Then 
	RefreshNodesWithTable("Z_EMAILS") 
End If 
Set Frm = Nothing 
End Sub 
