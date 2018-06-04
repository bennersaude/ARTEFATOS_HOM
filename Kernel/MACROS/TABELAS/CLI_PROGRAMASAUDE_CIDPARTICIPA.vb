'HASH: 45C087DBBA0E85CC42EBBAFA1AEB69B6
 


Public Sub CID_OnPopup(ShowPopup As Boolean)
  Dim dllProcura As Object
  Dim handlexx As Long
  ShowPopup = False
  Set dllProcura = CreateBennerObject("Procura.Procurar")
  handlexx = dllProcura.Exec(CurrentSystem, "SAM_CID", "ESTRUTURA|DESCRICAO", 2, "CID|Descrição", "HANDLE > 0", "Procura por CID", True, CID.Text)
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CID").Value = handlexx
  End If
  Set dllProcura = Nothing
End Sub
