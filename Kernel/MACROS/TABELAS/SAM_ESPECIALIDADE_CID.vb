'HASH: E764FB8CCF54DA916CDE99DA44D14AF7
 

Public Sub BOTAOEXCLUICID_OnClick()
  Dim Obj As Object

  Set Obj =CreateBennerObject("SamGerarCID.GerarCID")
  Obj.Excluir(CurrentSystem,CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger)

  Set Obj =Nothing
End Sub

Public Sub BOTAOINCLUICID_OnClick()
  Dim Obj As Object

  Set Obj =CreateBennerObject("SamGerarCID.GerarCID")
  Obj.Cadastrar(CurrentSystem,CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger)

  Set Obj =Nothing
End Sub

Public Sub CID_OnPopup(ShowPopup As Boolean)
  CID.AnyLevel = True
End Sub
