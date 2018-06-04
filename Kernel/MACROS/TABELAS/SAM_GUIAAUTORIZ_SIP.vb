'HASH: D42413B57E5B57F255FA7590C7D441E6
'#Uses "*bsShowMessage"
'#Uses "*VerificaPermissaoEdicaoTriagem"
'#Uses "*RecordHandleOfTableInterfacePEG"

Public Sub ABORTOCID_OnEnter()
  ABORTOCID.AnyLevel = True
End Sub

Public Sub ABORTOCID_OnPopup(ShowPopup As Boolean)

  Dim handlexx As Long
  Dim ProcuraX As Object

  ShowPopup = False
  Set ProcuraX = CreateBennerObject("Procura.Procurar")
  handlexx = ProcuraX.Exec(CurrentSystem, "SAM_CID", "ESTRUTURA|Z_DESCRICAO", 2, "CID|Descrição", "HANDLE > 0", "Procura por CID", False, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("ABORTOCID").Value = handlexx
  End If
  Set ProcuraX = Nothing
End Sub

Public Sub ALTAMATERNACID_OnEnter()
  ALTAMATERNACID.AnyLevel = True
End Sub

Public Sub ALTAMATERNACID_OnPopup(ShowPopup As Boolean)

  Dim handlexx As Long
  Dim ProcuraX As Object

  ShowPopup = False
  Set ProcuraX = CreateBennerObject("Procura.Procurar")
  handlexx = ProcuraX.Exec(CurrentSystem, "SAM_CID", "ESTRUTURA|Z_DESCRICAO", 2, "CID|Descrição", "HANDLE > 0", "Procura por CID", False, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("ALTAMATERNACID").Value = handlexx
  End If
  Set ProcuraX = Nothing
End Sub

Public Sub NEONATALCID_OnEnter()
  NEONATALCID.AnyLevel = True
End Sub

Public Sub NEONATALCID_OnPopup(ShowPopup As Boolean)

  Dim handlexx As Long
  Dim ProcuraX As Object

  ShowPopup = False
  Set ProcuraX = CreateBennerObject("Procura.Procurar")
  handlexx = ProcuraX.Exec(CurrentSystem, "SAM_CID", "ESTRUTURA|Z_DESCRICAO", 2, "CID|Descrição", "HANDLE > 0", "Procura por CID", False, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("NEONATALCID").Value = handlexx
  End If
  Set ProcuraX = Nothing
End Sub

Public Sub NEONATALUTICTICID_OnEnter()
  NEONATALUTICTICID.AnyLevel = True
End Sub

Public Sub NEONATALUTICTICID_OnPopup(ShowPopup As Boolean)

  Dim handlexx As Long
  Dim ProcuraX As Object

  ShowPopup = False
  Set ProcuraX = CreateBennerObject("Procura.Procurar")
  handlexx = ProcuraX.Exec(CurrentSystem, "SAM_CID", "ESTRUTURA|Z_DESCRICAO", 2, "CID|Descrição", "HANDLE > 0", "Procura por CID", False, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("NEONATALUTICTICID").Value = handlexx
  End If
  Set ProcuraX = Nothing
End Sub

Public Sub PUERPERIOCID_OnEnter()
  PUERPERIOCID.AnyLevel = True
End Sub

Public Sub PUERPERIOCID_OnPopup(ShowPopup As Boolean)

  Dim handlexx As Long
  Dim ProcuraX As Object

  ShowPopup = False
  Set ProcuraX = CreateBennerObject("Procura.Procurar")
  handlexx = ProcuraX.Exec(CurrentSystem, "SAM_CID", "ESTRUTURA|Z_DESCRICAO", 2, "CID|Descrição", "HANDLE > 0", "Procura por CID", False, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PUERPERIOCID").Value = handlexx
  End If
  Set ProcuraX = Nothing
End Sub

Public Sub RNALTACID_OnEnter()
  RNALTACID.AnyLevel = True
End Sub

Public Sub RNALTACID_OnPopup(ShowPopup As Boolean)

  Dim handlexx As Long
  Dim ProcuraX As Object

  ShowPopup = False
  Set ProcuraX = CreateBennerObject("Procura.Procurar")
  handlexx = ProcuraX.Exec(CurrentSystem, "SAM_CID", "ESTRUTURA|Z_DESCRICAO", 2, "CID|Descrição", "HANDLE > 0", "Procura por CID", False, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("RNALTACID").Value = handlexx
  End If
  Set ProcuraX = Nothing

End Sub

Public Sub RNSALACID_OnEnter()
  RNSALACID.AnyLevel = True
End Sub

Public Sub RNSALACID_OnPopup(ShowPopup As Boolean)

  Dim handlexx As Long
  Dim ProcuraX As Object

  ShowPopup = False
  Set ProcuraX = CreateBennerObject("Procura.Procurar")
  handlexx = ProcuraX.Exec(CurrentSystem, "SAM_CID", "ESTRUTURA|Z_DESCRICAO", 2, "CID|Descrição", "HANDLE > 0", "Procura por CID", False, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("RNSALACID").Value = handlexx
  End If
  Set ProcuraX = Nothing

End Sub




Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	CanContinue = VerificarPermissaoUsuarioPegTriado(True)
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	CanContinue = VerificarPermissaoUsuarioPegTriado(True)
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	CanContinue = VerificarPermissaoUsuarioPegTriado(True)
	RecordReadOnly = Not CanContinue
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim q1 As Object
  Set q1 = NewQuery
  If CurrentQuery.FieldByName("GUIA").AsInteger >0 Then
    q1.Add("SELECT M.SEXO FROM SAM_BENEFICIARIO B, SAM_MATRICULA M, SAM_GUIA G WHERE M.HANDLE=B.MATRICULA AND B.HANDLE=G.BENEFICIARIO AND G.HANDLE=" + Str(RecordHandleOfTableInterfacePEG("SAM_GUIA")))
  ElseIf CurrentQuery.FieldByName("AUTORIZACAO").AsInteger >0 Then
    q1.Add("SELECT M.SEXO FROM SAM_BENEFICIARIO B, SAM_MATRICULA M, SAM_AUTORIZ A WHERE M.HANDLE=B.MATRICULA AND B.HANDLE=A.BENEFICIARIO AND A.HANDLE=" + Str(RecordHandleOfTableInterfacePEG("SAM_AUTORIZ")))
  End If
  q1.Active = True
  If q1.FieldByName("SEXO").AsString <>"F" Then

    bsShowMessage("Somente para sexo feminino", "E")
    CanContinue = False
  End If
  q1.Active = False
  Set q1 = Nothing

End Sub

Public Sub TRANSTORNOMATERNOCID_OnEnter()
  TRANSTORNOMATERNOCID.AnyLevel = True
End Sub

Public Sub TRANSTORNOMATERNOCID_OnPopup(ShowPopup As Boolean)

  Dim handlexx As Long
  Dim ProcuraX As Object

  ShowPopup = False
  Set ProcuraX = CreateBennerObject("Procura.Procurar")
  handlexx = ProcuraX.Exec(CurrentSystem, "SAM_CID", "ESTRUTURA|Z_DESCRICAO", 2, "CID|Descrição", "HANDLE > 0", "Procura por CID", False, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("TRANSTORNOMATERNOCID").Value = handlexx
  End If
  Set ProcuraX = Nothing
End Sub

