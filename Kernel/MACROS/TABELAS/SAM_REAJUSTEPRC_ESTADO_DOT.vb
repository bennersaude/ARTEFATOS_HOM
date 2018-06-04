'HASH: 817791DFD6300484FEB2A51AE6DCDA61

Option Explicit
'#Uses "*bsShowMessage"


Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vColunas As String
  Dim vCriterios As String
  Dim vHandle As Long
  Dim vCampos As String
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_TGE.ESTRUTURA|SAM_TGE.DESCRICAO"

  vCriterios = ""
  vCampos = "ESTRUTURA|Descrição"

  vHandle = interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 2, vCampos, vCriterios, "Tabela de Eventos", False, "")
  CurrentQuery.FieldByName("EVENTO").AsInteger = vHandle
  Set interface = Nothing
  ShowPopup = False

End Sub


'Public Sub TABLE_AfterInsert()
'  Dim TIPO As Object
'  Set TIPO =NewQuery
'  TIPO.Add("SELECT HANDLE FROM SAM_REAJUSTEPRC_PARAMTIPO T WHERE T.REAJUSTEPRCPARAM = :PARAM AND ")
'  TIPO.Add("T.TIPODOREAJUSTE = 'D'")
'  TIPO.ParamByName("PARAM").Value=CurrentQuery.FieldByName("REAJUSTEPRCPARAM").AsInteger
'  TIPO.Active=True
'  If TIPO.EOF Then
'SetParamTipo=False
'  Else
'    setParamTipo=True
'    CurrentQuery.FieldByName("PARAMTIPO").Value=TIPO.FieldByName("HANDLE").AsInteger
'  End If
'  TIPO.Active=False
'End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim qMunicipio As Object
  Set qMunicipio = NewQuery
  qMunicipio.Active = False
  qMunicipio.Add("SELECT ESTADO FROM SAM_REAJUSTEPRC_PARAM WHERE HANDLE = :HANDLE")
  qMunicipio.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_REAJUSTEPRC_PARAM")
  qMunicipio.Active = True
  If checkPermissao(CurrentSystem, CurrentUser, "E", qMunicipio.FieldByName("ESTADO").AsInteger, "E") = "N" Then

    bsShowMessage("Permissão negada! Usuário não pode incluir.", "E")

    CanContinue = False
    Set qMunicipio = Nothing
    Exit Sub
  End If
  Set qMunicipio = Nothing
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim qMunicipio As Object
  Set qMunicipio = NewQuery
  qMunicipio.Active = False
  qMunicipio.Add("SELECT ESTADO FROM SAM_REAJUSTEPRC_PARAM WHERE HANDLE = :HANDLE")
  qMunicipio.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_REAJUSTEPRC_PARAM")
  qMunicipio.Active = True
  If checkPermissao(CurrentSystem, CurrentUser, "E", qMunicipio.FieldByName("ESTADO").AsInteger, "A") = "N" Then

    bsShowMessage("Permissão negada! Usuário não pode incluir.", "E")

    CanContinue = False
    Set qMunicipio = Nothing
    Exit Sub
  End If
  Set qMunicipio = Nothing
End Sub

