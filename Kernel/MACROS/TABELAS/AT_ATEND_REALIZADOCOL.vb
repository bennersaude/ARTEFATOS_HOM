'HASH: 25F7448261D992DCDE6C18A16D469FC6
'#Uses "*bsShowMessagE"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|SAM_TGE.DESCRICAO|NIVELAUTORIZACAO"

  vCriterio = "ULTIMONIVEL = 'S' "
  vCampos = "Evento|Descrição|Nível"

  vHandle = interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 1, vCampos, vCriterio, "Tabela Geral de Eventos", False, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("evento").Value = vHandle
  End If
  Set interface = Nothing

End Sub

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTabela As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vTabela = "SAM_GRAU"
  vColunas = "SAM_GRAU.DESCRICAO|SAM_GRAU.XTHM|SAM_GRAU.VERIFICAGRAUSVALIDOS"
  vCriterio = "SAM_GRAU.VERIFICAGRAUSVALIDOS = 'N' " + _
              "OR (EXISTS (SELECT GE.HANDLE FROM SAM_TGE_GRAU GE WHERE GE.EVENTO=" + _
              CurrentQuery.FieldByName("EVENTO").AsString + _
              " AND GE.GRAU=SAM_GRAU.HANDLE))"
  vCampos = "Descrição|XTHM|Graus Válidos"

  vHandle = interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Graus válidos", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAU").Value = vHandle
  End If
  Set interface = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim myQuery As Object

  Set myQuery = NewQuery
  myQuery.Active = False
  myQuery.Add("SELECT DATAATENDIMENTO,HORAATENDIMENTO FROM AT_ATEND WHERE HANDLE=:pHANDLE")
  myQuery.ParamByName("pHANDLE").Value = CurrentQuery.FieldByName("Atendimento").Value
  myQuery.Active = True
  If(myQuery.FieldByName("DATAATENDIMENTO").IsNull)Or(myQuery.FieldByName("HORAATENDIMENTO").IsNull)Then
  bsShowMessage("Atendimento não foi realizado.Os campos Data e Hora do Atendimento deverão ser preenchidos!", "E")
  CanContinue = False
End If
Set myQuery = Nothing

End Sub

