'HASH: 2EE236632561C33A7985F8553F7B6684
'Macro: SAM_ACOMODACAO_EVENTO
'#Uses "*bsShowMessage"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim qParamAtend As Object

  Set qParamAtend = NewQuery

  ShowPopup =False
  Set interface =CreateBennerObject("Procura.Procurar")

  vColunas ="ESTRUTURA|SAM_TGE.DESCRICAO|NIVELAUTORIZACAO"

  vCriterio ="ULTIMONIVEL = 'S' "
  vCampos ="Evento|Descrição|Nível"

  vHandle =interface.Exec(CurrentSystem,"SAM_TGE",vColunas,1,vCampos,vCriterio,"Tabela Geral de Eventos",True,"")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value =vHandle
  End If
  Set interface =Nothing

End Sub

Public Sub GRAUAGERAR_OnPopup(ShowPopup As Boolean)

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup =False
  Set interface =CreateBennerObject("Procura.Procurar")

  vColunas ="GRAU|DESCRICAO|SAM_GRAU.VERIFICAGRAUSVALIDOS"

  vCriterio ="TIPOGRAU IN (SELECT HANDLE FROM SAM_TIPOGRAU WHERE CLASSIFICACAO = '3') AND HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = " +CurrentQuery.FieldByName("EVENTO").AsString +" )"
  vCampos ="Grau|Descrição|Graus Válidos"
  vHandle =interface.Exec(CurrentSystem,"SAM_GRAU",vColunas,1,vCampos,vCriterio,"Graus válidos para o evento, classificados como diária",True,"")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAUAGERAR").Value =vHandle
  End If
  Set interface =Nothing

End Sub

Public Sub TABLE_AfterScroll()

  If (WebMode) Then
    EVENTO.WebLocalWhere = "A.ULTIMONIVEL = 'S'"

    GRAUAGERAR.WebLocalWhere = "A.TIPOGRAU IN (SELECT HANDLE FROM SAM_TIPOGRAU WHERE CLASSIFICACAO = '3') " + _
                             "AND A.HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = @CAMPO(EVENTO) )"
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim sql As Object
  Set sql =NewQuery
  If Len(UserVar("CAB_ACOMODACAO_EVENTO_DUPLIDADO")) = 0 Then
    sql.Add("SELECT A.DESCRICAO FROM SAM_ACOMODACAO_EVENTO AE,SAM_ACOMODACAO A  WHERE A.HANDLE=AE.ACOMODACAO AND AE.EVENTO=" +CurrentQuery.FieldByName("EVENTO").AsString)
    If CurrentQuery.State =2 Then
      sql.Add("AND AE.HANDLE <> :HANDLE")
      sql.ParamByName("HANDLE").Value =CurrentQuery.FieldByName("HANDLE").AsInteger
    End If
    sql.Active =True
    If Not sql.EOF Then
      bsShowMessage("Este evento já faz parte de outra acomodação: " +sql.FieldByName("DESCRICAO").AsString, "E")
      CanContinue =False
    End If
  End If

  sql.Active = False
  sql.Clear
  sql.Add("SELECT G.HANDLE ")
  sql.Add("  FROM SAM_GRAU G")
  sql.Add("  JOIN SAM_TIPOGRAU TG ON (G.TIPOGRAU = TG.HANDLE) ")
  sql.Add("  JOIN SAM_TGE_GRAU TGE ON (TGE.GRAU = G.HANDLE) ")
  sql.Add(" WHERE TG.CLASSIFICACAO = '3'  ")
  sql.Add("   AND TGE.EVENTO = :EVENTO ")
  sql.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  sql.Active = True

  If (sql.EOF) Then
    bsShowMessage("Grau Inválido para o Evento! ", "E")
    CanContinue =False
  End If

  'Douglas - 28/01/2005
  'SMS 38603 - Mostrar a acomodação do modulo do beneficiario na interface de autorização
  'Se o flag EVENTOPADRAOACOMODACAO for marcado, verificar se algum mais estiver com o flag marcado e desmarcar
  If CurrentQuery.FieldByName("EVENTOPADRAOACOMODACAO").AsString = "S" Then
    sql.Active = False
    sql.Clear
    sql.Add("UPDATE SAM_ACOMODACAO_EVENTO ")
    sql.Add("   SET EVENTOPADRAOACOMODACAO = 'N' ")
    sql.Add(" WHERE EVENTOPADRAOACOMODACAO = 'S' ")
    sql.Add("   AND HANDLE <> :CURRENTQUERY ")
    sql.Add("   AND ACOMODACAO = :ACOMODACAO ")
    sql.ParamByName("CURRENTQUERY").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    sql.ParamByName("ACOMODACAO").AsInteger = RecordHandleOfTable("SAM_ACOMODACAO")
    sql.ExecSQL
  End If

  Set sql =Nothing

End Sub
