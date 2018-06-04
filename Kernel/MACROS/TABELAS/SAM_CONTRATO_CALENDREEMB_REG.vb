'HASH: 4A6653135B1AF7C7D0B4374045987142
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  ' SMS 25689 - Calendário de Reembolso por Regime de Atendimento
' Ultimas Alterações
'   - 01/06/2004 - Douglas - Implementação

  Dim qCalendarioAtual As Object
  Dim qOutrosCalendarios As Object

  Set qCalendarioAtual = NewQuery
  Set qOutrosCalendarios = NewQuery

  'Busca dados do calendário atual
  'Douglas - 01/06/2004
  qCalendarioAtual.Add("SELECT SC.HANDLE         CONTRATO, ")
  qCalendarioAtual.Add("       SCC.HANDLE        CALENDARIO, ")
  qCalendarioAtual.Add("       SCC.DATAINICIAL   DATAINICIAL, ")
  qCalendarioAtual.Add("       SCC.DATAFINAL     DATAFINAL, ")
  qCalendarioAtual.Add("       SCC.DIASPAGAMENTO DIASPAGAMENTO,")
  qCalendarioAtual.Add("       SCC.TIPOPEG       TIPOPEG")
  qCalendarioAtual.Add("  FROM SAM_CONTRATO_CALENDREEMB SCC")
  qCalendarioAtual.Add("  JOIN SAM_CONTRATO             SC  ON SC.HANDLE = SCC.CONTRATO ")
  qCalendarioAtual.Add(" WHERE scc.HANDLE = :CALEND ")
  qCalendarioAtual.ParamByName("CALEND").AsInteger = CurrentQuery.FieldByName("CONTRATOCALENDARIOREEMB").AsInteger
  qCalendarioAtual.Active = True


  'Verifica se o Regime faz parte de algum outro calendário em que a vigência coincide com o atual
  'e se o Regime já está incluído neste Calendário
  'Douglas - 01/06/2004
  qOutrosCalendarios.Add("SELECT SCC.HANDLE       HANDLE, ")
  qOutrosCalendarios.Add("       SCC.DESCRICAO    DESCRICAO,")
  qOutrosCalendarios.Add("       SCC.DATAINICIAL  DATAINICIAL,")
  qOutrosCalendarios.Add("       SCC.DATAFINAL    DATAFINAL ")
  qOutrosCalendarios.Add("  FROM SAM_CONTRATO_CALENDREEMB_REG SCCR")
  qOutrosCalendarios.Add("  JOIN SAM_CONTRATO_CALENDREEMB     SCC  ON SCC.HANDLE = SCCR.CONTRATOCALENDARIOREEMB")
  qOutrosCalendarios.Add(" WHERE SCCR.REGIMEATENDIMENTO = :REGIME ")
  qOutrosCalendarios.Add("   AND :DATAINICIAL <= DATAFINAL ")
  qOutrosCalendarios.Add("   AND SCC.CONTRATO = :CONTRATO ")

  If Not qCalendarioAtual.FieldByName("TIPOPEG").IsNull Then
    qOutrosCalendarios.Add("   AND SCC.TIPOPEG = :TIPOPEG ")
  End If

  
  qOutrosCalendarios.ParamByName("REGIME").AsInteger = CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger
  qOutrosCalendarios.ParamByName("DATAINICIAL").AsDateTime = qCalendarioAtual.FieldByName("DATAINICIAL").AsDateTime
  qOutrosCalendarios.ParamByName("CONTRATO").AsInteger = qCalendarioAtual.FieldByName("CONTRATO").AsInteger

  If Not qCalendarioAtual.FieldByName("TIPOPEG").IsNull Then
    qOutrosCalendarios.ParamByName("TIPOPEG").AsInteger = qCalendarioAtual.FieldByName("TIPOPEG").AsInteger
  End If


  qOutrosCalendarios.Active = True

  If Not qOutrosCalendarios.EOF Then
    If qOutrosCalendarios.FieldByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CONTRATOCALENDARIOREEMB").AsInteger Then
      bsShowMessage("Este Regime de Atendimento já foi inserido neste Calendário. Não é possível inserir repetido.", "E")
      CanContinue = False
      Exit Sub
    Else
      bsShowMessage("Este Regime de Atendimento já foi inserido em outro calendário (" + qOutrosCalendarios.FieldByName("DESCRICAO").AsString + ") com vigência concomitante a do calendário atual.", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

End Sub

