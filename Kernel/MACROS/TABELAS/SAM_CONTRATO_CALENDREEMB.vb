'HASH: C7833A89D18CE77FFDC06AD05255F0AC

'SAM_CONTRATO_CALENDARIOREEMB
'Macro criada em 07/06/2004
'SMS 25689 - Douglas
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
 'sms 29652
  'Não permitir vigência cruzada com o mesmo tipo de PEG.
  'Não permitir vigência cruzada com tipo de PEG nulo.
  Dim Interface As Object
  Dim Resultado As String
  Dim Condicao As String

  'If CurrentQuery.FieldByName("TIPOPEG").IsNull Then
  '  Condicao = " AND TIPOPEG IS NULL"
  'Else
  '  Condicao = " AND TIPOPEG = " + CurrentQuery.FieldByName("TIPOPEG").AsString
  'End If

  'Condicao = ""

  'Set Interface =CreateBennerObject("SAMGERAL.Vigencia")

  'Resultado =Interface.Vigencia(CurrentSystem,"SAM_CONTRATO_CALENDREEMB","DATAINICIAL","DATAFINAL",CurrentQuery.FieldByName("DATAINICIAL").AsDateTime,CurrentQuery.FieldByName("DATAFINAL").AsDateTime,"CONTRATO",Condicao)

  'If Resultado ="" Then
  '  CanContinue =True
  'Else
  '  CanContinue =False
  '  MsgBox(Resultado)
  '  Set Interface =Nothing
  '  Exit Sub
  'End If
  'Set Interface =Nothing

'  Este código fazia com que não fosse permitido incluir dois calendários sem regime de atendimento para a mesma vigência.
  'Foi decidido deixar sem esse tratamento. Douglas/Larini

   Dim q1 As Object
   Dim q2 As Object
   Set q1 = NewQuery
   Set q2 = NewQuery

   'Seleciona os Calendário Do CONTRATO que possuam intersecção de vigência
   q1.Active = False
   q1.Add("SELECT SCC.HANDLE                   CALENDARIO")
   q1.Add("  FROM SAM_CONTRATO_CALENDREEMB SCC")
   q1.Add(" WHERE SCC.CONTRATO = :CONTRATO ")
   q1.Add("   AND ((:DATAINICIAL BETWEEN SCC.DATAINICIAL AND SCC.DATAFINAL)")
   q1.Add("    OR (:DATAFINAL BETWEEN SCC.DATAINICIAL AND SCC.DATAFINAL) ")
   q1.Add("    OR (:DATAINICIAL <= SCC.DATAINICIAL AND :DATAFINAL >= SCC.DATAFINAL)) ")
   q1.Add("   AND SCC.HANDLE <> :CALENDARIOATUAL")

   If Not CurrentQuery.FieldByName("TIPOPEG").IsNull Then
     q1.Add("    AND TIPOPEG = :TIPO")
   Else
     q1.Add("    AND TIPOPEG IS NULL")
   End If

   q1.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
   q1.ParamByName("DATAINICIAL").AsDateTime = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
   q1.ParamByName("DATAFINAL").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
   q1.ParamByName("CALENDARIOATUAL").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

   If Not CurrentQuery.FieldByName("TIPOPEG").IsNull Then
    q1.ParamByName("TIPO").AsInteger = CurrentQuery.FieldByName("TIPOPEG").AsInteger
   End If

   q1.Active = True

   'Verifica se na mesma vigência já existe algum calendário que não possua regimes cadastrados
   q2.Active = False
   q2.Add("SELECT HANDLE ")
   q2.Add("  FROM SAM_CONTRATO_CALENDREEMB_REG SCCR ")
   q2.Add(" WHERE SCCR.CONTRATOCALENDARIOREEMB = :CALENDARIO ")
   q2.Add("   AND SCCR.CONTRATOCALENDARIOREEMB <> :CALENDARIOATUAL")
   
   While Not q1.EOF
     q2.ParamByName("CALENDARIO").AsInteger = q1.FieldByName("CALENDARIO").AsInteger
     q2.ParamByName("CALENDARIOATUAL").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
     q2.Active = True
     If q2.EOF Then
       bsShowMessage("Não é possível incluir mais calendários nesta vigência." + Chr(13) + "Este Contrato já possui um calendário para todos os Regimes de Atendimento nesta vigência.", "E")
       q1.Active = False
       q2.Active = False
       Set q1 = Nothing
       Set q2 = Nothing
       CanContinue = False
       Exit Sub
     Else
       q2.Active = False
       q1.Next
     End If
   Wend

   q1.Active = False
   q2.Active = False
   Set q1 = Nothing
   Set q2 = Nothing

End Sub

