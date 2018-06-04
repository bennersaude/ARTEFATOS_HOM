'HASH: 9DE653E69E266CD0D2580D47E85A9FAF
'Macro: SAM_PLANO_MOD
'#Uses "*bsShowMessage"
'Daniela -SMS 12220 -Convênio no registro da ANS

Public Sub REGISTROMS_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set Interface = CreateBennerObject("Procura.Procurar")

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT CONVENIO")
  SQL.Add("FROM SAM_PLANO")
  SQL.Add("WHERE HANDLE = :HPLANO")
  SQL.ParamByName("HPLANO").Value = CurrentQuery.FieldByName("PLANO").AsInteger
  SQL.Active = True

  vCriterio = "CONVENIO = " + SQL.FieldByName("CONVENIO").AsString

  Set SQL = Nothing

  vColunas = "SAM_REGISTROMS.REGISTROMS|SAM_REGISTROMS.DESCRICAO|SAM_SEGMENTACAO.DESCRICAO|SAM_REGISTROMS.DATAVENCIMENTO"

  vCampos = "Registro|Descrição|Segmentação|Vencimento"

  vHandle = Interface.Exec(CurrentSystem, "SAM_REGISTROMS|SAM_SEGMENTACAO[SAM_REGISTROMS.SEGMENTACAO = SAM_SEGMENTACAO.HANDLE]", vColunas, 1, vCampos, vCriterio, "Registro no Ministério da Saúde", True, "")

  If vHandle <>0 Then
    CurrentQuery.FieldByName("REGISTROMS").Value = vHandle
  End If

  Set Interface = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  'Início Daniela Zardo -10/09/2002
  If(CurrentQuery.FieldByName("ACEITATITULAR").AsString = "N") _
     And(CurrentQuery.FieldByName("ACEITADEPENDENTES").AsString = "N") _
     And(CurrentQuery.FieldByName("ACEITAAGREGADOS").AsString = "N")Then
  bsShowMessage("Um dos campos Aceita Titular, Dependente ou Agregado devem estar ativos!", "E")
  CanContinue = False
End If

If CurrentQuery.FieldByName("VERIFICADEPENDENTEAGREGADO").AsString = "S" Then
  If(CurrentQuery.FieldByName("ACEITATITULAR").AsString = "N")Then
  bsShowMessage("Para verificar dependentes agregados, o campo Aceita Titular deve estar ativo!", "E")
  CanContinue = False
End If
If(CurrentQuery.FieldByName("ACEITADEPENDENTES").AsString = "S")Then
	bsShowMessage("Para verificar dependentes agregados, o campo Aceita Dependente deve estar inativo!", "E")
	CanContinue = False
End If
	If(CurrentQuery.FieldByName("ACEITAAGREGADOS").AsString = "S")Then
	bsShowMessage("Para verificar dependentes agregados, o campo Aceita Agregados deve estar inativo!", "E")
CanContinue = False
End If
	If(CurrentQuery.FieldByName("AUTOMATICO").AsString = "S")Then
	bsShowMessage("Para verificar dependentes agregados, o campo Automático deve estar inativo!", "E")
CanContinue = False
End If
	If(CurrentQuery.FieldByName("PROPAGAR").AsString = "S")Then
	bsShowMessage("Para verificar dependentes agregados, o campo Propagar deve estar inativo!", "E")
CanContinue = False
End If
End If
'Fim alteração Daniela

If(CurrentQuery.FieldByName("OBRIGATORIO").AsString = "S")And _
   (CurrentQuery.FieldByName("REGISTROMS").IsNull)Then
	bsShowMessage("O registro no ministério de saúde é obrigatório para módulos obrigatórios", "E")
	CanContinue = False
End If

Dim SQL As Object

Set SQL = NewQuery

If Not CurrentQuery.FieldByName("REGISTROMS").IsNull Then
  SQL.Clear
  SQL.Add("SELECT DATAVENCIMENTO")
  SQL.Add("FROM SAM_REGISTROMS")
  SQL.Add("WHERE HANDLE = :HREGISTROMS")
  SQL.ParamByName("HREGISTROMS").Value = CurrentQuery.FieldByName("REGISTROMS").AsInteger
  SQL.Active = True

  If Not SQL.FieldByName("DATAVENCIMENTO").IsNull And _
                         (SQL.FieldByName("DATAVENCIMENTO").AsDateTime <ServerDate)Then
    bsShowMessage("O Registro no Ministério da Saúde está vencido! Verifique", "E")
  End If
End If


If CurrentQuery.FieldByName("SEGUNDAPARCELA").AsString = "2" And _
                            CurrentQuery.FieldByName("PRIMEIRAPARCELA").AsString <>"2" Then
  CanContinue = False
  bsShowMessage("Para segunda parcela 'Proporcional' a primeira parcela deve ser integral", "E")
End If

'sms 49923
'Permitir a alocação de dois módulos iguais contanto que,pelo menos,uma das abrangências seja diferente.
If(CanContinue = True)Then
SQL.Active = False
SQL.Clear

SQL.Add("SELECT COUNT(*) NUMREGISTROS      ")
SQL.Add("    FROM SAM_PLANO_MOD            ")
SQL.Add("  WHERE PLANO = :PPLANO           ")
SQL.Add("        AND MODULO = :PMODULO     ")
SQL.Add("        AND HANDLE <> :PHCORRENTE ")

SQL.ParamByName("PPLANO").Value = CurrentQuery.FieldByName("PLANO").Value
SQL.ParamByName("PMODULO").Value = CurrentQuery.FieldByName("MODULO").Value
SQL.ParamByName("PHCORRENTE").Value = CurrentQuery.FieldByName("HANDLE").Value

'Considerar o Registro no Ministério da Saúde na verificação de duplicidade
If(CurrentQuery.FieldByName("REGISTROMS").IsNull)Then
SQL.Add("        AND REGISTROMS IS NULL")
Else
  SQL.Add("        AND REGISTROMS = :PREGISTROMS")
  SQL.ParamByName("PREGISTROMS").Value = CurrentQuery.FieldByName("REGISTROMS").AsInteger
End If


SQL.Active = True

If(SQL.FieldByName("NUMREGISTROS").Value >0)Then
	bsShowMessage("Não é permitido alocar módulos iguais no plano.", "E")
	CanContinue = False
	SQL.Active = False
	Set SQL = Nothing
	Exit Sub
End If
End If
Set SQL = Nothing
End Sub

