'HASH: CA425AE85395BFBFC5C9B3EA39790712
'Macro: SFN_MODELO_ESTRUTURA_CAMPO

'Última alteração: 09/01/02
'Por: Milton
'Sub: CAMPOCONTABILIDADE_OnPopup
'SMS: 5782

'#Uses "*bsShowMessage
'#Uses "*IsInt

Public Sub BOTAOCAMPOS_OnClick()
  Dim interface As Object
  Set interface = CreateBennerObject("Financeiro.Campos")
  interface.Exec(CurrentSystem, False, 4)
  Set interface = Nothing
End Sub

Public Sub CAMPOCONTABILIDADE_OnPopup(ShowPopup As Boolean)

  CAMPOCONTABILIDADE.LocalWhere = "SIS_CONTABCAMPOS.ORIGEM='9'"

End Sub

Public Sub COLFINAL_OnExit()
  If CurrentQuery.FieldByName("DESCRICAO").IsNull Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("DESCRICAO").Value = _
                             ZERAR(CurrentQuery.FieldByName("COLINICIAL").AsString) + "." + _
                             ZERAR(CurrentQuery.FieldByName("COLFINAL").AsString) + "- "
  End If
End Sub

Public Function ZERAR(NUM As String)As String
  ZERAR = Right("000" + NUM, 3)
End Function

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  ' sms 40765 Lopes
  Dim Qtemp As BPesquisa
  Set Qtemp = NewQuery
  Qtemp.Active = False
  Qtemp.Clear
  Qtemp.Add("Select NOME")
  Qtemp.Add("  FROM SIS_CONTABCAMPOS")
  Qtemp.Add(" WHERE HANDLE = :CAMPO")
  Qtemp.ParamByName("CAMPO").AsInteger = CurrentQuery.FieldByName("CAMPO").AsInteger
  Qtemp.Active = True
  If Qtemp.FieldByName("NOME").AsString = "BAIXAJURO" Or Qtemp.FieldByName("NOME").AsString = "BAIXAMULTA" Then
    Qtemp.Active = False
    Qtemp.Clear
    Qtemp.Add("Select COUNT(HANDLE) QT ")
    Qtemp.Add("  FROM SFN_MODELO_ESTRUTURA_CAMPO")
    Qtemp.Add(" WHERE MODELOESTRUTURA = :MOD")
    Qtemp.Add("   And CAMPO In (Select HANDLE FROM SIS_CONTABCAMPOS WHERE  NOME = 'BAIXAJUROSMULTA') ")
    Qtemp.ParamByName("MOD").AsInteger = CurrentQuery.FieldByName("MODELOESTRUTURA").AsInteger
    Qtemp.Active = True
    If Qtemp.FieldByName("QT").AsInteger > 0 Then
      bsShowMessage("O modelo não pode possuir o campo 'Baixa juros e multa' juntamente com o campo 'Baixa juro' ou com o campo 'Baixa multa'", "I")
      CanContinue = False
    End If
  Else
    If Qtemp.FieldByName("NOME").AsString = "BAIXAJUROSMULTA" Then
      Qtemp.Active = False
      Qtemp.Clear
      Qtemp.Add("Select COUNT(HANDLE) QT ")
      Qtemp.Add("  FROM SFN_MODELO_ESTRUTURA_CAMPO ")
      Qtemp.Add(" WHERE MODELOESTRUTURA = :MOD")
      Qtemp.Add("   And CAMPO In (Select HANDLE FROM SIS_CONTABCAMPOS WHERE  NOME = 'BAIXAJURO' OR NOME = 'BAIXAMULTA')")
      Qtemp.ParamByName("MOD").AsInteger = CurrentQuery.FieldByName("MODELOESTRUTURA").AsInteger
      Qtemp.Active = True
      If Qtemp.FieldByName("QT").AsInteger > 0 Then
      	bsShowMessage("O modelo não pode possuir o campo 'Baixa juros e multa' juntamente com o campo 'Baixa juro' ou com o campo 'Baixa multa'", "I")
        CanContinue = False
      End If
    End If
  End If

  Qtemp.Active = False
  Qtemp.Clear
  Qtemp.Add("Select COUNT(HANDLE) QT ")
  Qtemp.Add("  FROM SFN_MODELO_ESTRUTURA")
  Qtemp.Add(" WHERE HANDLE = :MOD AND TABTIPO = '17'")
  Qtemp.ParamByName("MOD").AsInteger = CurrentQuery.FieldByName("MODELOESTRUTURA").AsInteger
  Qtemp.Active = True

  If (Qtemp.FieldByName("QT").AsInteger > 0) And Not IsInt(CurrentQuery.FieldByName("ORDEMCAMPO").AsString) Then
    bsShowMessage("O modelo permite somente valores numéricos para o campo 'Ordem'!", "I")
    CanContinue = False
  End If

  Qtemp.Active = False
  Set Qtemp = Nothing

  Dim qLeiauteCartao As BPesquisa
  Set qLeiauteCartao = NewQuery

  qLeiauteCartao.Active = False
  qLeiauteCartao.Clear
  qLeiauteCartao.Add("SELECT M.HANDLE                              ")
  qLeiauteCartao.Add("  FROM SFN_MODELO_ESTRUTURA ME               ")
  qLeiauteCartao.Add("  JOIN SFN_MODELO M ON (M.HANDLE = ME.MODELO)")
  qLeiauteCartao.Add(" WHERE M.TABTIPO = 8                         ")
  qLeiauteCartao.Add("   AND ME.HANDLE = :PHANDLECAMPO             ")
  qLeiauteCartao.ParamByName("PHANDLECAMPO").AsInteger = CurrentQuery.FieldByName("MODELOESTRUTURA").AsInteger
  qLeiauteCartao.Active = True

  If (Not qLeiauteCartao.FieldByName("HANDLE").IsNull) Then
  	If(CurrentQuery.FieldByName("ORDEMCAMPO").IsNull) Then
		BsShowMessage("Quando for Leiaute de Cartão é necessário preencher o campo Ordem.", "E")
		CanContinue = False
  	End If
  End If

  Set qLeiauteCartao = Nothing

End Sub

Public Sub TABLE_NewRecord()
  Select Case NodeInternalCode
    Case 1
      CurrentQuery.FieldByName("TIPOREGISTRO").Value = "C"
    Case 2
      CurrentQuery.FieldByName("TIPOREGISTRO").Value = "D"
    Case 3
      CurrentQuery.FieldByName("TIPOREGISTRO").Value = "R"
  End Select
End Sub

Public Sub TABLE_AfterPost()
	Dim atualizaModelo As BPesquisa
    Set atualizaModelo = NewQuery

    atualizaModelo.Clear
    atualizaModelo.Add(" UPDATE SFN_MODELO                              ")
    atualizaModelo.Add("    SET USUARIOALTERACAO  = :HUSUARIOALTERACAO, ")
	atualizaModelo.Add("	    DATAHORAALTERACAO = :DATAHORAALTERACAO  ")
    atualizaModelo.Add("  WHERE HANDLE = (SELECT MODELO FROM SFN_MODELO_ESTRUTURA WHERE HANDLE = :HMODELOESTRUTURA) ")
    atualizaModelo.ParamByName("HUSUARIOALTERACAO").AsInteger = CurrentSystem.CurrentUser
    atualizaModelo.ParamByName("DATAHORAALTERACAO").AsDateTime = CurrentSystem.ServerNow
    atualizaModelo.ParamByName("HMODELOESTRUTURA").AsInteger = CurrentQuery.FieldByName("MODELOESTRUTURA").AsInteger
    atualizaModelo.ExecSQL

    Set atualizaModelo = Nothing
End Sub

