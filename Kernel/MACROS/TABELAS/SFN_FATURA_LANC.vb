'HASH: 800F8C9A663A7188A120ABF18B7E19CC
'Macro: SFN_FATURA_LANC

'#Uses "*bsShowMessage
'#uses "*PermissaoAlteracao"
'Última alteração:
' Milton - 12/03/2002
' (Descrições dos erros na contabilização)
'Henrique 31/12/2002 - colocado a cláusula "FROM" no comando "DELETE"   --  não funcionava em DB2


Public Sub BOTAOCONTABILIZA_OnClick()
  Dim erro As Long
  Dim sql As Object
  Dim OLE As Object
  Dim DescErro As String
  Dim vOperacaoRetorno As String
  Dim vsMensagem As String

  If bsShowMessage("Confirma a contabilização ?", "Q") = vbYes Then

  Set sql = NewQuery

  sql.Clear
  sql.Active = False
  sql.Add("SELECT CONTABILIZA FROM SFN_PARAMETROSFIN")
  sql.Active = True

  If sql.FieldByName("CONTABILIZA").AsString <> "S" Then
    bsShowMessage("Nos Parâmetros Gerais - Financeiro, o campo 'Contabiliza' não está marcado !", "I")
    Exit Sub
  End If

  Set OLE = CreateBennerObject("Financeiro.Geral")

  If Not InTransaction Then StartTransaction

  sql.Clear
  sql.Add("DELETE FROM SFN_CONTAB_LANC_DEBCRE ")
  sql.Add("WHERE HANDLE IN (SELECT E.HANDLE FROM SFN_CONTAB_LANC_DEBCRE E, SFN_CONTAB_LANC D ")
  sql.Add("                 WHERE E.CONTABLANC = D.HANDLE And D.FATURALANC = " + CurrentQuery.FieldByName("HANDLE").AsString + ")")
  sql.ExecSQL

  sql.Clear
  sql.Add("DELETE FROM SFN_CONTAB_LANC WHERE FATURALANC = " + CurrentQuery.FieldByName("HANDLE").AsString)
  sql.ExecSQL

  If InTransaction Then Commit

  sql.Clear
  sql.Add("SELECT P.HANDLE                                          ")
  sql.Add("  FROM SAM_PEG         P                                 ")
  sql.Add("  LEFT JOIN SFN_FATURA      F  ON (P.PEG = F.PEG)        ")
  sql.Add("  LEFT JOIN SFN_FATURA_LANC FL ON (F.HANDLE = FL.FATURA) ")
  sql.Add(" WHERE FL.HANDLE = :HND                                  ")
  sql.ParamByName("HND").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.Active = True

  erro = OLE.ContabilizaTreeView(CurrentSystem, _
                                 CurrentQuery.FieldByName("CLASSEGERENCIAL").AsInteger, _
                                 CurrentQuery.FieldByName("OPERACAO").AsInteger, _
                                 CurrentQuery.FieldByName("DATACONTABIL").AsDateTime, _
                                 CurrentQuery.FieldByName("VALOR").AsFloat, _
                                 CurrentQuery.FieldByName("NATUREZA").AsString, _
                                 1, _
                                 0, _
                                 CurrentQuery.FieldByName("FATURA").AsInteger, _
                                 CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                 0, _
                                 0, _
                                 CurrentQuery.FieldByName("TESOURARIA").AsInteger, _
                                 0, _
                                 0, _
                                 PermissaoAlteracao(sql.FieldByName("HANDLE").AsInteger, 0, 0, True, vsMensagem), _
                                 CurrentQuery.FieldByName("TIPOLANCFIN").AsInteger, _
                                 vOperacaoRetorno)

  If erro > 0 Then
    bsShowMessage("Contabilização efetuada com sucesso!", "I")
  Else
    Select Case erro
      Case -10
        DescErro = "Sem Regra Financeira"
      Case -11
        DescErro = "Código de Operacao de Lançamento Inválido"
      Case -12
        DescErro = "Sem Regra de Contabilização para Classe Gerencial"
      Case -13
        DescErro = "Sem Classe Contábil - Débito"
      Case -14
        DescErro = "Sem Classe Contábil - Crédito"
      Case -15
        DescErro = "Classe Contábil da Conta Financeira Inválida"
      Case -16
        DescErro = "Classe Contábil - Tesouraria Inválida"
      Case -17
        DescErro = "Classe Gerencial Inválida"
      End

  End Select

  If DescErro = "" Then
    DescErro = "Erro desconhecido !"
  End If

  bsShowMessage("Erro no processo de contabilização: " + DescErro, "I")
End If

CurrentQuery.Active = False
CurrentQuery.Active = True


Set OLE = Nothing
Set sql = Nothing

End If


End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCONTABILIZA"
			BOTAOCONTABILIZA_OnClick
	End Select
End Sub
