'HASH: 68ACE75633B38AF6C0CA12986D32B602
'Macro: SFN_PARAMETROSFIN
'#Uses "*PrimeiroDiaCompetencia"
'#Uses "*bsShowMessage"

'SMS 73362 - Débora Rebello - 08/12/2006
Public Sub CLASSETAXAADMINISTRACAOEMSUP_OnPopup(ShowPopup As Boolean)
  Dim dllProcura_Procurar As Object
  Dim viHandle As Long
  Dim vsCampos As String
  Dim vsColunas As String
  Dim vsCriterio As String


  ShowPopup = False
  Set dllProcura_Procurar = CreateBennerObject("Procura.Procurar")

  vsColunas = "ESTRUTURA|DESCRICAO"

  vsCriterio = ""

  vsCampos = "Estrutura|Descrição"

  viHandle = dllProcura_Procurar.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vsColunas, 1, vsCampos, vsCriterio, "Classes Gerenciais", False, "")

  If viHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLASSETAXAADMINISTRACAOEMSUP").Value = viHandle
  End If
  Set dllProcura_Procurar = Nothing
End Sub

'Public Sub CONTRIBUICAOCONTAFINANCEIRA_OnPopup(ShowPopup As Boolean)
'  Dim Interface As Object
'  Dim vHandle As Long
'  Dim vCampos As String
'  Dim vColunas As String
'  Dim vCriterio As String
'
'
'  ShowPopup = False
'  Set Interface = CreateBennerObject("Procura.Procurar")
'
'  vColunas = "SFN_PESSOA.NOME"
'
'  vCriterio = "SFN_CONTAFIN.TABRESPONSAVEL=3 AND  SFN_PESSOA.EHFISCO='S'"
'
'  vCampos = "Nome|Geração"
'
'  vHandle = Interface.Exec(CurrentSystem, "SFN_CONTAFIN|SFN_PESSOA[SFN_PESSOA.HANDLE=SFN_CONTAFIN.PESSOA]", vColunas, 1, vCampos, vCriterio, "Conta Financeira", True, CONTRIBUICAOCONTAFINANCEIRA.Text)
'
'  If vHandle <>0 Then
'    CurrentQuery.Edit
'    CurrentQuery.FieldByName("CONTRIBUICAOCONTAFINANCEIRA").Value = vHandle
'  End If
'  Set Interface = Nothing
'End Sub

'SMS 119240 - Ricardo Rocha - 14/08/2009
Public Sub CONTRIBUICAOPESSOA_OnPopup(ShowPopup As Boolean)
Dim Interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String


  ShowPopup = False
  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_PESSOA.CNPJCPF|SFN_PESSOA.NOME"

  vCriterio = "SFN_CONTAFIN.TABRESPONSAVEL=3 AND SFN_PESSOA.EHFISCO='S'"

  vCampos = "CNPJ/CPF|Nome"

  vHandle = Interface.Exec(CurrentSystem, "SFN_PESSOA|SFN_CONTAFIN[SFN_PESSOA.HANDLE=SFN_CONTAFIN.PESSOA]", vColunas, 2, vCampos, vCriterio, "Pessoa", True, CONTRIBUICAOPESSOA.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRIBUICAOPESSOA").Value = vHandle
  End If
  Set Interface = Nothing
End Sub

Public Sub INSSCONTAFINANCEIRA_OnPopup(ShowPopup As Boolean)
  INSSCONTAFINANCEIRA.LocalWhere = "EHFISCO='S'"
End Sub

Public Sub IRRFCONTAFINANCEIRA_OnPopup(ShowPopup As Boolean)
  IRRFCONTAFINANCEIRA.LocalWhere = "EHFISCO='S'"
End Sub


Public Sub TABLE_AfterScroll()
  INSSCONTAFINANCEIRA.WebLocalWhere = "A.EHFISCO='S'"
  IRRFCONTAFINANCEIRA.WebLocalWhere = "A.EHFISCO='S'"
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Query As Object

  If CurrentQuery.FieldByName("DIASPERMITIDOSATRASO").AsInteger > DateDiff("d", "01/01/1900", ServerDate) Then
	  bsShowMessage("A quantidade de dias permitidos de atraso deve ser menor", "E")
	  CanContinue = False
  End If

  If(CurrentQuery.FieldByName("UTILIZACENTROCUSTO").AsString = "S")And(CurrentQuery.FieldByName("CENTROCUSTOBAIXA").IsNull)And(CurrentQuery.FieldByName("TABFATURARECEBIMENTOMAIOR").AsInteger = 1)Then
  bsShowMessage("Deve-se informar o centro de custo para recebimento a maior", "E")
  CanContinue = False
  Exit Sub
End If

If(CurrentQuery.FieldByName("UTILIZACENTROCUSTO").AsString = "S")And((CurrentQuery.FieldByName("ISSCENTROCUSTO").IsNull)Or(CurrentQuery.FieldByName("IRRFCENTROCUSTO").IsNull)Or(CurrentQuery.FieldByName("CENTROCUSTOPADRAO").IsNull))Then
bsShowMessage("Falta informar um ou mais campos referentes ao centro de custo", "E")
CanContinue = False
Exit Sub
End If


If CurrentQuery.FieldByName("DEBDATAFINAL").AsDateTime < CurrentQuery.FieldByName("DEBDATAINICIAL").AsDateTime Then
  CanContinue = False
  TABLE.ActivePage(0)
  DEBDATAFINAL.SetFocus
  bsShowMessage("Data final de adequação de débito não pode ser menor que a data inicial!", "E")
  Exit Sub
End If


Set Query = NewQuery

Query.Clear
Query.Add("SELECT * FROM FIS_REGAUXCOMPETENCIA WHERE COMPETENCIA >= :PERIODOCONTINICIAL ORDER BY COMPETENCIA")
Query.ParamByName("PERIODOCONTINICIAL").AsDateTime = PrimeiroDiaCompetencia(CurrentQuery.FieldByName("PERIODOFATCONINICIAL").AsDateTime)
Query.Active = True

While(Not Query.EOF)
If Query.FieldByName("PROVISORIO").AsString = "N" Then
  bsShowMessage("Data contábil inicial não pode ser menor ou igual a competência do último registro auxiliar", "E")
  Set Query = Nothing
  CanContinue = False
  Exit Sub
End If

Query.Next
Wend

Set Query = Nothing

If CurrentQuery.FieldByName("LOCALTAXAADM").AsInteger = 1 Then 'Leonam - SMS 36033
  If CurrentQuery.FieldByName("TAXAADMINISTRACAOPRESTADOR").IsNull Then
    CanContinue = False
    TABLE.ActivePage(8)
    TAXAADMINISTRACAOPRESTADOR.SetFocus
    bsShowMessage("Classe gerencial 'Taxa de Administração' deve ser preenchida!", "E")
    Exit Sub
  End If
End If

If CurrentQuery.FieldByName("LOCALIMPOSTOS").AsString = "P" Then 'Henrique SMS: 21399
  If CurrentQuery.FieldByName("ISSCLASSEGERENCIAL").IsNull Then
    CanContinue = False
    TABLE.ActivePage(8)
    ISSCLASSEGERENCIAL.SetFocus
    bsShowMessage("Classe gerencial 'Retenção de ISS' deve ser preenchida !", "E")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("ISSREVISAOCLASSEGER").IsNull Then
    CanContinue = False
    TABLE.ActivePage(8)
    ISSREVISAOCLASSEGER.SetFocus
    bsShowMessage("Classe gerencial 'Revisão de ISS' deve ser preenchida !", "E")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("IRRFCLASSEGERENCIAL").IsNull Then
    CanContinue = False
    TABLE.ActivePage(8)
    IRRFCLASSEGERENCIAL.SetFocus
    bsShowMessage("Classe gerencial 'Retenção de IRRF' deve ser preenchida !", "E")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("INSSCLASSEGERENCIAL").IsNull Then
    CanContinue = False
    TABLE.ActivePage(8)
    INSSCLASSEGERENCIAL.SetFocus
    bsShowMessage("Classe gerencial 'Retenção de INSS' deve ser preenchida !", "E")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("INSSREVISAOCLASSEGER").IsNull Then
    CanContinue = False
    TABLE.ActivePage(8)
    INSSREVISAOCLASSEGER.SetFocus
    bsShowMessage("Classe gerencial 'Revisão de INSS' deve ser preenchida !", "E")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("CSLLCLASSEGERENCIAL").IsNull Then
    CanContinue = False
    TABLE.ActivePage(8)
    CSLLCLASSEGERENCIAL.SetFocus
    bsShowMessage("Classe gerencial 'Revisão CSLL' deve ser preenchida!", "E")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("COFINSCLASSEGERENCIAL").IsNull Then
    CanContinue = False
    TABLE.ActivePage(8)
    COFINSCLASSEGERENCIAL.SetFocus
    bsShowMessage("Classe gerencial 'Revisão COFINS' deve ser preenchida!", "E")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("PISPASEPCLASSEGERENCIAL").IsNull Then
    CanContinue = False
    TABLE.ActivePage(8)
    PISPASEPCLASSEGERENCIAL.SetFocus
    bsShowMessage("Classe gerencial 'Revisão PIS/PASEP' deve ser preenchida!", "E")
    Exit Sub
  End If

End If

If ((CurrentQuery.FieldByName("ENVIARTODASSEG").AsString = "S") And (CurrentQuery.FieldByName("VARIOSREGISTROSPROD").AsString <> "S")) Then
  CanContinue = False
  bsShowMessage("O campo 'Enviar todas as Segmentações' só pode ser marcado, caso o campo 'Vários registros por produto' esteja marcado!", "E")
  Exit Sub
End If

If ((CurrentQuery.FieldByName("PROVISIONARPEGSREEMBOLSO").AsString = "S") And (CurrentQuery.FieldByName("PROVISIONARPEGSCREDENCIAMENTO").AsString <> "S")) Then
  CanContinue = False
  bsShowMessage("O campo 'Provisionar PEG's de reembolso' só pode ser marcado, caso o campo 'Provisionar PEG's de credenciamento' esteja marcado!", "E")
  Exit Sub
End If


If ((CurrentQuery.FieldByName("PROVISIONARPEGSREEMBOLSO").AsString <> "S") Or (CurrentQuery.FieldByName("PROVISIONARPEGSCREDENCIAMENTO").AsString <> "S")) And (CurrentQuery.FieldByName("PROVISIONARRECUPERACAO").AsString = "S")Then
  CanContinue = False
  bsShowMessage("O campo 'Provisionar Recuperação - Contrato CO' só pode ser marcado, caso os campos 'Provisionar PEG's de credenciamento' e 'Provisionar PEG's de reembolso' estiverem marcados!", "E")
  Exit Sub
End If


If ((CurrentQuery.FieldByName("CONTROLADOTORC").AsInteger = 2) And ((CurrentQuery.FieldByName("PROVISIONARPEGSREEMBOLSO").AsString = "S") Or (CurrentQuery.FieldByName("PROVISIONARPEGSCREDENCIAMENTO").AsString = "S"))) Then
  CanContinue = False
  bsShowMessage("Não é possivel provisionar PEG's e controlar dotação orçamentária ao mesmo tempo!", "E")
  Exit Sub
End If


End Sub


Public Sub TAXAADMINISTRACAOPRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim dllProcura_Procurar As Object
  Dim viHandle As Long
  Dim vsCampos As String
  Dim vsColunas As String
  Dim vsCriterio As String


  ShowPopup = False
  Set dllProcura_Procurar = CreateBennerObject("Procura.Procurar")

  vsColunas = "ESTRUTURA|DESCRICAO"

  vsCriterio = ""

  vsCampos = "Estrutura|Descrição"

  viHandle = dllProcura_Procurar.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vsColunas, 1, vsCampos, vsCriterio, "Classes Gerenciais", False, "")

  If viHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("TAXAADMINISTRACAOPRESTADOR").Value = viHandle
  End If
  Set dllProcura_Procurar = Nothing
End Sub
