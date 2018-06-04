'HASH: 3DBDDB62DD68DA928551ECCAFA13A28E
'Macro: SAM_INCOMP_EVENTOS_PRESTADOR
'#Uses "*bsShowMessage"

'Alterada por: Soares - SMS: 60815 - 23/05/2006 - Início

Public Sub BOTAODUPLICAR_OnClick()
  Dim DuplicarIncomp As Object

  If CurrentQuery.State = 3 Then 'Verifica se a query está em edicao, se tiver nao permitir acesso ao botao.
    MsgBox("O registro não pode estar em edição.")
  Else
    Set DuplicarIncomp = CreateBennerObject("SAMDUPINCOMP.DUPLICAINCOMP")
    DuplicarIncomp.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 'Chama a interface de duplicacao de incompatibilidades
    Set DuplicarIncomp = Nothing
  End If
End Sub

Public Sub INCOMPATIBILIDADE_OnChange()
  TABLE_AfterScroll
End Sub


Public Sub INCOMPATIBILIDADE_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim vCampos As String
  Dim vColunas As String
  Dim ProcuraGrau As Long

  ShowPopup = False


  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_INCOMP_EVENTOS_GERAL.INCOMPATIBILIDADE"
  vCampos = "Incompatibilidade"

  'início sms 55308 - Edilson.Castro - 22/12/2005
  'ProcuraGrau =interface.Exec(CurrentSystem,"SAM_INCOMP_EVENTOS_GERAL",vColunas,2,vCampos,"SAM_INCOMP_EVENTOS_GERAL.TIPO = 'E' ","Incompatibilidades específicas",True,"")
  ProcuraGrau =Interface.Exec(CurrentSystem,"SAM_INCOMP_EVENTOS_GERAL",vColunas,1,vCampos,"DATAFINAL IS NULL OR DATAFINAL >= " + SQLDate(ServerDate)  ,"Incompatibilidades",True, "")
  'fim sms 55308

  Set Interface = Nothing
  If ProcuraGrau >0 Then 'selecionou alguma incompatibilidade específica
    CurrentQuery.Edit
    CurrentQuery.FieldByName("INCOMPATIBILIDADE").Value = ProcuraGrau
  End If
  '  INCOMPATIBILIDADE_OnChange
  TABLE_AfterScroll
End Sub

Public Sub MOTIVOGLOSAANTERIOR_OnChange()
  'SMS 50015 - 12.09.2005
  CurrentQuery.FieldByName("PERCENTGLOSAANTERIOR").AsFloat = 100
End Sub

Public Sub MOTIVOGLOSAPOSTERIOR_OnChange()
  'SMS 50015 - 12.09.2005
  CurrentQuery.FieldByName("PERCENTGLOSAPOSTERIOR").AsFloat = 100
End Sub

Public Sub TABLE_AfterScroll()
  Dim SQL As Object
  Dim SQLGrau As Object

  ShowPopup = False
  INCOMPATIBILIDADE.ReadOnly = False
  If (CurrentQuery.State = 1) Then TableReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
  If CurrentQuery.FieldByName("INCOMPATIBILIDADE").IsNull Then
    EVENTOANTERIOR.Text = ""
    EVENTOPOSTERIOR.Text = ""
    GRAUANTERIOR.Text = ""  'Soares - SMS: 60815 - 23/05/2006
    GRAUPOSTERIOR.Text = "" 'Soares - SMS: 60815 - 23/05/2006
    Exit Sub
  End If

  Set SQL = NewQuery
  '============================================================
  'EVENTO ANTERIOR
  SQL.Clear
  SQL.Add("SELECT	T.ESTRUTURA"+SQLConcatStr+" '-' "+SQLConcatStr+"T.DESCRICAO EVENTO")
  SQL.Add("  FROM	SAM_TGE T, ")
  SQL.Add("	    SAM_INCOMP_EVENTOS_GERAL G")
'início sms 55308 - Edilson.Castro - 22/12/2005
'SQL.Add(" WHERE G.TIPO = 'E' AND G.HANDLE = " +CurrentQuery.FieldByName("INCOMPATIBILIDADE").AsString)
SQL.Add(" WHERE G.HANDLE = " +CurrentQuery.FieldByName("INCOMPATIBILIDADE").AsString)
'fim sms 55308
  SQL.Add("	AND G.EVENTOANTERIOR = T.HANDLE")
  SQL.Active = True

  If Not SQL.EOF Then
    EVENTOANTERIOR.Text = SQL.FieldByName("EVENTO").AsString
  Else
    EVENTOANTERIOR.Text = ""
  End If
  '============================================================
  'EVENTO POSTERIOR
  SQL.Clear
  SQL.Add("SELECT	T.ESTRUTURA"+SQLConcatStr+" '-' "+SQLConcatStr+"T.DESCRICAO EVENTO")
  SQL.Add("  FROM	SAM_TGE T, ")
  SQL.Add("	    SAM_INCOMP_EVENTOS_GERAL G")
'início sms 55308 - Edilson.Castro - 22/12/2005
'SQL.Add(" WHERE G.TIPO = 'E' AND G.HANDLE = " +CurrentQuery.FieldByName("INCOMPATIBILIDADE").AsString)
SQL.Add(" WHERE G.HANDLE = " +CurrentQuery.FieldByName("INCOMPATIBILIDADE").AsString)
'fim sms 55308
  SQL.Add("   AND G.EVENTOPOSTERIOR = T.HANDLE")
  SQL.Active = True

  If Not SQL.EOF Then
    EVENTOPOSTERIOR.Text = SQL.FieldByName("EVENTO").AsString
  Else
    EVENTOPOSTERIOR.Text = ""
  End If

  Set SQL = Nothing

  'Soares - SMS: 60815 - 23/05/2006 - Início
  'Busca o grau anterior
  'Pesquisa para exibir os valores nos rotulos "grau anterior e grau posterior"
  Set SQLGrau = NewQuery

  SQLGrau.Clear
  SQLGrau.Add("SELECT GR.GRAU, GR.DESCRICAO         ")
  SQLGrau.Add("  FROM SAM_GRAU                 GR,  ")
  SQLGrau.Add("    	  SAM_INCOMP_EVENTOS_GERAL GE   ")
  SQLGrau.Add(" WHERE GE.HANDLE = " + CurrentQuery.FieldByName("INCOMPATIBILIDADE").AsString)
  SQLGrau.Add("   AND GE.GRAUANTERIOR = GR.HANDLE   ")
  SQLGrau.Active = True

  If Not SQLGrau.EOF Then
    GRAUANTERIOR.Text = SQLGrau.FieldByName("grau").AsString + " - " + SQLGrau.FieldByName("DESCRICAO").AsString
  Else
    GRAUANTERIOR.Text = ""
  End If

  'Busca o grau posterior

  SQLGrau.Clear
  SQLGrau.Add("SELECT GR.GRAU, GR.DESCRICAO         ")
  SQLGrau.Add("  FROM SAM_GRAU                 GR,  ")
  SQLGrau.Add("    	  SAM_INCOMP_EVENTOS_GERAL GE   ")
  SQLGrau.Add(" WHERE GE.HANDLE = " + CurrentQuery.FieldByName("INCOMPATIBILIDADE").AsString)
  SQLGrau.Add("   AND GE.GRAUPOSTERIOR = GR.HANDLE  ")
  SQLGrau.Active = True

  If Not SQLGrau.EOF Then
    GRAUPOSTERIOR.Text = SQLGrau.FieldByName("grau").AsString + " - " + SQLGrau.FieldByName("DESCRICAO").AsString
  Else
    GRAUPOSTERIOR.Text = ""
  End If

  Set SQLGrau = Nothing
  'Soares - SMS: 60815 - 23/05/2006 - Fim
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qVerificaSeJaExiste As Object
  Set qVerificaSeJaExiste = NewQuery
  qVerificaSeJaExiste.Add("SELECT COUNT(*) QTD FROM SAM_INCOMP_EVENTOS_PRESTADOR WHERE PRESTADOR = :PRESTADOR")
  qVerificaSeJaExiste.Add("   AND INCOMPATIBILIDADE = :INCOMPATIBILIDADE AND HANDLE <> :HANDLE")
  qVerificaSeJaExiste.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  qVerificaSeJaExiste.ParamByName("INCOMPATIBILIDADE").Value = CurrentQuery.FieldByName("INCOMPATIBILIDADE").AsInteger
  qVerificaSeJaExiste.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qVerificaSeJaExiste.Active = True

  If qVerificaSeJaExiste.FieldByName("QTD").AsInteger >0 Then
    bsShowMessage("Incompatibilidade já cadastrada para este prestador", "E")
    CanContinue = False
  End If

  If Not (VigenciaValida()) Then
	CanContinue = False
    bsShowMessage("Vigência inválida, data final não pode ser inferior a data inicial.", "E")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("TABTIPOACAO").AsInteger <>1 Then

    Set qVerificaSeJaExiste = Nothing

    CurrentQuery.FieldByName("MOTIVOGLOSAANTERIOR").Clear
    CurrentQuery.FieldByName("MOTIVOGLOSAPOSTERIOR").Clear
    CurrentQuery.FieldByName("MOTIVONEGACAO").Clear
    CurrentQuery.FieldByName("PERCENTGLOSAANTERIOR").Value = 0
    CurrentQuery.FieldByName("PERCENTGLOSAPOSTERIOR").Value = 0
  End If

  'Eduardo - 27/06/2006 - SMS 64160
  'Checagens duplicadas da tabela SAM_INCOMP_EVENTOS_GERAL
    If CurrentQuery.FieldByName("TABTIPOACAO").AsInteger = 1 Then
      If CurrentQuery.FieldByName("MOTIVOGLOSAANTERIOR").IsNull And CurrentQuery.FieldByName("MOTIVOGLOSAPOSTERIOR").IsNull Then
        bsShowMessage("Deve ser preenchido pelo menos um motivo de Glosa (Anterior/Posterior)", "E")
        CanContinue = False
        Exit Sub
      End If

      If(Not CurrentQuery.FieldByName("MOTIVOGLOSAANTERIOR").IsNull And CurrentQuery.FieldByName("PERCENTGLOSAANTERIOR").AsFloat = 0)Or _
         (Not CurrentQuery.FieldByName("MOTIVOGLOSAPOSTERIOR").IsNull And CurrentQuery.FieldByName("PERCENTGLOSAPOSTERIOR").AsFloat = 0)Then
        bsShowMessage("Quando selecionado um motivo de Glosa, deve-se preencher o campo % de Glosa !", "E")
        CanContinue = False
        Exit Sub
      End If

      If(CurrentQuery.FieldByName("PERCENTGLOSAANTERIOR").AsFloat <>0)And(CurrentQuery.FieldByName("MOTIVOGLOSAANTERIOR").IsNull)Then
        CanContinue = False
        bsShowMessage("Motivo de glosa do evento anterior é obrigatório quando percentual da glosa é diferente de zero", "E")
      End If

      If(CurrentQuery.FieldByName("PERCENTGLOSAPOSTERIOR").AsFloat <>0)And(CurrentQuery.FieldByName("MOTIVOGLOSAPOSTERIOR").IsNull)Then
        CanContinue = False
        bsShowMessage("Motivo de glosa do evento posterior é obrigatório quando percentual da glosa é diferente de zero", "E")
      End If
    End If
    If CurrentQuery.FieldByName("TABTIPOACAONEGACAO").AsInteger = 1 Then
      If CurrentQuery.FieldByName("MOTIVONEGACAOANTERIOR").IsNull And CurrentQuery.FieldByName("MOTIVONEGACAOPOSTERIOR").IsNull Then
        bsShowMessage("Deve ser preenchido pelo menos um motivo de negação (Anterior/Posterior)", "E")
        CanContinue = False
        Exit Sub
      End If
    End If
  'fim SMS 64160

  Dim validaCBOCBOS As Boolean
  Dim dllValidarCBOCBOS As Object
  Set dllValidarCBOCBOS = CreateBennerObject("Especifico.uEspecifico")
  validaCBOCBOS = dllValidarCBOCBOS.PRO_ValidaConsideraCBOCBOS(CurrentSystem)

  'SMS 167056 - Anderson Silva
  If ((validaCBOCBOS And (CurrentQuery.FieldByName("CONSIDERACBOCBOS").AsString = "M")) Or (CurrentQuery.FieldByName("CONSIDERACBOCBOS").AsString = "C")) Then
    If Not ((CurrentQuery.FieldByName("CONSIDERAEXECUTOR").AsString = "M") And (CurrentQuery.FieldByName("CONSIDERALOCALEXECUCAO").AsString = "M")) Then
	  bsShowMessage("CBO/CBOS – Parametrização permitida somente para “mesmos” Executores e Locais de Execução!", "I")
      CanContinue = False
      Exit Sub
    End If
  End If
  'SMS 167056 - Anderson Silva

	Set dllValidarCBOCBOS = Nothing
End Sub

Function VigenciaValida() As Boolean
	VigenciaValida = ((CurrentQuery.FieldByName("DATAFINAL").IsNull) Or (CurrentQuery.FieldByName("DATAFINAL").AsDateTime >= CurrentQuery.FieldByName("DATAINICIAL").AsDateTime))
End Function
