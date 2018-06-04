'HASH: 5CB21970244B27E58C5B6179601AADEF
 

Public Sub UTILIZAREMPENHO_OnPopup(ShowPopup As Boolean)

  ShowPopup = False

  Dim vHandle As Long

  vCriterio = ""

  vHandle = FiltrarEmpenho(True, UTILIZAREMPENHO.Text)

  If vHandle <> 0 Then
    	CurrentQuery.FieldByName("UTILIZAREMPENHO").AsInteger = vHandle
  End If

End Sub

Public Sub UTILIZAREMPENHOENCARGO_OnPopup(ShowPopup As Boolean)

  ShowPopup = False

  Dim vHandle As Long

  vCriterio = ""

  vHandle = FiltrarEmpenho(True, UTILIZAREMPENHOENCARGO.Text)

  If vHandle <> 0 Then
    	CurrentQuery.FieldByName("UTILIZAREMPENHOENCARGO").AsInteger = vHandle
  End If

End Sub

Public Sub VERIFICAREMPENHO_OnPopup(ShowPopup As Boolean)

  ShowPopup = False

  Dim vHandle As Long

  vCriterio = ""

  vHandle = FiltrarEmpenho(False, VERIFICAREMPENHO.Text)

  If vHandle <> 0 Then
    	CurrentQuery.FieldByName("VERIFICAREMPENHO").AsInteger = vHandle
  End If

End Sub

Public Sub VERIFICAREMPENHOENCARGO_OnPopup(ShowPopup As Boolean)

  ShowPopup = False

  Dim vHandle As Long

  vCriterio = ""

  vHandle = FiltrarEmpenho(False, VERIFICAREMPENHOENCARGO.Text)

  If vHandle <> 0 Then
    	CurrentQuery.FieldByName("VERIFICAREMPENHOENCARGO").AsInteger = vHandle
  End If

End Sub


Public Function FiltrarEmpenho(filtrarExercicio As Boolean, texto As String) As Long

  Dim INTERFACE As Object
  Dim vColunas, vCriterio, vCampos, vTabela As String
  Set INTERFACE = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_DOTACAOEXERCICIO.EXERCICIO|SFN_EMPENHO.NUMERO|SFN_EMPENHO.DESCRICAO|NATUREZADESPESA.DESCRICAO|UGRESPONSAVEL.DESCRICAO"
  vCampos = "Exercício|Número do Empenho|Descrição do Empenho|Natureza|Dotação"
  vTabela = "SFN_EMPENHO|SFN_DOTACAONATUREZA[SFN_EMPENHO.DOTACAONATUREZA = SFN_DOTACAONATUREZA.HANDLE]|"
  vTabela = vTabela + "NATUREZADESPESA[SFN_DOTACAONATUREZA.NATUREZADESPESA=NATUREZADESPESA.HANDLE]|"
  vTabela = vTabela + "SFN_DOTACAO[SFN_DOTACAO.HANDLE = SFN_DOTACAONATUREZA.DOTACAO]|"
  vTabela = vTabela + "UGRESPONSAVEL[UGRESPONSAVEL.HANDLE = SFN_DOTACAO.UGRESPONSAVEL]|"
  vTabela = vTabela + "SFN_DOTACAOEXERCICIO[SFN_DOTACAOEXERCICIO.HANDLE = SFN_DOTACAO.EXERCICIO]"

  vCriterio = ""

  FiltrarEmpenho = INTERFACE.Exec(CurrentSystem, vTabela, vColunas, 3, vCampos, vCriterio, "Empenho", True, texto)

End Function

Public Function FiltrarNaturezaDespesa(texto As String) As Long

  Dim INTERFACE As Object
  Dim vColunas, vCriterio, vCampos, vTabela As String
  Set INTERFACE = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_DOTACAOEXERCICIO.EXERCICIO|NATUREZADESPESA.DESCRICAO|UGRESPONSAVEL.DESCRICAO"
  vCampos = "Exercício|Natureza|Dotação"
  vTabela = "SFN_DOTACAONATUREZA|"
  vTabela = vTabela + "NATUREZADESPESA[SFN_DOTACAONATUREZA.NATUREZADESPESA=NATUREZADESPESA.HANDLE]|"
  vTabela = vTabela + "SFN_DOTACAO[SFN_DOTACAO.HANDLE = SFN_DOTACAONATUREZA.DOTACAO]|"
  vTabela = vTabela + "UGRESPONSAVEL[UGRESPONSAVEL.HANDLE = SFN_DOTACAO.UGRESPONSAVEL]|"
  vTabela = vTabela + "SFN_DOTACAOEXERCICIO[SFN_DOTACAOEXERCICIO.HANDLE = SFN_DOTACAO.EXERCICIO]"

  vCriterio = ""

  FiltrarNaturezaDespesa = INTERFACE.Exec(CurrentSystem, vTabela, vColunas, 2, vCampos, vCriterio, "Natureza da Despesa", True, texto)

End Function

Public Sub VERIFICARNATUREZADESPESA_OnPopup(ShowPopup As Boolean)

  ShowPopup = False

  Dim vHandle As Long

  vCriterio = ""

  vHandle = FiltrarNaturezaDespesa(VERIFICARNATUREZADESPESA.Text)

  If vHandle <> 0 Then
    	CurrentQuery.FieldByName("VERIFICARNATUREZADESPESA").AsInteger = vHandle
  End If

End Sub

Public Sub VERIFICARNATUREZADESPESAENCARG_OnPopup(ShowPopup As Boolean)

  ShowPopup = False

  Dim vHandle As Long

  vCriterio = ""

  vHandle = FiltrarNaturezaDespesa(VERIFICARNATUREZADESPESAENCARG.Text)

  If vHandle <> 0 Then
    	CurrentQuery.FieldByName("VERIFICARNATUREZADESPESAENCARG").AsInteger = vHandle
  End If

End Sub

Public Function FiltrarDotacao(texto As String) As Long

  Dim INTERFACE As Object
  Dim vColunas, vCriterio, vCampos, vTabela As String
  Set INTERFACE = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_DOTACAOEXERCICIO.EXERCICIO|UGRESPONSAVEL.DESCRICAO"
  vCampos = "Exercício|Dotação"
  vTabela = "SFN_DOTACAO|"
  vTabela = vTabela + "UGRESPONSAVEL[UGRESPONSAVEL.HANDLE = SFN_DOTACAO.UGRESPONSAVEL]|"
  vTabela = vTabela + "SFN_DOTACAOEXERCICIO[SFN_DOTACAOEXERCICIO.HANDLE = SFN_DOTACAO.EXERCICIO]"

  vCriterio = ""

  FiltrarDotacao = INTERFACE.Exec(CurrentSystem, vTabela, vColunas, 2, vCampos, vCriterio, "Dotação", True, texto)

End Function

Public Sub VERIFICARDOTACAO_OnPopup(ShowPopup As Boolean)

  ShowPopup = False

  Dim vHandle As Long

  vCriterio = ""

  vHandle = FiltrarDotacao(VERIFICARDOTACAO.Text)

  If vHandle <> 0 Then
    	CurrentQuery.FieldByName("VERIFICARDOTACAO").AsInteger = vHandle
  End If

End Sub

Public Sub VERIFICARDOTACAOENCARGO_OnPopup(ShowPopup As Boolean)

  ShowPopup = False

  Dim vHandle As Long

  vCriterio = ""

  vHandle = FiltrarDotacao(VERIFICARDOTACAOENCARGO.Text)

  If vHandle <> 0 Then
    	CurrentQuery.FieldByName("VERIFICARDOTACAOENCARGO").AsInteger = vHandle
  End If

End Sub
