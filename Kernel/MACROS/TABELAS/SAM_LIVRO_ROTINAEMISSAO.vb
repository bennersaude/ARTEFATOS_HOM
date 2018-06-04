'HASH: AAE9D5E40478E88BCA55BB6B12F8C375
'#Uses "*bsShowMessage"

Public Sub BOTAOPROCESSAR_OnClick()
	Dim SQL As Object
	Set SQL = NewQuery

	If CurrentQuery.State = 2 Or CurrentQuery.State = 3 Then
		bsShowMessage("Registro em edição !", "I")
		Exit Sub
	End If

	If CurrentQuery.FieldByName("TIPOEMISSAO").AsString = "E" Then
		SQL.Add("SELECT * FROM SAM_LIVROENCARTE WHERE LIVRO = :LIVRO")
		SQL.ParamByName("LIVRO").Value = CurrentQuery.FieldByName("LIVRO").Value
		SQL.Active = True

		If SQL.EOF Then
			bsShowMessage("Este livro não possue encartes !", "I")
			Set SQL = Nothing
			Exit Sub
		End If
	End If

	'SMS 87652 - Ricardo Rocha - 19/12/2007
	'Adequacao para Web
	If CurrentQuery.FieldByName("TIPORELATORIO").AsString = "C" Then
		SessionVar("HRotinaEmissaoCompleto") = CStr(CurrentQuery.FieldByName("HANDLE").AsInteger)
	Else
		SessionVar("HRotinaEmissaoResumido") = CStr(CurrentQuery.FieldByName("HANDLE").AsInteger)
	End If

	If CurrentQuery.FieldByName("TIPORELATORIO").AsString = "C" Then
		SQL.Active = False
		SQL.Clear
		SQL.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'PRE060'")
		SQL.Active = True

		If SQL.EOF Then
			bsShowMessage("Relatório 'PRE060' não encontrado !" + Chr(10) + _
										"Verique se: " + Chr(10) + _
										"- o relatório foi importado;" + Chr(10) + _
										"- o código do relatório está correto." + Chr(10) + _
										Chr(10) + _
										"Para importar ou corrigir o código entre no" + Chr(10) + _
										"Módulo ´Adm´ / carga: ´Gerador de relatórios/Relatórios/...", "I")
			Exit Sub
		Else
			ReportPreview(SQL.FieldByName("HANDLE").Value, "", True, False)
			RefreshNodesWithTable("SAM_LIVRO_ROTINAEMISSAO")
		End If
	Else
		SQL.Active = False
		SQL.Clear
		SQL.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'CRE062'")
		SQL.Active = True

		If SQL.EOF Then
			bsShowMessage("Relatório 'CRE062' não encontrado !" + Chr(10) + _
										"Verique se: " + Chr(10) + _
										"- o relatório foi importado;" + Chr(10) + _
										"- o código do relatório está correto." + Chr(10) + _
										Chr(10) + _
										"Para importar ou corrigir o código entre no" + Chr(10) + _
										"Módulo ´Adm´ / carga: ´Gerador de relatórios/Relatórios/...", "I")
			Exit Sub
		Else
			ReportPreview(SQL.FieldByName("HANDLE").Value, "", True, False)
			RefreshNodesWithTable("SAM_LIVRO_ROTINAEMISSAO")
		End If
	End If

  Set SQL = Nothing

End Sub

Public Sub LIVRO_OnPopup(ShowPopup As Boolean)
	CurrentQuery.FieldByName("LIVROENCARTE").Clear
End Sub

Public Sub TABLE_AfterScroll()
  If Not (CurrentQuery.FieldByName("CONFIGURACAO").IsNull) Then
	Dim Query As Object
	Set Query = NewQuery

	Query.Add("SELECT EMITIRAFASTADOS, ")
	Query.Add("		  EMITIRBLOQUEADOS ")
	Query.Add("  FROM SAM_LIVROCONFIG  ")
	Query.Add(" WHERE HANDLE = :HANDLE ")
	Query.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CONFIGURACAO").AsInteger
	Query.Active = True

	If Query.FieldByName("EMITIRAFASTADOS").AsString = "S" Then
		EMITIRAFASTADOS.Text = "Emitir Afastados: Sim"
	Else
		EMITIRAFASTADOS.Text = "Emitir Afastados: Não"
	End If

	If Query.FieldByName("EMITIRBLOQUEADOS").AsString = "S" Then
		EMITIRBLOQUEADOS.Text = "Emitir Bloqueados: Sim"
	Else
		EMITIRBLOQUEADOS.Text = "Emitir Bloqueados: Não"
	End If

		ExecutarFiltros
  End If

  vCondicao = ""

  If VisibleMode Then
    vCondicao = vCondicao + "SAM_LIVROENCARTE.HANDLE "
    vCondicao = vCondicao + "IN (SELECT HANDLE FROM SAM_LIVROENCARTE WHERE LIVRO = @LIVRO)"

  	LIVROENCARTE.LocalWhere = vCondicao
  Else
    vCondicao = vCondicao + "A.HANDLE "
    vCondicao = vCondicao + "IN (SELECT HANDLE FROM SAM_LIVROENCARTE WHERE LIVRO = @CAMPO(LIVRO))"

  	LIVROENCARTE.WebLocalWhere = vCondicao
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If CurrentQuery.FieldByName("TIPOEMISSAO").AsString = "E" And CurrentQuery.FieldByName("LIVROENCARTE").IsNull Then
		bsShowMessage("Campo Encarte obrigatório !", "E")
		CanContinue = False
		Exit Sub
	Else
		If CurrentQuery.FieldByName("TIPOEMISSAO").AsString = "C" And Not CurrentQuery.FieldByName("LIVROENCARTE").IsNull Then
			CurrentQuery.FieldByName("LIVROENCARTE").Value = Null
		End If
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOPROCESSAR"
			BOTAOPROCESSAR_OnClick
	End Select
End Sub

Public Sub ExecutarFiltros
  Dim SQL As Object
  Set SQL = NewQuery
  Dim vContador As Integer
  Dim vAreas As String
  Dim vEspecialidades As String
  Dim vEstados As String
  Dim vMunicipios As String
  Dim vPrestadores As String
  Dim vRede As String
  Dim vRegiao As String
  Dim vTipos As String
  Dim vFiltro As String

  'Áreas do Livro
  vAreas = "Áreas: "
  vContador = 0
  SQL.Clear
  SQL.Add(" SELECT A.DESCRICAO ")
  SQL.Add("   FROM SAM_AREALIVRO            A ")
  SQL.Add("   JOIN SAM_LIVROCONFIG_FILTROAREA F ON (F.AREALIVRO = A.HANDLE) ")
  SQL.Add("  WHERE F.LIVROCONFIGURACAO =  " + CurrentQuery.FieldByName("CONFIGURACAO").AsString)
  SQL.Active = True
  While (Not SQL.EOF)
  	If vContador < 1 Then
  	  vAreas = vAreas + SQL.FieldByName("DESCRICAO").AsString
  	Else
      vAreas = vAreas + ", " + SQL.FieldByName("DESCRICAO").AsString
    End If
    vContador = vContador + 1
    SQL.Next
  Wend
  If (vContador > 0) Then
    vAreas= vAreas + Chr(13)
  Else
    vAreas = ""
  End If

  'Especialidades
  vEspecialidades = "Especialidades: "
  vContador = 0
  SQL.Clear
  SQL.Add(" SELECT A.DESCRICAO ")
  SQL.Add("   FROM SAM_ESPECIALIDADE           A ")
  SQL.Add("   JOIN SAM_LIVROCONFIG_FILTROESPEC E ON (E.ESPECIALIDADE = A.HANDLE) ")
  SQL.Add("  WHERE E.LIVROCONFIGURACAO =  " + CurrentQuery.FieldByName("CONFIGURACAO").AsString)
  SQL.Active = True
  While (Not SQL.EOF)
  	If vContador < 1 Then
  	  vEspecialidades = vEspecialidades + SQL.FieldByName("DESCRICAO").AsString
  	Else
      vEspecialidades = vEspecialidades + ", " + SQL.FieldByName("DESCRICAO").AsString
    End If
    vContador = vContador + 1
    SQL.Next
  Wend
  If (vContador > 0) Then
    vEspecialidades = vEspecialidades + Chr(13)
  Else
    vEspecialidades = ""
  End If

  'Estados
  vEstados = "Estados: "
  vContador = 0
  SQL.Clear
  SQL.Add(" SELECT A.NOME ")
  SQL.Add("   FROM ESTADOS                      A ")
  SQL.Add("   JOIN SAM_LIVROCONFIG_FILTROESTADO E ON (E.ESTADO = A.HANDLE) ")
  SQL.Add("  WHERE E.LIVROCONFIGURACAO =  " + CurrentQuery.FieldByName("CONFIGURACAO").AsString)
  SQL.Active = True
  While (Not SQL.EOF)
  	If vContador < 1 Then
	  vEstados = vEstados + SQL.FieldByName("NOME").AsString
	Else
      vEstados = vEstados + ", " + SQL.FieldByName("NOME").AsString
    End If
    vContador = vContador + 1
    SQL.Next
  Wend
  If (vContador > 0) Then
    vEstados = vEstados + Chr(13)
  Else
    vEstados = ""
  End If

  'Municipios
  vMunicipios = "Municípios: "
  vContador = 0
  SQL.Clear
  SQL.Add(" SELECT A.NOME ")
  SQL.Add("   FROM MUNICIPIOS                  A ")
  SQL.Add("   JOIN SAM_LIVROCONFIG_FILTROMUNIC M ON (M.MUNICIPIO = A.HANDLE) ")
  SQL.Add("  WHERE M.LIVROCONFIGURACAO =  " + CurrentQuery.FieldByName("CONFIGURACAO").AsString)
  SQL.Active = True
  While (Not SQL.EOF)
  	If vContador < 1 Then
      vMunicipios = vMunicipios + SQL.FieldByName("NOME").AsString
    Else
      vMunicipios = vMunicipios + ", " + SQL.FieldByName("NOME").AsString
    End If
    vContador = vContador + 1
    SQL.Next
  Wend
  If (vContador > 0) Then
    vMunicipios = vMunicipios + Chr(13)
  Else
    vMunicipios = ""
  End If

  'Prestadores
  vPrestadores = "Prestadores: "
  vContador = 0
  SQL.Clear
  SQL.Add(" SELECT A.NOME ")
  SQL.Add("   FROM SAM_PRESTADOR               A ")
  SQL.Add("   JOIN SAM_LIVROCONFIG_FILTROPREST P ON (P.PRESTADOR = A.HANDLE) ")
  SQL.Add("  WHERE P.LIVROCONFIGURACAO =  " + CurrentQuery.FieldByName("CONFIGURACAO").AsString)
  SQL.Active = True
  While (Not SQL.EOF)
    If vContador < 1 Then
	  vPrestadores = vPrestadores + SQL.FieldByName("NOME").AsString
	Else
      vPrestadores = vPrestadores + ", " + SQL.FieldByName("NOME").AsString
    End If
    vContador = vContador + 1
    SQL.Next
  Wend
  If (vContador > 0) Then
    vPrestadores = vPrestadores + Chr(13)
  Else
    vPrestadores = ""
  End If

  'Prestadores
  vRede = "Redes Restritas: "
  vContador = 0
  SQL.Clear
  SQL.Add(" SELECT A.DESCRICAO ")
  SQL.Add("   FROM SAM_REDERESTRITA           A ")
  SQL.Add("   JOIN SAM_LIVROCONFIG_FILTROREDE R ON (R.REDERESTRITA = A.HANDLE) ")
  SQL.Add("  WHERE R.LIVROCONFIGURACAO =  " + CurrentQuery.FieldByName("CONFIGURACAO").AsString)
  SQL.Active = True
  While (Not SQL.EOF)
    If vContador < 1 Then
      vRede = vRede + SQL.FieldByName("DESCRICAO").AsString
    Else
      vRede = vRede + ", " + SQL.FieldByName("DESCRICAO").AsString
    End If
    vContador = vContador + 1
    SQL.Next
  Wend
  If (vContador > 0) Then
    vRede = vRede + Chr(13)
  Else
    vRede = ""
  End If

  'Regiões
  vRegiao = "Regiões: "
  vContador = 0
  SQL.Clear
  SQL.Add(" SELECT A.NOME ")
  SQL.Add("   FROM SAM_REGIAO                   A ")
  SQL.Add("   JOIN SAM_LIVROCONFIG_FILTROREGIAO R ON (R.REGIAO = A.HANDLE) ")
  SQL.Add("  WHERE R.LIVROCONFIGURACAO =  " + CurrentQuery.FieldByName("CONFIGURACAO").AsString)
  SQL.Active = True
  While (Not SQL.EOF)
    If vContador < 1 Then
      vRegiao = vRegiao + SQL.FieldByName("NOME").AsString
    Else
      vRegiao = vRegiao + ", " + SQL.FieldByName("NOME").AsString
    End If
    vContador = vContador + 1
    SQL.Next
  Wend
  If (vContador > 0) Then
    vRegiao = vRegiao + Chr(13)
  Else
    vRegiao = ""
  End If

  'Tipo Prestador
  vTipos = "Tipos de Prestadores: "
  vContador = 0
  SQL.Clear
  SQL.Add(" SELECT A.DESCRICAO ")
  SQL.Add("   FROM SAM_TIPOPRESTADOR          A ")
  SQL.Add("   JOIN SAM_LIVROCONFIG_FILTROTIPO T ON (T.TIPOPRESTADOR = A.HANDLE) ")
  SQL.Add("  WHERE T.LIVROCONFIGURACAO =  " + CurrentQuery.FieldByName("CONFIGURACAO").AsString)
  SQL.Active = True
  While (Not SQL.EOF)
    If vContador < 1 Then
      vTipos = vTipos + SQL.FieldByName("DESCRICAO").AsString
    Else
      vTipos = vTipos + ", " + SQL.FieldByName("DESCRICAO").AsString
    End If
    vContador = vContador + 1
    SQL.Next
  Wend
  If (vContador > 0) Then
    vTipos = vTipos + Chr(13)
  Else
    vTipos = ""
  End If

  vFiltro = vAreas + vEspecialidades + vEstados + vMunicipios + vPrestadores + vRede + vRegiao + vTipos

  FILTRO.Text = vFiltro

End Sub
