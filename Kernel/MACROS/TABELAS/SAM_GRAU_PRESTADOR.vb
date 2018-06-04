'HASH: F2EFF2E43CCA3EC469A80BFE720EE3C1
'SAM_GRAU_PRESTADOR
'#Uses "*bsShowMessage"

Option Explicit

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraGrau

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("GRAU").AsInteger = vHandle
	End If
End Sub

Public Function ProcuraGrau()As Long
	Dim interface As Object
	Dim vCampos As String
	Dim vColunas As String
	Set interface = CreateBennerObject("Procura.Procurar")

	vColunas = "SAM_GRAU.GRAU|SAM_GRAU.Z_DESCRICAO|SAM_GRAU.VERIFICAGRAUSVALIDOS"
	vCampos = "Código do Grau|Descrição|Graus Válidos"
	ProcuraGrau = interface.Exec(CurrentSystem, "SAM_GRAU", vColunas, 2, vCampos, "", "Graus de Atuação", True, GRAU.Text)

	Set interface = Nothing
End Function

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraPrestador

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("PRESTADOR").AsInteger = vHandle
	End If
End Sub

Public Function ProcuraPrestador()As Long
	Dim interface As Object
	Dim vCampos As String
	Dim vColunas As String
	Set interface = CreateBennerObject("Procura.Procurar")

	vColunas = "SAM_PRESTADOR.CPFCNPJ|SAM_PRESTADOR.NOME"
	vCampos = "CPF/CNPJ|Nome"
	ProcuraPrestador = interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 2, vCampos, "", "Prestadores", True, "")

	Set interface = Nothing
End Function

'WebMenuCode: T3874 - Prestador/Associações
'  		      T4178 - Prestador/Grupo Empresarial
'			  T3875 - Prestador/Pessoa Física
'			  T3876 - Prestador/Pessoa Jurídica
'			  T3878 - Prestador/Por Categoria
'			  T3877 - Prestador/Por Especialidade
'			  T3879 - Prestador/Tipo Prestador
'			  T3880 - Prestador/Todos

Public Sub TABLE_AfterScroll()
	If WebMode Then
		If (WebMenuCode = "T3874") Or (WebMenuCode = "T4178") Or (WebMenuCode = "T3875") Or (WebMenuCode = "T3876") Or _
		   (WebMenuCode = "T3878") Or (WebMenuCode = "T3877") Or (WebMenuCode = "T3879") Or (WebMenuCode = "T3880") Then 'SMS 95891 - Ricardo Rocha - 03/06/2008
		   PRESTADOR.ReadOnly = True
		End If
	End If
End Sub
