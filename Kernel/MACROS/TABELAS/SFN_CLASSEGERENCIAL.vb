'HASH: A18305465624947FBDBBE45CF8AAA6FB
'Macro: SFN_CLASSEGERENCIAL
'#Uses "*bsShowMessage"

Public Sub BOTAODUPLICAR_OnClick()
  Dim viRetorno As Long
  Dim vcContainer     As Object
  Dim BSINTERFACE0002 As Object
  Dim vsMensagem      As String

  Set vcContainer = NewContainer
  Set BSINTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

  viRetorno = BSINTERFACE0002.Exec(CurrentSystem, _
								   1, _
								   "TV_FORM0036", _
								   "Duplica estrutura", _
								   0, _
								   140, _
								   310, _
								   False, _
								   vsMensagem, _
								   vcContainer)

  Set vcContainer = Nothing

  Select Case viRetorno
	Case -1
		bsShowMessage("Operação cancelada pelo usuário!", "I")
	Case 1
		bsShowMessage(vsMensagem, "I")
  End Select
  Set BSINTERFACE0002 = Nothing

End Sub

Public Sub BOTAOPADRONIZACAO_OnClick()
  Dim interface As Object
  If CurrentQuery.FieldByName("ULTIMONIVEL").AsString = "S" Then
    bsShowMessage("Último nível. Não existem registros a serem atulizados.", "I")
    Exit Sub
  End If

  Set interface = CreateBennerObject("SfnGerencial.Rotinas")
  interface.ReplicaClasseGer(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set interface = Nothing
End Sub

Public Sub CLASSECORRECAO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_CLASSEGERENCIAL.ESTRUTURA|SFN_CLASSEGERENCIAL.DESCRICAO|SFN_CLASSEGERENCIAL.CODIGOREDUZIDO|SFN_CLASSEGERENCIAL.NATUREZA|SFN_CLASSEGERENCIAL.HISTORICO"

  vCriterio = "HANDLE>0"

  vCampos = "Estrutura|Descrição|Código|D/C|Historico"

  vHandle = interface.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vColunas, 1, vCampos, vCriterio, "Classe Gerencial", False, CLASSECORRECAO.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLASSECORRECAO").Value = vHandle
  End If
  Set interface = Nothing

End Sub

Public Sub CLASSEDESCONTO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String


  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_CLASSEGERENCIAL.ESTRUTURA|SFN_CLASSEGERENCIAL.DESCRICAO|SFN_CLASSEGERENCIAL.CODIGOREDUZIDO|SFN_CLASSEGERENCIAL.NATUREZA|SFN_CLASSEGERENCIAL.HISTORICO"

  vCriterio = "HANDLE>0"

  vCampos = "Estrutura|Descrição|Código|D/C|Historico"

  vHandle = interface.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vColunas, 1, vCampos, vCriterio, "Classe Gerencial", False, CLASSEDESCONTO.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLASSEDESCONTO").Value = vHandle
  End If
  Set interface = Nothing

End Sub

Public Sub CLASSEGERENCIALPFCONTRATO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  'FERNANDO SMS 14974

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_CLASSEGERENCIAL.ESTRUTURA|SFN_CLASSEGERENCIAL.DESCRICAO"

  vCriterio = "ULTIMONIVEL='S'"

  vCampos = "Estrutura|Descrição"

  vHandle = interface.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vColunas, 1, vCampos, vCriterio, "Classe Gerencial", False, CLASSEGERENCIALPFCONTRATO.LocateText)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLASSEGERENCIALPFCONTRATO").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub CLASSEGERENCIALSUPLEMENTACAOPF_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  'FERNANDO SMS 14974

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_CLASSEGERENCIAL.ESTRUTURA|SFN_CLASSEGERENCIAL.DESCRICAO"

  vCriterio = "ULTIMONIVEL='S'"

  vCampos = "Estrutura|Descrição"

  vHandle = interface.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vColunas, 1, vCampos, vCriterio, "Classe Gerencial", False, CLASSEGERENCIALSUPLEMENTACAOPF.LocateText)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLASSEGERENCIALSUPLEMENTACAOPF").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub CLASSEJURO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_CLASSEGERENCIAL.ESTRUTURA|SFN_CLASSEGERENCIAL.DESCRICAO|SFN_CLASSEGERENCIAL.CODIGOREDUZIDO|SFN_CLASSEGERENCIAL.NATUREZA|SFN_CLASSEGERENCIAL.HISTORICO"

  vCriterio = "HANDLE>0"

  vCampos = "Estrutura|Descrição|Código|D/C|Historico"

  vHandle = interface.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vColunas, 1, vCampos, vCriterio, "Classe Gerencial", False, CLASSEJURO.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLASSEJURO").Value = vHandle
  End If
  Set interface = Nothing

End Sub

Public Sub CLASSEMULTA_OnPopup(ShowPopup As Boolean)
  Dim interface As Object

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_CLASSEGERENCIAL.ESTRUTURA|SFN_CLASSEGERENCIAL.DESCRICAO|SFN_CLASSEGERENCIAL.CODIGOREDUZIDO|SFN_CLASSEGERENCIAL.NATUREZA|SFN_CLASSEGERENCIAL.HISTORICO"

  vCriterio = "HANDLE>0"

  vCampos = "Estrutura|Descrição|Código|D/C|Historico"

  vHandle = interface.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vColunas, 1, vCampos, vCriterio, "Classe Gerencial", True, CLASSEMULTA.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLASSEMULTA").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub CLASSEPREVISAO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_CLASSEGERENCIAL.ESTRUTURA|SFN_CLASSEGERENCIAL.DESCRICAO|SFN_CLASSEGERENCIAL.CODIGOREDUZIDO|SFN_CLASSEGERENCIAL.NATUREZA|SFN_CLASSEGERENCIAL.HISTORICO"

  vCriterio = "HANDLE>0"

  vCampos = "Estrutura|Descrição|Código|D/C|Historico"

  vHandle = interface.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vColunas, 1, vCampos, vCriterio, "Classe Gerencial", False, CLASSEPREVISAO.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLASSEPREVISAO").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub CLASSERESSARCIMENTOINSS_OnPopup(ShowPopup As Boolean)
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

  viHandle = dllProcura_Procurar.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vsColunas, 1, vsCampos, vsCriterio, "Classes Gerenciais", False, CLASSERESSARCIMENTOINSS.LocateText)

  If viHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLASSERESSARCIMENTOINSS").Value = viHandle
  End If
  Set dllProcura_Procurar = Nothing
End Sub

Public Sub CLASSERESSARCIMENTOISS_OnPopup(ShowPopup As Boolean)
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

  viHandle = dllProcura_Procurar.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vsColunas, 1, vsCampos, vsCriterio, "Classes Gerenciais", False, CLASSERESSARCIMENTOISS.LocateText)

  If viHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLASSERESSARCIMENTOISS").Value = viHandle
  End If
  Set dllProcura_Procurar = Nothing
End Sub

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

  viHandle = dllProcura_Procurar.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vsColunas, 1, vsCampos, vsCriterio, "Classes Gerenciais", False, CLASSETAXAADMINISTRACAOEMSUP.LocateText)

  If viHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLASSETAXAADMINISTRACAOEMSUP").Value = viHandle
  End If
  Set dllProcura_Procurar = Nothing
End Sub

Public Sub CLASSEVALOREXCEDENTE_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_CLASSEGERENCIAL.ESTRUTURA|SFN_CLASSEGERENCIAL.DESCRICAO"

  vCriterio = "ULTIMONIVEL='S'"

  vCampos = "Estrutura|Descrição"

  vHandle = interface.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vColunas, 1, vCampos, vCriterio, "Classe Gerencial", False, CLASSEVALOREXCEDENTE.LocateText)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLASSEVALOREXCEDENTE").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT TABTIPOGESTAO FROM EMPRESAS WHERE HANDLE=" + Str(CurrentCompany))
  SQL.Active = True

  If SQL.FieldByName("TABTIPOGESTAO").AsInteger <>3 Then
    TIPOATO.Visible = False
  Else
    TIPOATO.Visible = True
  End If

  Set SQL = Nothing

  If WebMode Then
  	CLASSECORRECAO.WebLocalWhere = HANDLE>0
	CLASSEDESCONTO.WebLocalWhere = HANDLE>0
	CLASSEGERENCIALPFCONTRATO.WebLocalWhere = "ULTIMONIVEL='S'"
	CLASSEGERENCIALSUPLEMENTACAOPF.WebLocalWhere = "ULTIMONIVEL='S'"
	CLASSEJURO.WebLocalWhere = "HANDLE>0"
    CLASSEMULTA.WebLocalWhere = "HANDLE>0"
    CLASSEPREVISAO.WebLocalWhere = "HANDLE>0"
    CLASSERESSARCIMENTOINSS.WebLocalWhere = ""
    CLASSERESSARCIMENTOISS.WebLocalWhere = ""
    CLASSETAXAADMINISTRACAOEMSUP.WebLocalWhere = ""
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Dim SQL1 As Object
  Set SQL = NewQuery
  Set SQL1 = NewQuery
  Dim vHandle As Long

  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT HANDLE ")
  SQL.Add("  FROM SFN_CLASSEGERENCIAL ")
  SQL.Add(" WHERE CODIGOREDUZIDO =:CODRED ")
  SQL.Add(" AND HANDLE <>:HANDLE")
  SQL.ParamByName("CODRED").AsInteger = CurrentQuery.FieldByName("CODIGOREDUZIDO").AsInteger
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  vHandle = SQL.FieldByName("HANDLE").AsInteger

  If vHandle > 0 Then
    bsShowMessage("Já existe uma classe gerencial cadastrada com este Código reduzido!", "E")
    Set SQL = Nothing
    Set SQL1 = Nothing
    CODIGOREDUZIDO.SetFocus
    CanContinue = False
    Exit Sub
  End If

  SQL1.Clear
  SQL1.Active = False
  SQL1.Add("SELECT HANDLE ")
  SQL1.Add("  FROM SFN_CLASSEGERENCIAL")
  SQL1.Add(" WHERE ESTRUTURA =:ESTRUTURA")
  SQL1.Add(" AND HANDLE <>:HANDLE")
  SQL1.ParamByName("ESTRUTURA").AsString = CurrentQuery.FieldByName("ESTRUTURA").AsString
  SQL1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL1.Active = True

  vHandle = SQL1.FieldByName("HANDLE").AsInteger

  If vHandle > 0 Then
    bsShowMessage("Já existe uma classe gerencial cadastrada com esta Estrutura!", "E")
    Set SQL = Nothing
    Set SQL1 = Nothing
    ESTRUTURA.SetFocus
    CanContinue = False
    Exit Sub
  End If

  Set SQL = Nothing
  Set SQL1 = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAODUPLICAR"
			BOTAODUPLICAR_OnClick
		Case "BOTAOPADRONIZACAO"
			BOTAOPADRONIZACAO_OnClick
		Case ""
	End Select
End Sub
