'HASH: C65AC1F8DDA7ABD876E8A8A02AB43B4C
'#Uses "*bsShowMessage"

Public Sub EVENTOPADRAO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  Dim vData As String
  Dim S As Object
  Dim Interface As Object
  Dim vColunas As String
  Dim vCriterio As String
  Dim vCampos As String

  ShowPopup = False

  Set Interface = CreateBennerObject("Procura.Procurar")
  vColunas = "ESTRUTURA|DESCRICAO"
  vCriterio = " SAM_TGE.ULTIMONIVEL = 'S' AND (SAM_TGE.ESTRUTURA LIKE '1.01.01.%' OR SAM_TGE.ESTRUTURA LIKE '0.00.01.%' OR SAM_TGE.ESTRUTURA LIKE '00.01.%') "

  vCampos = "Estrutura|Descrição"
  vHandle = Interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 1, vCampos, vCriterio, "Pessoa", True, "")

  If vHandle > 0 Then
    CurrentQuery.FieldByName("EVENTOPADRAO").AsInteger = vHandle
  End If

  Set Interface = Nothing


End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  Dim vData As String
  Dim S As Object
  Dim Interface As Object
  Dim vColunas As String
  Dim vCriterio As String
  Dim vCampos As String


  Set Interface = CreateBennerObject("Procura.Procurar")
  vColunas = "PRESTADOR|NOME"
  vCriterio = " PRESTADOR IS NOT NULL "

  vCampos = "Prestador|Nome"
  vHandle = Interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 1, vCampos, vCriterio, "Pessoa", True, "")

  If vHandle > 0 Then
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
    Dim SQL As Object
    Set SQL= NewQuery

    SQL.Add("SELECT RECEBEDOR,EXECUTOR,SOLICITANTE,LOCALEXECUCAO FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").Value = vHandle
    SQL.Active = True

    If Not SQL.FieldByName("RECEBEDOR").IsNull Then
      CurrentQuery.FieldByName("RECEBEDOR").AsString = SQL.FieldByName("RECEBEDOR").AsString
    End If
    If Not SQL.FieldByName("EXECUTOR").IsNull Then
      CurrentQuery.FieldByName("EXECUTOR").AsString = SQL.FieldByName("EXECUTOR").AsString
    End If
    If Not SQL.FieldByName("SOLICITANTE").IsNull Then
      CurrentQuery.FieldByName("SOLICITANTE").AsString = SQL.FieldByName("SOLICITANTE").AsString
    End If
    If Not SQL.FieldByName("LOCALEXECUCAO").IsNull Then
      CurrentQuery.FieldByName("LOCALEXECUCAO").AsString = SQL.FieldByName("LOCALEXECUCAO").AsString
    End If

  End If

  Set Interface = Nothing



End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim SQL As Object
	Dim sMSG As String
	Set SQL= NewQuery

	SQL.Add("SELECT RECEBEDOR,EXECUTOR,SOLICITANTE,LOCALEXECUCAO FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
	SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
	SQL.Active = True

	If Not SQL.FieldByName("EXECUTOR").IsNull And CurrentQuery.FieldByName("EXECUTOR").AsString = "S" And SQL.FieldByName("EXECUTOR").AsString = "N" Then
		sMSG = "Prestador não é executor!"
		bsShowMessage(sMSG, "E")
		CanContinue = False
		Exit Sub
	End If
	If Not SQL.FieldByName("RECEBEDOR").IsNull And CurrentQuery.FieldByName("RECEBEDOR").AsString = "S" And SQL.FieldByName("RECEBEDOR").AsString = "N" Then
		sMSG = "Prestador não é recebedor!"
		bsShowMessage(sMSG, "E")
		CanContinue = False
		Exit Sub
	End If
	If Not SQL.FieldByName("SOLICITANTE").IsNull And CurrentQuery.FieldByName("SOLICITANTE").AsString = "S" And SQL.FieldByName("SOLICITANTE").AsString = "N" Then
		sMSG = "Prestador não é solicitante!"
		bsShowMessage(sMSG, "E")
		CanContinue = False
		Exit Sub
	End If
	If Not SQL.FieldByName("LOCALEXECUCAO").IsNull And CurrentQuery.FieldByName("LOCALEXECUCAO").AsString = "S" And SQL.FieldByName("LOCALEXECUCAO").AsString = "N" Then
		sMSG = "Prestador não é local de execução!"
		bsShowMessage(sMSG, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
