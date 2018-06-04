'HASH: EC67ACA12AC83D860E30131D0B4F08EC

'################### CENTRAL DE ATENDIMENTO ##########################

Public Sub SCTEBENEFICIARIO_OnPopup(ShowPopup As Boolean)
  Dim vCriterios As String
  Dim vCampos As String
  Dim Interface As Variant
  Set Interface = CreateBennerObject("Procura.Procurar")
  vColunas = "BENEFICIARIO|NOME"
  vCriterios = ""
  vCampos = "Beneficiário|Nome"
  vHandle = Interface.Exec(CurrentSystem, "SAM_BENEFICIARIO", vColunas, 2, vCampos, vCriterios, "Tabela de beneficiários", False, "")
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("SCTEBENEFICIARIO").Value = vHandle
  End If
  ShowPopup = False
  Set Interface = Nothing
End Sub


Public Sub SCTEPRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim vColuna As String
  Dim vCriterios As String
  Dim vCampos As String
  Dim Interface As Variant
  Set Interface = CreateBennerObject("Procura.Procurar")
  vColunas = "PRESTADOR|NOME"
  vCriterios = ""
  vCampos = "CNPJ/CPF|Nome"
  vHandle = Interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 2, vCampos, vCriterios, "Tabela de prestadores", False, "")
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("SCTEPRESTADOR").Value = vHandle
  End If
  ShowPopup = False
  Set Interface = Nothing
End Sub

'######################################################################

Public Sub TABLE_AfterInsert()
Dim vANO As String
Dim Sequencia As Long
Dim AnoAtual As Long
Dim vSQL As Object

' CurrentQuery.FieldByName("SCTEBENEFICIARIO").Value=SessionVar("BENEFICIARIO")
	If WebMode Then
		Set vSQL = NewQuery
		vSQL.Clear
		vSQL.Add("SELECT FILIALPADRAO FROM Z_GRUPOUSUARIOS WHERE HANDLE = :HANDLE")
		vSQL.ParamByName("HANDLE").Value = CurrentUser
		vSQL.Active = True

		CurrentQuery.FieldByName("FILIAL").Value = vSQL.FieldByName("FILIALPADRAO").AsInteger

		vSQL.Clear
		vSQL.Add("SELECT SEVERIDADE, TIPOSERVICO, CLASSEOUTROSSERVICOS FROM SAM_PARAMETROSWEB")
		vSQL.Active = True

		If vSQL.FieldByName("SEVERIDADE").IsNull Then
			Err.Raise(1,,"Falta parametrizar severidade em Parâmetros da Web.")
			Exit Sub
		End If
		If vSQL.FieldByName("TIPOSERVICO").IsNull Then
			Err.Raise(1,,"Falta parametrizar Tipo de Serviço em Parâmetros da Web.")
			Exit Sub
		End If
		If vSQL.FieldByName("CLASSEOUTROSSERVICOS").IsNull Then
			Err.Raise(1,,"Falta parametrizar Classificação em Parâmetros da Web.")
			Exit Sub
		End If

		CurrentQuery.FieldByName("SEVERIDADE").Value = vSQL.FieldByName("SEVERIDADE").AsInteger
		CurrentQuery.FieldByName("TIPOSERVICO").Value = vSQL.FieldByName("TIPOSERVICO").AsInteger
		CurrentQuery.FieldByName("CLASSEOUTROSSERVICOS").Value = vSQL.FieldByName("CLASSEOUTROSSERVICOS").AsInteger


		vANO = Format(ServerDate,"yyyy")
		CurrentQuery.FieldByName("ANO").Value = ("01/01/" + vANO)
		AnoAtual = CLng(vANO)
		NewCounter("CA_OUTROS", AnoAtual, 1, Sequencia)
		CurrentQuery.FieldByName("NUMERO").Value = Sequencia
		CurrentQuery.FieldByName("PROTOCOLO").Value = Format(ServerDate,"yyyy") + Format(Sequencia,"######000000")
		CurrentQuery.FieldByName("TABSOLICITADO").Value = 3
		CurrentQuery.FieldByName("TABRESPOSTA").Value = 4
		CurrentQuery.FieldByName("DATASERVICO").AsDateTime = ServerNow

		vSQL.Active = False
		vSQL.Clear
		vSQL.Add("SELECT PRESTADOR FROM Z_GRUPOUSUARIOS_PRESTADOR WHERE USUARIO = :USUARIO")
		vSQL.ParamByName("USUARIO").AsInteger = CurrentUser
		vSQL.Active = True

		If Not vSQL.FieldByName("PRESTADOR").IsNull Then
		  CurrentQuery.FieldByName("SCTEPRESTADOR").AsInteger = vSQL.FieldByName("PRESTADOR").AsInteger
		  CurrentQuery.FieldByName("TABSOLICITANTE").AsInteger=2
	End If


		Set vSQL = Nothing






	End If
End Sub

Public Sub TABLE_AfterPost()
	If WebMode Then
		InfoDescription = "Por favor, anote o número do protocolo: " + CurrentQuery.FieldByName("PROTOCOLO").AsString
	End If
End Sub

Public Sub TABLE_AfterScroll()
  SCTEBENEFICIARIO.ResultFields = "BENEFICIARIO|NOME|"
  SCTEPRESTADOR.ResultFields = "PRESTADOR|NOME|"
  SCDOBENEFICIARIO.ResultFields = "BENEFICIARIO|NOME|"
  SCDOPRESTADOR.ResultFields = "PRESTADOR|NOME|"
  RESPONSAVELUSUARIO.ResultFields = "APELIDO|NOME|"
  RECOMPRESTADOR.ResultFields = "PRESTADOR|NOME|"
  RECOMBENEFICIARIO.ResultFields = "BENEFICIARIO|NOME|"
  SEVERIDADE.ResultFields = "CODIGO|DESCRICAO|"

  If CurrentQuery.FieldByName("PROTOCOLOOUVIDORIA").IsNull Then
    PROTOCOLOOUVIDORIA.Visible = False
    PRAZOOUVIDORIA.Visible = False
    ANEXO.Visible = False
    HISTORICOOUVIDORIA.Visible = False
  Else
    PROTOCOLOOUVIDORIA.Visible = True
    PRAZOOUVIDORIA.Visible = True
    ANEXO.Visible = True
    HISTORICOOUVIDORIA.Visible = True
  End If


End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vANO As Integer
  Dim Sequencia As Long
  vANO = DatePart("yyyy", CurrentQuery.FieldByName("ANO").AsDateTime)
  NewCounter("CA_OUTROS", vANO, 1, Sequencia)
  If CurrentQuery.State = 3 Then
    CurrentQuery.FieldByName("NUMERO").Value = Sequencia
  End If

  If VisibleMode Then

  If CurrentQuery.FieldByName("TABSOLICITANTE").Value = 1 Then
    If CurrentQuery.FieldByName("sctebeneficiario").IsNull Then
      MsgBox("Beneficiário solicitante é obrigaório.")
      CanContinue = False
      Exit Sub
    End If
  Else
    If CurrentQuery.FieldByName("TABSOLICITANTE").Value = 2 Then
      If CurrentQuery.FieldByName("scteprestador").IsNull Then
        MsgBox("Prestador solicitante é obrigaório.")
        CanContinue = False
        Exit Sub
      End If
    Else
      If CurrentQuery.FieldByName("TABSOLICITANTE").Value = 3 Then
        If CurrentQuery.FieldByName("outronome").IsNull Then
          MsgBox("Nome do solicitante é obrigaório.")
          CanContinue = False
          Exit Sub
        End If
      End If
    End If
  End If
  End If

End Sub

Public Sub TABLE_UpdateRequired()
If WebMode Then
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT SEVERIDADE, TIPOSERVICO, CLASSEOUTROSSERVICOS FROM SAM_PARAMETROSWEB")
  SQL.Active = True

  If CurrentQuery.FieldByName("TIPOSERVICO").AsInteger <> SQL.FieldByName("TIPOSERVICO").AsInteger Then
    SQL.Active = False

    SQL.Clear
    SQL.Add("SELECT HANDLE FROM CA_CLASSEOUTROS WHERE TIPOSERVICOPADRAO = :TIPO AND CODIGO = ")
    SQL.Add("(SELECT MIN(HANDLE) FROM CA_CLASSEOUTROS WHERE TIPOSERVICOPADRAO = :TIPO) ")
    SQL.ParamByName("TIPO").AsInteger = CurrentQuery.FieldByName("TIPOSERVICO").AsInteger
    SQL.Active = True

    CurrentQuery.FieldByName("CLASSEOUTROSSERVICOS").AsInteger = SQL.FieldByName("HANDLE").AsInteger
	End If

  If Not CurrentQuery.FieldByName("TELEFONERESPOSTA").IsNull Then
    CurrentQuery.FieldByName("TABRESPOSTA").AsInteger = 3
  End If

  If Not CurrentQuery.FieldByName("FAXRESPOSTA").IsNull Then
    CurrentQuery.FieldByName("TABRESPOSTA").AsInteger = 2
  End If


  If WebVisionCode = "W_CA_SOLICITOUTROS" Then


    SQL.Active = False

    If CurrentQuery.FieldByName("EMAILRESPOSTA").IsNull Then
      SQL.Clear
      SQL.Clear
      SQL.Add("SELECT EMAIL FROM SAM_BENEFICIARIO WHERE HANDLE = :HANDLE")
      SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("sctebeneficiario").AsInteger
      SQL.Active = True

      If (Not SQL.FieldByName("EMAIL").IsNull) And (CurrentQuery.FieldByName("EMAILRESPOSTA").IsNull) Then
        CurrentQuery.FieldByName("EMAILRESPOSTA").AsString = SQL.FieldByName("EMAIL").AsString
      End If
    End If
    CurrentQuery.FieldByName("SCTEPRESTADOR").Clear
    CurrentQuery.FieldByName("TABSOLICITANTE").AsInteger=1

  End If

  If WebVisionCode = "W_CA_SOLICITOUTROS_PRESTADOR" Then
    If CurrentQuery.FieldByName("EMAILRESPOSTA").IsNull Then
      SQL.Clear
      SQL.Clear
      SQL.Add("SELECT EMAIL FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
      SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("SCTEPRESTADOR").AsInteger
      SQL.Active = True

      If (Not SQL.FieldByName("EMAIL").IsNull) And (CurrentQuery.FieldByName("EMAILRESPOSTA").IsNull) Then
        CurrentQuery.FieldByName("EMAILRESPOSTA").AsString = SQL.FieldByName("EMAIL").AsString
      End If
      CurrentQuery.FieldByName("SCTEBENEFICIARIO").Clear
      CurrentQuery.FieldByName("TABSOLICITANTE").AsInteger=2
    End If


  End If

  Set SQL = Nothing



  End If
End Sub
