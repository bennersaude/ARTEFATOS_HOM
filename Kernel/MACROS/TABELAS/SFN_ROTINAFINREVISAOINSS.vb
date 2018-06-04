'HASH: 031950D6C17B5A27DB79F1FCE2B11EDE

'Macro: SFN_ROTINAFINREVISAOINSS


'#Uses "*bsShowMessage

Option Explicit

Public Sub BOTAOCANCELAR_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  Dim vsMensagem As String

  If CurrentQuery.FieldByName("SITUACAO").AsInteger = 1 Then
    bsShowMessage("A Rotina não foi processada", "I")
    Exit Sub
  End If

  If VisibleMode Then
    Set Obj = CreateBennerObject("BSINTERFACE0027.CancelaRevisaoINSS")
    Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMensagem)

  Else
    Dim viRetorno As Long
    Dim vcContainer As CSDContainer
    Set vcContainer = NewContainer

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "SfnRecolhimento", _
                                     "CancelaRevisaoINSS", _
                                     "Rotina de Cancelamento de Revisão de INSS", _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SFN_ROTINAFINREVISAOINSS", _
                                     "SITUACAO", _
                                     "", _
                                     "", _
                                     "C", _
                                     False, _
                                     vsMensagem, _
                                     Null)

      If viRetorno = 0 Then
       bsShowMessage("Processo enviado para execução no servidor!", "I")
      Else
       bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagem, "I")
      End If
   End If

   CurrentQuery.Active = False
   CurrentQuery.Active = True
   Set Obj = Nothing
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <>1 Then
   bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  Dim vsMensagem As String

  If CurrentQuery.FieldByName("SITUACAO").AsInteger <> 1 Then
    bsShowMessage("A Rotina não está aberta", "I")
    Exit Sub
  End If

  If VisibleMode Then
    Set Obj = CreateBennerObject("BSINTERFACE0027.RevisaoINSS")
    Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMensagem)

  Else
    Dim viRetorno As Long

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "SfnRecolhimento", _
                                     "RevisaoINSS", _
                                     "Rotina de Revisão de INSS", _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SFN_ROTINAFINREVISAOINSS", _
                                     "SITUACAO", _
                                     "", _
                                     "", _
                                     "P", _
                                     False, _
                                     vsMensagem, _
                                     Null)

      If viRetorno = 0 Then
       bsShowMessage("Processo enviado para execução no servidor!", "I")
      Else
       bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagem, "I")
      End If
   End If

   CurrentQuery.Active = False
   CurrentQuery.Active = True

   Set Obj = Nothing

End Sub

Public Sub PESSOA_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_PESSOA.CNPJCPF|SFN_PESSOA.NOME"

  vCriterio = "HANDLE>0"

  vCampos = "CNPJ/CPF|Nome"

  vHandle = interface.Exec(CurrentSystem, "SFN_PESSOA", vColunas, 2, vCampos, vCriterio, "Procura Pessoa", False, PESSOA.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PESSOA").AsInteger = vHandle
  End If

  Set interface = Nothing
End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_PRESTADOR.CPFCNPJ|SAM_PRESTADOR.NOME"

  vCriterio = "HANDLE>0"

  vCampos = "CNPJ/CPF|Nome"

  vHandle = interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 2, vCampos, vCriterio, "Procura Prestador", False, PRESTADOR.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").AsInteger = vHandle
  End If

  Set interface = Nothing
End Sub

Public Sub TABLE_AfterScroll()

	BOTAOPROCESSAR.Enabled = CurrentQuery.FieldByName("SITUACAO").AsString = "1"
	BOTAOCANCELAR.Enabled = CurrentQuery.FieldByName("SITUACAO").AsString = "5"

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If CurrentQuery.FieldByName("SITUACAO").AsInteger <> 1 Then
	    CanContinue = False
    	bsShowMessage("Rotina não está aberta. Alteração não permitida!", "E")
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT C.COMPETENCIA")
  SQL.Add("FROM SFN_ROTINAFIN R, SFN_COMPETFIN C")
  SQL.Add("WHERE R.COMPETFIN=C.HANDLE AND R.HANDLE=:ROTINAFIN")
  SQL.ParamByName("ROTINAFIN").AsInteger = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
  SQL.Active = True


  If Year(CurrentQuery.FieldByName("DATAGERACAO").AsDateTime)<Year(SQL.FieldByName("COMPETENCIA").AsDateTime) Or _
     (Year(CurrentQuery.FieldByName("DATAGERACAO").AsDateTime)=Year(SQL.FieldByName("COMPETENCIA").AsDateTime) And _
  		   Month(SQL.FieldByName("COMPETENCIA").AsDateTime)>Month(CurrentQuery.FieldByName("DATAGERACAO").AsDateTime)) Then
    bsShowMessage("A data de geração não pode ser menor que a competência da rotina !", "E")
    CanContinue = False
    Exit Sub
  End If


  If WebMode Then
	PESSOA.WebLocalWhere = "HANDLE>0"
	PRESTADOR.WebLocalWhere = "HANDLE>0"
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOPROCESSAR"
			BOTAOPROCESSAR_OnClick
	End Select
End Sub
