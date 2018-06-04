'HASH: B14365359ECD58C96D2A004E92FF38E7

'Macro: SAM_ROTINASIMULACAO
Option Explicit

'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"


Public Sub BOTAOPLANILHA_OnClick()
  If CurrentQuery.State <> 1 Then
		bsShowMessage("Os parâmetros não podem estar em edição", "I")
		Exit Sub
  End If

  Dim interface As Object

  If VisibleMode Then
    Set interface = CreateBennerObject("BSINTERFACE0035.SimulacaoReajuste")
    interface.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Else
    Dim vsmensagemerro As String
    Dim viRetorno      As Long
    Dim vcContainer    As CSDContainer

    Set vcContainer = NewContainer

    vcContainer.AddFields("HANDLE:INTEGER")
    vcContainer.Insert
    vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

    Set interface = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = interface.ExecucaoImediata(CurrentSystem, _
                                           "BSPre008", _
                                           "SimulacaoReajuste", _
                                           "Simulação de Reajuste", _
                                           0, _
                                           "", _
                                           "", _
                                           "", _
                                           "", _
                                           "P", _
                                           False, _
                                           vsmensagemerro, _
                                           vcContainer)

    Set vcContainer = Nothing

    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!" + Chr(13) + _
                    "As planilhas de simulação serão enviadas para o e-mail do usuário!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsmensagemerro, "I")
    End If
  End If

  Set interface = Nothing

  If VisibleMode Then
    CurrentQuery.Active = False
    CurrentQuery.Active = True
  End If
End Sub

Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
  '  If Len(EVENTO.Text)=0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTOINICIAL.Text)
  If vHandle <>0 Then
    If CurrentQuery.State = 1 Then
      CurrentQuery.Edit
    End If
    CurrentQuery.FieldByName("EVENTOINICIAL").Value = vHandle
  End If
  '  End If
End Sub

Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)
  '  If Len(EVENTO.Text)=0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTOFINAL.Text)
  If vHandle <>0 Then
    If CurrentQuery.State = 1 Then
      CurrentQuery.Edit
    End If
    CurrentQuery.FieldByName("EVENTOFINAL").Value = vHandle
  End If
  '  End If
End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  ShowPopup = False
  Dim vHandle As Long
  Dim qVerifica As Object
  Dim Interface As Object
  Dim vColunas As String
  Dim vCriterio As String
  Dim vCampos As String

  Set qVerifica = NewQuery

  vCriterio = ""
  Set Interface = CreateBennerObject("Procura.Procurar")
  vColunas = "CPFCNPJ|Z_NOME"
  vCampos = "CPFCNPJ|Nome"
  vHandle = Interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 1, vCampos, vCriterio, "Prestador", True, "")
  CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  Set Interface = Nothing
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If Not CheckEventosFx Then
    CanContinue = False
    Exit Sub
  End If

  Dim SQLTGE As Object

  Set SQLTGE = NewQuery
  SQLTGE.Add("SELECT T1.ESTRUTURANUMERICA ESTRUTURAINICIAL,")
  SQLTGE.Add("       T2.ESTRUTURANUMERICA ESTRUTURAFINAL")
  SQLTGE.Add("  FROM SAM_TGE  T1,")
  SQLTGE.Add("       SAM_TGE  T2 ")
  SQLTGE.Add(" WHERE T1.HANDLE = :HADLEINICIAL")
  SQLTGE.Add("   AND T2.HANDLE = :HANDLEFINAL ")
  SQLTGE.ParamByName("HADLEINICIAL").AsInteger = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
  SQLTGE.ParamByName("HANDLEFINAL").AsInteger  = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
  SQLTGE.Active = True
  If SQLTGE.FieldByName("ESTRUTURAINICIAL").IsNull Then
    bsShowMessage("A estrutura numérica do evento inicial está nula, verifique o cadastro!", "E")
    CanContinue =False
    Exit Sub
  End If
  If SQLTGE.FieldByName("ESTRUTURAFINAL").IsNull Then
    bsShowMessage("A estrutura numérica do evento final está nula, verifique o cadastro!", "E")
    CanContinue =False
    Exit Sub
  End If

  Dim EstruturaI As String
  Dim EstruturaF As String

  ' Atribuir ESTRUTURAINICIAL E FINAL
  Dim Estrutura As String

  ' Atribuir ESTRUTURAINICIAL
  SQLTGE.Clear
  SQLTGE.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTO")
  SQLTGE.ParamByName("HEVENTO").AsInteger = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
  SQLTGE.Active = True
  EstruturaI = SQLTGE.FieldByName("ESTRUTURA").Value
  CurrentQuery.FieldByName("ESTRUTURAINICIAL").Value = EstruturaI

  ' Atribuir ESTRUTURAFINAL
  SQLTGE.Active = False
  SQLTGE.ParamByName("HEVENTO").AsInteger = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
  SQLTGE.Active = True
  EstruturaF = SQLTGE.FieldByName("ESTRUTURA").Value
  CurrentQuery.FieldByName("ESTRUTURAFINAL").Value = EstruturaF
  SQLTGE.Active = False
  Set SQLTGE = Nothing

  CurrentQuery.FieldByName("USUARIO").Value = CurrentUser

End Sub


Public Function CheckEventosFx As Boolean
  CheckEventosFx = True
  If Not CurrentQuery.FieldByName("EVENTOINICIAL").IsNull Then
    If CurrentQuery.FieldByName("EVENTOFINAL").IsNull Then
      CurrentQuery.FieldByName("EVENTOFINAL").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
    Else
      If CurrentQuery.FieldByName("EVENTOINICIAL").Value <>CurrentQuery.FieldByName("EVENTOFINAL").Value Then
        Dim SQLI, SQLF As Object
        Set SQLI = NewQuery
        SQLI.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTOI")
        SQLI.ParamByName("HEVENTOI").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
        SQLI.Active = True

        Set SQLF = NewQuery
        SQLF.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTOF")
        SQLF.ParamByName("HEVENTOF").Value = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
        SQLF.Active = True

        If SQLF.FieldByName("ESTRUTURA").Value <SQLI.FieldByName("ESTRUTURA").Value Then
          bsShowMessage("Evento final não pode ser menor que o evento inicial!", "E")
          CheckEventosFx = False
        End If
        Set SQLI = Nothing
        Set SQLF = Nothing
      End If
    End If
  End If
End Function

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
    ' Ricardo Matiello - SMS 90364
	If (CommandID = "BOTAOPLANILHA") Then
		BOTAOPLANILHA_OnClick
	End If
End Sub
