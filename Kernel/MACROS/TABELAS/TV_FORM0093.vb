'HASH: 0C0CDF96244D29869DDD0A5E316F586A
'#Uses "*bsShowMessage"

Option Explicit

Dim vgNovaAcomodacao       As Long
Dim vgNovoTipoAcomodacao   As Long
Dim vgEvento               As String
Dim vgGrau                 As String
Dim vgTipoAcomodacao       As String
Dim vgDiariaTipoAcomodacao As String
Dim vgHandleAutoriz        As Long
Dim vgHandleProtocolo      As Long

Public Sub TABLE_AfterScroll()

  GRUPONOVAACOMODACAO.Visible = False
  ACOMODACAO.Visible = False
  TIPOACOMODACAO.Visible = False
  EVENTO.Visible = False
  GRAU.Visible = False

End Sub

Public Sub TABLE_NewRecord()

  Dim qParametros As Object
  Set qParametros = NewQuery

  qParametros.Clear
  qParametros.Add(" Select P.TIPOACOMODACAO,            ")
  qParametros.Add("        P.DIARIATIPOACOMODACAOTISS   ")
  qParametros.Add("    FROM SAM_PARAMETROSATENDIMENTO P ")
  qParametros.Active = True

  vgTipoAcomodacao = qParametros.FieldByName("TIPOACOMODACAO").AsString
  vgDiariaTipoAcomodacao = qParametros.FieldByName("DIARIATIPOACOMODACAOTISS").AsString

  Set qParametros = Nothing

  SessionVar("TIPOACOMODACAOAUTORIZ") = vgTipoAcomodacao

  If (SessionVar("HANDLEAUTORIZACAO") <> "") Then
    vgHandleAutoriz = CLng(SessionVar("HANDLEAUTORIZACAO"))
  Else
    vgHandleAutoriz = 0
  End If

  If (SessionVar("PROTOCTRANSAUTOR") <> "") Then
    vgHandleProtocolo = CLng(SessionVar("PROTOCTRANSAUTOR"))
  Else
    vgHandleProtocolo = 0
  End If

  If VisibleMode Then
    CurrentQuery.FieldByName("ACOMODACAOATUAL").AsInteger = CLng(SessionVar("ACOMODACAOATUAL"))
  Else
    If (vgHandleAutoriz <> 0) Then
      Dim qAcomodacao As Object
      Set qAcomodacao = NewQuery

      qAcomodacao.Clear
      qAcomodacao.Add("  SELECT ACOMODACAO       ")
      qAcomodacao.Add("   FROM SAM_AUTORIZ       ")
      qAcomodacao.Add("  WHERE HANDLE = :HANDLE  ")

      qAcomodacao.ParamByName("HANDLE").AsInteger = vgHandleAutoriz
      qAcomodacao.Active = True

      CurrentQuery.FieldByName("ACOMODACAOATUAL").AsInteger = qAcomodacao.FieldByName("ACOMODACAO").AsInteger

      Set qAcomodacao = Nothing
    Else
      CurrentQuery.FieldByName("ACOMODACAOATUAL").Clear
    End If
  End If

End Sub

Public Sub NOVAACOMODACAO_OnPopup(ShowPopup As Boolean)

  ShowPopup = False

  Dim interface As Object
  Dim vHandle As String
  Dim vTemDiariaProrrogacao As Boolean

  Dim q1 As Object
  Set q1 = NewQuery

  q1.Clear
  q1.Add("  Select COUNT(*) QTD             ")
  q1.Add("  FROM SAM_AUTORIZ_EVENTOGERADO   ")
  If vgHandleProtocolo > 0 Then
  	q1.Add(" WHERE PROTOCOLOTRANSACAO = :HANDLE ")
  Else
  	q1.Add(" WHERE AUTORIZACAO = :HANDLE ")
  End If
  q1.Add("   And SITUACAO <> 'C'            ")
  q1.Add("   And TIPOEVENTO IN ('D', 'P')   ")
  If vgHandleProtocolo > 0 Then
	q1.ParamByName("HANDLE").AsInteger = vgHandleProtocolo
  Else
  	q1.ParamByName("HANDLE").AsInteger = vgHandleAutoriz
  End If
  q1.Active = True

  If q1.FieldByName("QTD").AsInteger > 0 Then
    vTemDiariaProrrogacao = True
  Else
    vTemDiariaProrrogacao = False
  End If

  Set q1 = Nothing

  Dim qAutoriz As Object
  Set qAutoriz = NewQuery

  If (vgHandleAutoriz > 0) Then
	  qAutoriz.Add(" Select A.HANDLE,                                                 ")
	  qAutoriz.Add("        ES.EVENTO,                                                ")
	  qAutoriz.Add("        ES.GRAU,                                                  ")
	  qAutoriz.Add("        ES.EXECUTOR,                                              ")
	  qAutoriz.Add("        ES.RECEBEDOR,                                             ")
	  qAutoriz.Add("        ES.DATAATENDIMENTO,                                       ")
	  qAutoriz.Add("        A.BENEFICIARIO,                                           ")
	  qAutoriz.Add("        A.ACIDENTEPESSOAL,                                        ")
	  qAutoriz.Add("        A.FINALIDADEATENDIMENTO,                                  ")
	  qAutoriz.Add("        A.CONDICAOATENDIMENTO,                                    ")
	  qAutoriz.Add("        A.LOCALATENDIMENTO,                                       ")
	  qAutoriz.Add("        A.OBJETIVOTRATAMENTO,                                     ")
	  qAutoriz.Add("        A.TIPOTRATAMENTO,                                         ")
	  qAutoriz.Add("        A.REGIMEATENDIMENTO,                                      ")
	  qAutoriz.Add("        A.TIPOACOMODACAO                                          ")
	  qAutoriz.Add("   FROM SAM_AUTORIZ A                                             ")
	  qAutoriz.Add("   Join SAM_AUTORIZ_EVENTOSOLICIT ES On ES.AUTORIZACAO = A.Handle ")
	  qAutoriz.Add("  WHERE A.HANDLE = :HANDLE                                        ")
	  qAutoriz.ParamByName("HANDLE").AsInteger = vgHandleAutoriz
	  qAutoriz.Active = True


  Set interface = CreateBennerObject("SamAuto.Rotinas")

  SessionVar("TROCAACOMODACAO") = "1"
  vHandle = interface.ProcuraMostra(CurrentSystem, _
                           qAutoriz.FieldByName("HANDLE").AsInteger, _
                           qAutoriz.FieldByName("EVENTO").AsInteger, _
                           qAutoriz.FieldByName("GRAU").AsInteger, _
                           qAutoriz.FieldByName("EXECUTOR").AsInteger, _
                           qAutoriz.FieldByName("RECEBEDOR").AsInteger, _
                           qAutoriz.FieldByName("DATAATENDIMENTO").AsDateTime, _
                           False, _
                           SessionVar("TipoEvento"), _
                           qAutoriz.FieldByName("BENEFICIARIO").AsInteger, _
    					   qAutoriz.FieldByName("ACIDENTEPESSOAL").AsString, _
    					   qAutoriz.FieldByName("FINALIDADEATENDIMENTO").AsInteger, _
    					   qAutoriz.FieldByName("CONDICAOATENDIMENTO").AsInteger, _
    					   qAutoriz.FieldByName("LOCALATENDIMENTO").AsInteger, _
    					   qAutoriz.FieldByName("OBJETIVOTRATAMENTO").AsInteger, _
    					   qAutoriz.FieldByName("TIPOTRATAMENTO").AsInteger, _
    					   qAutoriz.FieldByName("REGIMEATENDIMENTO").AsInteger, _
    					   qAutoriz.FieldByName("TIPOACOMODACAO").AsInteger, _
    					   vTemDiariaProrrogacao)
  End If

  If vHandle <> "" Then
    BuscaRegistro(vHandle)

  End If

SessionVar("TROCAACOMODACAO") = ""
Set interface = Nothing
Set qAutoriz = Nothing

End Sub

Public Sub BuscaRegistro(vAcomodacao As String)

  Dim qAcomodacao As Object
  Set qAcomodacao = NewQuery

  qAcomodacao.Clear

  If vgTipoAcomodacao = "G" Then
    qAcomodacao.Add(" SELECT A.Handle,									          ")
    qAcomodacao.Add(" AE.GRAU EVENTOGRAU                                          ")
    qAcomodacao.Add("   FROM SAM_ACOMODACAO A								      ")
    qAcomodacao.Add("  JOIN SAM_ACOMODACAO_GRAU AE ON AE.ACOMODACAO = A.HANDLE    ")
  Else
    qAcomodacao.Add(" Select A.Handle,									          ")
    qAcomodacao.Add(" AE.EVENTO EVENTOGRAU,                                       ")
    qAcomodacao.Add(" AE.GRAUAGERAR                                               ")
    qAcomodacao.Add("   FROM SAM_ACOMODACAO A									  ")
    qAcomodacao.Add("  JOIN SAM_ACOMODACAO_EVENTO AE ON AE.ACOMODACAO = A.HANDLE  ")
  End If

  qAcomodacao.Add("  WHERE AE.Handle = :HANDLE                                    ")
  qAcomodacao.ParamByName("HANDLE").AsInteger = CLng(vAcomodacao)
  qAcomodacao.Active = True

  CurrentQuery.Edit
  CurrentQuery.FieldByName("NOVAACOMODACAO").AsInteger = qAcomodacao.FieldByName("HANDLE").AsInteger

  vgNovaAcomodacao = qAcomodacao.FieldByName("HANDLE").AsInteger

  vgEvento = ""
  vgGrau = ""

  If vgTipoAcomodacao = "G" Then
    vgGrau = CStr(qAcomodacao.FieldByName("EVENTOGRAU").AsInteger)
  Else
    vgEvento = CStr(qAcomodacao.FieldByName("EVENTOGRAU").AsInteger)
    vgGrau = CStr(qAcomodacao.FieldByName("GRAUAGERAR").AsInteger)

    SessionVar("HANDLEVENTOGRAU") = vAcomodacao
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If VisibleMode Then

    If Not CurrentQuery.FieldByName("NOVAACOMODACAO").IsNull Then
      SessionVar("NOVAACOMODACAO") = CStr(vgNovaAcomodacao)
      SessionVar("EVENTOACOMODACAO") = vgEvento
      SessionVar("GRAUACOMODACAO") = vgGrau
    Else
      bsShowMessage("Campo Nova Acomodação obrigatório!", "E")
      CanContinue = False
      Exit Sub
    End If

  Else

    Dim vsRetorno As String
    Dim vsGerar As String

    If vgDiariaTipoAcomodacao = "S" Then
      If CurrentQuery.FieldByName("TIPOACOMODACAO").AsString <> "" Then
        SessionVar("NOVOTIPOACOMODACAO") = CurrentQuery.FieldByName("TIPOACOMODACAO").AsString
      Else
        bsShowMessage("Campo Tipo Acomodação obrigatório!", "E")
        CanContinue = False
        Exit Sub
      End If
    End If

    If vgTipoAcomodacao = "G" Then
      vsGerar = CurrentQuery.FieldByName("GRAU").AsString
    Else
      vsGerar = CurrentQuery.FieldByName("EVENTO").AsString
    End If

	If vsGerar <> "" Then
	  BuscaRegistro(vsGerar)

      SessionVar("NOVAACOMODACAO") = CStr(vgNovaAcomodacao)
      SessionVar("EVENTOACOMODACAO") = vgEvento
      SessionVar("GRAUACOMODACAO") = vgGrau

      Dim dllTrocarAcomodacao   As Object
      Set dllTrocarAcomodacao = CreateBennerObject("CA043.AUTORIZACAO")

      vsRetorno = dllTrocarAcomodacao.TrocarAcomodacao(CurrentSystem, vgHandleAutoriz, vgHandleProtocolo)

      If vsRetorno <> "" Then
        bsShowMessage(vsRetorno, "E")
        CanContinue = False

        Set dllTrocarAcomodacao = Nothing
        Exit Sub
      End If

      Set dllTrocarAcomodacao = Nothing
	Else
	  bsShowMessage("Campo Gerar obrigatório!","E")
	  CanContinue = False
	  Exit Sub
	End If

  End If
End Sub

Public Sub TABLE_BeforeCancel(CanContinue As Boolean)

  SessionVar("NOVAACOMODACAO") = ""
  SessionVar("EVENTOACOMODACAO") = ""
  SessionVar("TIPOACOMODACAOAUTORIZ") = ""
  SessionVar("GRAUACOMODACAO") = ""
  SessionVar("HANDLEVENTOGRAU") = ""
End Sub
