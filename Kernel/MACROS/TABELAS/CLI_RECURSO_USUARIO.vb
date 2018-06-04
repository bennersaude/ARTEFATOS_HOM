'HASH: 388C0787602C04C5E7B47A62CF7FC113
'#Uses "*bsShowMessage"

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "PRESTADOR|NOME|SOLICITANTE|EXECUTOR|RECEBEDOR|NAOFATURARGUIAS"
  vCriterio = ""
  vCampos = "Prestador|Nome|Solicitante|Executor|Recebedor|Não faturar guias"
  vHandle = interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 2, vCampos, vCriterio, "Prestador", False, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
'SMS 46142 - 28/12/2005
  If CurrentQuery.FieldByName("PRESTADOR").IsNull Then
    If Not ((CurrentQuery.FieldByName("PRONTUARIOMEDICO").AsBoolean) Or (CurrentQuery.FieldByName("PRONTUARIOSOCIAL").AsBoolean) Or (CurrentQuery.FieldByName("PRONTUARIOODONTOLOGICO").AsBoolean) Or (CurrentQuery.FieldByName("PRONTUARIOENFERMAGEM").AsBoolean) Or (CurrentQuery.FieldByName("ACESSAATENDIMENTOS").AsBoolean)) Then
      bsShowMessage("Para salvar o recurso sem o prestador é necessário marcar pelo menos um dos demais campos", "E")
      CanContinue =False
    End If
  End If
'FIM SMS 46142
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT U.NOME")
  SQL.Add("  FROM CLI_RECURSO_USUARIO R,")
  SQL.Add("       Z_GRUPOUSUARIOS U")
  SQL.Add(" WHERE R.PRESTADOR = :PRESTADOR")
  SQL.Add("   AND R.USUARIO <> :USUARIO")
  SQL.Add("   AND R.USUARIO = U.HANDLE")
  SQL.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  SQL.ParamByName("USUARIO").AsInteger = CurrentQuery.FieldByName("USUARIO").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    bsShowMessage("O usuário " + SQL.FieldByName("NOME").AsString + " já foi cadastrado para esse prestador!", "E")
    CanContinue = False
  End If

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT COUNT(R.HANDLE) RECURSOS,                       ")
  SQL.Add("       U.NOME                                          ")
  SQL.Add("  FROM CLI_RECURSO_USUARIO R                           ")
  SQL.Add("  JOIN Z_GRUPOUSUARIOS     U ON (U.HANDLE = R.USUARIO) ")
  SQL.Add(" WHERE R.USUARIO = :USUARIO                            ")
  SQL.Add("   AND R.HANDLE <> :HANDLE                             ")
  SQL.Add(" GROUP BY U.NOME                                       ")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("USUARIO").AsInteger = CurrentQuery.FieldByName("USUARIO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("RECURSOS").AsInteger > 0 Then
    bsShowMessage("O usuário " + SQL.FieldByName("NOME").AsString + " já possui prestador vinculado!", "E")
    CanContinue = False
  End If

  Set SQL = Nothing
End Sub

