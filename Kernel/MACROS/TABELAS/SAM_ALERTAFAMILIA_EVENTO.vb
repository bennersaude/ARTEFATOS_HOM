'HASH: 149D17631FFF16FCC07575F7AEB5CAAE
'Macro: SAM_ALERTAFAMILIA_EVENTO
Option Explicit

'#Uses "*ProcuraEvento"
'#Uses "*ProcuraGrau"
'#Uses "*bsShowMessage"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  '  If Len(EVENTO.Text) = 0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTO.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
  '  End If
End Sub

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
  ''  If Len(GRAU.Text) = 0 Then
  '      Dim vHandle As Long
  '      ShowPopup = False
  '      vHandle = ProcuraGrau
  '      If vHandle<>0 Then
  '         CurrentQuery.Edit
  '         CurrentQuery.FieldByName("GRAU").Value=vHandle
  '      End If
  ''  End If

  'Inicio SMS 37619 Wagner Santos 20/07/2005
  ShowPopup = False
  Dim vHandle As Long
  Dim vColunas As String
  Dim vCriterio As String
  Dim vCampos As String
  Dim ProcuraDll As Object
  Dim SQL As Object
  Dim vGrau As String
  Dim vContador As Integer
  Dim vColunaIndice As Integer

  vCriterio = ""
  vColunaIndice = 1
  vGrau = TiraAcento(GRAU.Text, True)

  For vContador = 1 To Len(GRAU.Text)
    If InStr(GRAU.Text, "0") Or _
              InStr(GRAU.Text, "1") Or _
              InStr(GRAU.Text, "2") Or _
              InStr(GRAU.Text, "3") Or _
              InStr(GRAU.Text, "4") Or _
              InStr(GRAU.Text, "5") Or _
              InStr(GRAU.Text, "6") Or _
              InStr(GRAU.Text, "7") Or _
              InStr(GRAU.Text, "8") Or _
              InStr(GRAU.Text, "9") Then
      vColunaIndice = 1
    Else
      vColunaIndice = 2
      Exit For
    End If
  Next

  Set SQL = NewQuery
  SQL.Add("SELECT GRAUSVALIDOSNOALERTA FROM SAM_PARAMETROSPRESTADOR")
  SQL.Active = True

  Set ProcuraDll = CreateBennerObject("PROCURA.PROCURAR")
  vColunas = "SAM_GRAU.GRAU|SAM_GRAU.DESCRICAO|SAM_TIPOGRAU.DESCRICAO|SAM_GRAU.VERIFICAGRAUSVALIDOS"

  If GRAU.Text = "" Then
    If Not CurrentQuery.FieldByName("EVENTO").IsNull Then
      If SQL.FieldByName("GRAUSVALIDOSNOALERTA").AsString = "S" Then
        vCriterio = "SAM_GRAU.HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString + ")"
      Else
        vCriterio = ""
      End If
    Else
      bsShowMessage("Informar evento!", "I")
      EVENTO.SetFocus
      Exit Sub
    End If
  Else
    If vColunaIndice = 1 Then
      vCriterio = "GRAU = " + vGrau
    Else
      vCriterio = "Z_DESCRICAO LIKE '" + vGrau + "%'"
      vGrau = ""
    End If
  End If

  vCampos = "Código do Grau|Descrição|Tipo do Grau|Graus Válidos"
  vHandle = ProcuraDll.Exec(CurrentSystem, "SAM_GRAU|SAM_TIPOGRAU[SAM_TIPOGRAU.HANDLE = SAM_GRAU.TIPOGRAU]", vColunas, vColunaIndice, vCampos, vCriterio, "Graus de atuação", True, vGrau)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAU").Value = vHandle
  End If
  Set ProcuraDll = Nothing
  Set SQL = Nothing
  'Fim SMS 37619

End Sub

Public Sub TABLE_AfterScroll()
  If WebMode Then
  	Dim SQL As Object

	Set SQL = NewQuery
  	SQL.Add("SELECT GRAUSVALIDOSNOALERTA FROM SAM_PARAMETROSPRESTADOR")
  	SQL.Active = True

	If SQL.FieldByName("GRAUSVALIDOSNOALERTA").AsString = "S" Then
  		GRAU.WebLocalWhere = "A.HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = @CAMPO(EVENTO))"
	Else
  		GRAU.WebLocalWhere = ""
	End If
  End If

  Dim Q As Object

  Set Q = NewQuery
  Q.Add("SELECT * FROM SAM_ALERTAFAMILIA WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("FAMILIAALERTA").AsInteger
  Q.Active = True

  CurrentQuery.RequestLive = Q.FieldByName("DATAFINAL").IsNull
  EVENTO .ReadOnly = Not Q.FieldByName("DATAFINAL").IsNull
  GRAU .ReadOnly = Not Q.FieldByName("DATAFINAL").IsNull

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim q1 As Object
  Set q1 = NewQuery
  q1.Add("SELECT USUARIO FROM SAM_ALERTAFAMILIA WHERE HANDLE=:HANDLE")
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("FAMILIAALERTA").AsInteger
  q1.Active = True
  If q1.FieldByName("USUARIO").AsInteger<>CurrentUser Then
    bsShowMessage("O usuário atual não tem permissão para alteração", "E")
    CanContinue = False
  End If
  q1.Active = False
  Set q1 = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim q1 As Object
  Set q1 = NewQuery
  q1.Add("SELECT USUARIO FROM SAM_ALERTAFAMILIA WHERE HANDLE=:HANDLE")
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("FAMILIAALERTA").AsInteger
  q1.Active = True
  If q1.FieldByName("USUARIO").AsInteger<>CurrentUser Then
    bsShowMessage("O usuário atual não tem permissão para alteração", "E")
    CanContinue = False
  End If
  q1.Active = False
  Set q1 = Nothing


  '******************************************************************************************************
  Dim Q As Object

  Set Q = NewQuery
  Q.Add("SELECT * FROM SAM_ALERTAFAMILIA WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("FAMILIAALERTA").AsInteger
  Q.Active = True

  If Not Q.FieldByName("DATAFINAL").IsNull Then
    bsShowMessage("Cadastro não permitido, pois a vigência do alerta está fechada", "E")
    CanContinue = False
    Exit Sub
  End If
  '******************************************************************************************************


  '**********************  Durval 07/05/2003 *************************************************88
  Q.Clear
  Q.Add("SELECT * FROM SAM_ALERTAFAMILIA_EVENTO")
  Q.Add(" WHERE (EVENTO = :EVENTO )")
  Q.Add("   AND (HANDLE <> :HANDLE)")
  If Not CurrentQuery.FieldByName("grau").IsNull Then
    Q.Add("   AND (GRAU = :GRAU OR GRAU IS NULL)")
    Q.ParamByName("GRAU").Value = CurrentQuery.FieldByName("GRAU").AsInteger
  End If
  Q.Add("   AND (FAMILIAALERTA = :FAMILIAALERTA)")
  Q.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTO").AsInteger
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Q.ParamByName("FAMILIAALERTA").Value = CurrentQuery.FieldByName("FAMILIAALERTA").AsInteger
  Q.Active = True

  If Not Q.EOF Then
    bsShowMessage("Já existe este evento cadastrado para este alerta", "E")
    CanContinue = False
    Exit Sub
  End If
  Set Q = Nothing
  '**********************  Durval 07/05/2003 *************************************************88




End Sub

