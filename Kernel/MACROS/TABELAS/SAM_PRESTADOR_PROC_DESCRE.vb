'HASH: 0187B055D4157256D0EB06276C457AF5
'Macro: SAM_PRESTADOR_PROC_DESCRE

'#Uses "*bsShowMessage"

Dim EstadoDaTabela As Long
Dim Mensagem As String

Public Function Ok As Boolean
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT DATAFINAL,RESPONSAVEL FROM SAM_PRESTADOR_PROC WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC")
  SQL.Active = True
  Ok = IIf(SQL.FieldByName("DATAFINAL").IsNull And SQL.FieldByName("RESPONSAVEL").AsInteger = CurrentUser, True, False)
  If Not SQL.FieldByName("DATAFINAL").IsNull Then
    Mensagem = "Processo finalizado! Operação não permitida." + Chr(13)
  End If
  If SQL.FieldByName("RESPONSAVEL").AsInteger <> CurrentUser Then
    Mensagem = Mensagem + "Usuário não é o responsável!"
  End If
  Set SQL = Nothing
End Function

Public Sub TABLE_AfterInsert()
  If Not Ok Then
    RefreshNodesWithTable "SAM_PRESTADOR_PROC"
    bsShowMessage(Mensagem, "E")
    CurrentQuery.Cancel
    RefreshNodesWithTable "SAM_PRESTADOR_PROC_HABIL"
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  EstadoDaTabela = CurrentQuery.State

  If CurrentQuery.FieldByName("TABPARECERSEDE").AsInteger = 1 Then
    CurrentQuery.FieldByName("DATADESCRESEDE").Clear
    CurrentQuery.FieldByName("MOTIVODESCRESEDE").Clear
    CurrentQuery.FieldByName("NOVACATEGORIASEDE").Clear
  Else
    ' Para parecer da sede, nao permitir gravar data de descredenciamento menor do que credenciamento
    If CurrentQuery.FieldByName("TABPARECERSEDE").AsInteger = 2 Then
      CurrentQuery.FieldByName("REVISAREMSEDE").Clear
      Dim SQL As Object
      Set SQL = NewQuery
      SQL.Add("SELECT DATACREDENCIAMENTO FROM SAM_PRESTADOR WHERE HANDLE = :PRESTADOR")
      SQL.ParamByName("PRESTADOR").Value = RecordHandleOfTable("SAM_PRESTADOR")
      SQL.Active = True

      If SQL.FieldByName("datacredenciamento").Value > CurrentQuery.FieldByName("DATADESCRESEDE").Value Then
        bsShowMessage("Data de descredenciamento (" + CurrentQuery.FieldByName("DATADESCRESEDE").Value + _
        ") menor do que a data de credenciamento (" + SQL.FieldByName("datacredenciamento").Value + ").", "E")
        CanContinue = False
        Exit Sub
      End If
    End If
  End If

  If CurrentQuery.FieldByName("TABPARECERREPRE").AsInteger = 1 Then
    CurrentQuery.FieldByName("MOTIVODESCREREPRE").Clear
    CurrentQuery.FieldByName("DATADESCREREPRE").Clear
  Else
    If CurrentQuery.FieldByName("TABPARECERREPRE").AsInteger = 2 Then
      If CurrentQuery.FieldByName("MOTIVODESCREREPRE").IsNull Then
        bsShowMessage("Motivo do parecer da filial é obrigatório.", "E")
        CanContinue = False
        Exit Sub
      End If
      If CurrentQuery.FieldByName("DATADESCREREPRE").IsNull Then
        bsShowMessage("Data do parecer da filial é obrigatório.", "E")
        CanContinue = False
        Exit Sub
      End If
      CurrentQuery.FieldByName("REVISAREMREPRE").Clear
    End If
  End If


  Dim s As Object
  Dim m As String

  Set s = NewQuery
  s.Add("SELECT T.EXIGEPARECERREPRESENTACAO, T.EXIGEPARECERSEDE")
  s.Add("  FROM SAM_PRESTADOR P,")
  s.Add("       SAM_TIPOPRESTADOR T")
  s.Add(" WHERE P.HANDLE = :HANDLE")
  s.Add("   AND T.HANDLE = P.TIPOPRESTADOR")
  s.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR")
  s.Active = True

  If s.FieldByName("EXIGEPARECERSEDE").AsString = "S" Then
    If CurrentQuery.FieldByName("TABPARECERSEDE").IsNull Then m = "Tipo do prestador exige parecer da sede"
    If CurrentQuery.FieldByName("JUSTIFICATIVASEDE").IsNull Then m = m + Chr(13) + "Tipo do prestador exige justificativa da sede"
    If m<>"" Then
      CanContinue = False
      bsShowMessage(m, "E")
      Exit Sub
    End If
  End If

  If s.FieldByName("EXIGEPARECERREPRESENTACAO").AsString = "S" Then
    If CurrentQuery.FieldByName("TABPARECERREPRE").IsNull Then m = "Tipo do prestador exige parecer da filial"
    If CurrentQuery.FieldByName("JUSTIFICATIVAREPRE" ).IsNull Then m = m + Chr(13) + "Tipo do prestador exige justificativa da filial"
    If m<>"" Then
      CanContinue = False
      bsShowMessage(m, "E")
      Exit Sub
    End If
  End If

  RefreshNodesWithTable("SAM_PRESTADOR_PROC_DESCRE")

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = Ok
  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = Ok
  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub


Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOVAGAS" Then
		BOTAOVAGAS_OnClick
	End If
End Sub

Public Sub TABPARECERREPRE_OnChanging(AllowChange As Boolean)
  AllowChange = Ok
  If Not AllowChange Then
    MsgBox Mensagem
  End If
End Sub

Public Sub TABPARECERSEDE_OnChanging(AllowChange As Boolean)
  AllowChange = Ok
  If Not AllowChange Then
    MsgBox Mensagem
  End If
End Sub


Public Sub BOTAOVAGAS_OnClick()
  Dim Interface As Object

  Set Interface = CreateBennerObject("SamVaga.Atendimento")
  Interface.Consultarvaga(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Set Interface = Nothing
End Sub

