'HASH: 5D381D6B558B785CD792A13A859437E0
'Macro: SAM_PRESTADOR_TPSERVICO
'Mauricio Ibelli -14/08/2001 -sms3858 -colocar na condicao a classificacao
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
  Condicao = " AND TIPOSERVICO = " + CurrentQuery.FieldByName("TIPOSERVICO").AsString
  Condicao = Condicao + " AND classificacao = " + "'" + CurrentQuery.FieldByName("classificacao").AsString + "'"

  If VisibleMode Then
    Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_TPSERVICO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", Condicao)
    If Linha = "" Then
      CanContinue = True
    Else
      CanContinue = False
      bsShowMessage(Linha, "E")
    End If
  End If

  Set Interface = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
      bsShowMessage(Msg, "E")
      CanContinue = False
      Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
      bsShowMessage(Msg, "E")
      CanContinue = False
      Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
      bsShowMessage(Msg, "E")
      CanContinue = False
      Exit Sub
  End If
End Sub

