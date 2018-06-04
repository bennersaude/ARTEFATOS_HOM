'HASH: 889AF33AA626145F865D47D5CA90E9C9
Option Explicit

'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"


Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
  '  If Len(EVENTO.Text)=0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(False, EVENTOINICIAL.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOINICIAL").Value = vHandle
  End If
  '  End If
End Sub


Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)
  '  If Len(EVENTO.Text)=0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(False, EVENTOFINAL.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOFINAL").Value = vHandle
  End If
  '  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String
  Dim EstruturaI As String
  Dim EstruturaF As String

  ' Atribuir ESTRUTURAINICIAL E FINAL
  Dim SQLTGE, SQLMASC As Object
  Dim Estrutura As String

  ' Atribuir ESTRUTURAINICIAL
  Set SQLTGE = NewQuery
  SQLTGE.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTO")
  SQLTGE.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
  SQLTGE.Active = True
  CurrentQuery.FieldByName("ESTRUTURAINICIAL").Value = SQLTGE.FieldByName("ESTRUTURA").Value

  ' Atribuir ESTRUTURAFINAL
  SQLTGE.Active = False
  SQLTGE.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
  SQLTGE.Active = True
  Estrutura = SQLTGE.FieldByName("ESTRUTURA").Value
  SQLTGE.Active = False
  Set SQLTGE = Nothing

  ' Completar ESTRUTURAFinal com 99999
  Set SQLMASC = NewQuery
  SQLMASC.Add("SELECT M.MASCARA MASCARA FROM Z_TABELAS T, Z_MASCARAS M")
  SQLMASC.Add("WHERE T.NOME = 'SAM_TGE' AND M.TABELA = T.HANDLE")
  SQLMASC.Active = True
  Estrutura = Estrutura + Mid(SQLMASC.FieldByName("MASCARA").AsString, Len(Estrutura) + 1)
  CurrentQuery.FieldByName("ESTRUTURAFINAL").Value = Estrutura
  SQLMASC.Active = False
  Set SQLMASC = Nothing



  ' Checar Vigencia
  EstruturaI = CurrentQuery.FieldByName("ESTRUTURAINICIAL").AsString
  EstruturaF = CurrentQuery.FieldByName("ESTRUTURAFINAL").AsString

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")


  If CurrentQuery.FieldByName("ISSMUNICIPIO").IsNull Then
    Condicao = Condicao + " ISSMUNICIPIO IS NULL"
  Else
    Condicao = Condicao + " ISSMUNICIPIO = " + CurrentQuery.FieldByName("ISSMUNICIPIO").AsString
  End If

  If Not CurrentQuery.FieldByName("TIPOPRESTADOR").IsNull Then
    Linha = Interface.EventoFx(CurrentSystem, "SFN_ISS_MUNICIPIO_REDUCAOFX", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, _
            CurrentQuery.FieldByName("DATAFINAL").AsDateTime, EstruturaI, EstruturaF, Condicao, _
            CurrentQuery.FieldByName("TIPOPRESTADOR").AsInteger)
  Else
    Linha = Interface.EventoFx(CurrentSystem, "SFN_ISS_MUNICIPIO_REDUCAOFX", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, _
            CurrentQuery.FieldByName("DATAFINAL").AsDateTime, EstruturaI, EstruturaF, Condicao, _
            0)
  End If

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
    Exit Sub
  End If
  Set Interface = Nothing
End Sub

