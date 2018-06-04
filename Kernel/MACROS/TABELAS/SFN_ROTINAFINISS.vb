'HASH: B5F75D1F28F78F223BC797711B12C98C
'MACRO SFN_ROTINAFINISS

'#Uses "*bsShowMessage"
'#Uses "*ProcuraPrestador"


Option Explicit

Public Sub BOTAOCANCELAR_OnClick()
  If bsShowMessage("Confirma o cancelamento da rotina ?", "Q") = vbYes Then
    Dim Interface As Object
    Dim viRetorno As Integer
    Dim vsMensagem As String

    If CurrentQuery.FieldByName("SITUACAO").AsString <>"5" Then 'Verifica se a rotina não foi processada
      bsShowMessage("A rotina não foi processada !", "I")
      Exit Sub
    End If

    If VisibleMode Then
       Set Interface = CreateBennerObject("BSINTERFACE0042.RotinaFinISS")
       Interface.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    Else
      Dim qSQL As Object
      Set qSQL = NewQuery

      qSQL.Clear
      qSQL.Add("SELECT SFAT.DESCRICAO DESCRICAOTIPOFATURAMENTO,")
      qSQL.Add("       CFIN.COMPETENCIA,")
      qSQL.Add("       RFIN.SEQUENCIA")
      qSQL.Add("FROM SFN_ROTINAFIN       RFIN")
      qSQL.Add("JOIN SFN_COMPETFIN       CFIN ON RFIN.COMPETFIN       = CFIN.HANDLE")
      qSQL.Add("JOIN SIS_TIPOFATURAMENTO SFAT ON CFIN.TIPOFATURAMENTO = SFAT.HANDLE")
      qSQL.Add("WHERE RFIN.HANDLE = :HROTINAFIN")
      qSQL.ParamByName("HROTINAFIN").AsInteger = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
      qSQL.Active = True

      Set Interface = CreateBennerObject("BSServerExec.ProcessosServidor")
      viRetorno = Interface.ExecucaoImediata(CurrentSystem, _
                                             "SfnRecolhimento", _
                                             "RotinaFinISS_CancelaISS", _
                                             "Rotina de Recolhimento de ISS (Cancelamento) -" + _
                                               " Competência: " + Str(Format(qSQL.FieldByName("COMPETENCIA").AsDateTime, "mm/yyyy")) + _
                                               " Sequência: "   + qSQL.FieldByName("SEQUENCIA").AsString, _
                                             CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                             "SFN_ROTINAFINISS", _
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

      Set qSQL = Nothing
    End If

    CurrentQuery.Active = False
    CurrentQuery.Active = True

    Set Interface = Nothing
  End If
End Sub

Public Sub BOTAOGERARARQUIVO_OnClick()
'SMS 90432 - Marcelo Barbosa - 17/04/2008
'Essa geracao de arquivo é especifica para a Camed e ISSFacil de Maringá (PAM), não será feito para a WEB
'  Dim SQL As Object
'  Set SQL = NewQuery

'  SQL.Active = False
'  SQL.Add("SELECT SITUACAO FROM SFN_ROTINAFIN WHERE HANDLE = :HROTINAFIN")
'  SQL.ParamByName("HROTINAFIN").AsInteger = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
'  SQL.Active = True

'  If SQL.FieldByName("SITUACAO").AsString <>"P" Then 'Verifica se a rotina não foi processada
'    MsgBox "A rotina não foi processada !", "Informação"
'    Set SQL = Nothing
'    Exit Sub
'  End If

'  Dim Interface As Object
'  Set Interface = CreateBennerObject("SfnRecolhimento.Rotinas")
'  Interface.GerarArquivoISS(CurrentSystem, CurrentQuery.FieldByName("ROTINAFIN").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
'  Set Interface = Nothing
'  Set SQL = Nothing
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim Interface As Object
  Dim viRetorno As Integer
  Dim vsMensagem As String

  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição", "I")
    Exit Sub
  End If


  If VisibleMode Then
     Set Interface = CreateBennerObject("BSINTERFACE0042.RotinaFinISS")
     Interface.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Else
    Dim qSQL As Object
    Set qSQL = NewQuery

    qSQL.Clear
    qSQL.Add("SELECT SFAT.DESCRICAO DESCRICAOTIPOFATURAMENTO,")
    qSQL.Add("       CFIN.COMPETENCIA,")
    qSQL.Add("       RFIN.SEQUENCIA")
    qSQL.Add("FROM SFN_ROTINAFIN       RFIN")
    qSQL.Add("JOIN SFN_COMPETFIN       CFIN ON RFIN.COMPETFIN       = CFIN.HANDLE")
    qSQL.Add("JOIN SIS_TIPOFATURAMENTO SFAT ON CFIN.TIPOFATURAMENTO = SFAT.HANDLE")
    qSQL.Add("WHERE RFIN.HANDLE = :HROTINAFIN")
    qSQL.ParamByName("HROTINAFIN").AsInteger = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
    qSQL.Active = True

    Set Interface = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Interface.ExecucaoImediata(CurrentSystem, _
                                          "SfnRecolhimento", _
                                          "RotinaFinISS_ProcessaISS", _
                                          "Rotina de Recolhimento de ISS (Processamento) -" + _
                                            " Competência: " + Str(Format(qSQL.FieldByName("COMPETENCIA").AsDateTime, "mm/yyyy")) + _
                                            " Sequência: "   + qSQL.FieldByName("SEQUENCIA").AsString, _
                                          CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                          "SFN_ROTINAFINISS", _
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

   Set qSQL = Nothing
 End If

 CurrentQuery.Active = False
 CurrentQuery.Active = True

 Set Interface = Nothing
End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim viHandlePrestador As Integer
  Dim vsCPFNome As String

  If (IsNumeric(PRESTADOR.Text)) Then
    vsCPFNome = "C"
  Else
    vsCPFNome = "N"
  End If

  viHandlePrestador = ProcuraPrestador(vsCPFNome, "T", PRESTADOR.Text)

  If viHandlePrestador <> 0 Then
       CurrentQuery.Edit
       CurrentQuery.FieldByName("PRESTADOR").Value = viHandlePrestador
  End If

End Sub

Public Sub TABLE_AfterScroll()
    BOTAOPROCESSAR.Enabled = CurrentQuery.FieldByName("SITUACAO").AsString = "1"
    BOTAOCANCELAR.Enabled = CurrentQuery.FieldByName("SITUACAO").AsString  = "5"
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If CurrentQuery.FieldByName("SITUACAO").AsString = "5" Then
    bsShowMessage("Não foi possível excluir, a rotina está processada !", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If CurrentQuery.FieldByName("SITUACAO").AsString = "5" Then
    bsShowMessage("Alteração negada, a rotina está processada !", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOGERARARQUIVO"
			BOTAOGERARARQUIVO_OnClick
		Case "BOTAOPROCESSAR"
			BOTAOPROCESSAR_OnClick
	End Select
End Sub
