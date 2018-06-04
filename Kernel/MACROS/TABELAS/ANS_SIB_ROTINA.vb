'HASH: B41E41712A5F4DD83E5736715F08124D
'Macro: ANS_SIB_ROTINA
'#Uses "*bsShowMessage"

Option Explicit
Dim obj As Object

Public Sub BOTAOAJUSTARBENEFICIARIO_OnClick()

  SessionVar("HANDLEROTINASIB") = CurrentQuery.FieldByName("HANDLE").AsString

  If VisibleMode Then

    Dim qVerificaXml As BPesquisa
    Set qVerificaXml = NewQuery
    qVerificaXml.Add("SELECT A.HANDLE                                                ")
    qVerificaXml.Add("  FROM ANS_SIBPADRAO_XML A                                     ")
    qVerificaXml.Add(" WHERE (COMPETENCIA = :COMP                                    ")
    qVerificaXml.Add("   AND TABENVIORETORNO = 1                                     ")
    qVerificaXml.Add("   AND ROTINASIB = :ROTINA)				  					 ")
    qVerificaXml.Add(" ORDER BY A.COMPETENCIA, A.DATAHORA                            ")
    qVerificaXml.ParamByName("COMP").AsDateTime = CurrentQuery.FieldByName("COMPETENCIA").AsDateTime
    qVerificaXml.ParamByName("ROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qVerificaXml.Active = True

   If qVerificaXml.EOF Then
      Dim form As CSVirtualForm
      Set form = NewVirtualForm

      form.Caption = "Localizar Beneficiário"
      form.TableName = "TV_ANS002"
      form.Show

      Set form = Nothing
    Else
      BsShowMessage("Impossível realizar ajuste de beneficiário. XML já processado!", "I")
    End If

    qVerificaXml.Active = False
    Set qVerificaXml = Nothing

  End If

End Sub

Public Sub BOTAOCANCELAR_OnClick()
  Dim Interface As Object
  Dim vsMensagemRetorno As String
  Dim viRetorno As Long

  If CurrentQuery.State <>1 Then
	bsShowMessage("A tabela não pode estar em edição", "I")
	Exit Sub
  End If

  If VisibleMode Then
  	Set obj = CreateBennerObject("BSINTERFACE0056.Cancelar")
  	obj.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMensagemRetorno)
  	Set Interface = Nothing
  Else
	Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = obj.ExecucaoImediata(CurrentSystem, _
                                    "BSANS002", _
                                    "Cancelar", _
                                    "SIB - Sistema de informações de Beneficiarios - Cancelar", _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    "ANS_SIB_ROTINA", _
                                    "SITUACAOPROCESSO", _
                                    "", _
                                    "", _
                                    "C", _
                                    False, _
                                    vsMensagemRetorno, _
                                    Null)

 	 If viRetorno = 0 Then
  		bsShowMessage("Processo enviado para execução no servidor!", "I")
 	 Else
  		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemRetorno, "I")
  	 End If
  End If

  Dim QUERY As Object
  Set QUERY = NewQuery

  QUERY.Active = False
  QUERY.Clear
  QUERY.Add("DELETE FROM ANS_TSS_CONTRATO WHERE COMPETENCIA = :HROTINA")
  QUERY.ParamByName("HROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QUERY.ExecSQL

  QUERY.Active = False
  QUERY.Clear
  QUERY.Add("DELETE FROM ANS_TSS_BENEFICIARIO WHERE COMPETENCIA = :HROTINA")
  QUERY.ParamByName("HROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QUERY.ExecSQL

  QUERY.Active = False
  QUERY.Clear
  QUERY.Add("DELETE FROM ANS_TSS_RESUMO WHERE SIBROTINA = :HROTINA")
  QUERY.ParamByName("HROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QUERY.ExecSQL

  Set QUERY = Nothing

  If Not WebMode Then
	RefreshNodesWithTable("ANS_SIB_ROTINA")
  End If

End Sub


Public Sub BOTAOCCO_OnClick()
  Dim Interface As Object
  Dim vsMensagemRetorno As String
  Dim viRetorno As Long

  If VisibleMode Then
  	Set obj = CreateBennerObject("BSINTERFACE0056.SincronizarCCO")
  	obj.SincronizarCCO(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMensagemRetorno)
  	Set Interface = Nothing
  Else
	Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = obj.ExecucaoImediata(CurrentSystem, _
                                    "BSANS002", _
                                    "SincronizarCCO", _
                                    "SIB - Sistema de informações de Beneficiarios - Sincronizar CCO", _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    "ANS_SIB_ROTINA", _
                                    "SITUACAOCCO", _
                                    "", _
                                    "", _
                                    "P", _
                                    False, _
                                    vsMensagemRetorno, _
                                    Null)

 	 If viRetorno = 0 Then
  		bsShowMessage("Processo enviado para execução no servidor!", "I")
 	 Else
  		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemRetorno, "I")
  	 End If
  End If

  If Not WebMode Then
  	RefreshNodesWithTable("ANS_SIB_ROTINA")
  End If
End Sub

Public Sub BOTAOSINCRONIZARCNS_OnClick()
  Dim Interface As Object
  Dim vsMensagemRetorno As String
  Dim viRetorno As Long

  If VisibleMode Then
  	Set obj = CreateBennerObject("BSINTERFACE0056.SincronizarCNS")
  	obj.SincronizarCNS(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMensagemRetorno)
  	Set Interface = Nothing
  Else
	Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = obj.ExecucaoImediata(CurrentSystem, _
                                    "BSANS002", _
                                    "SincronizarCNS", _
                                    "SIB - Sistema de informações de Beneficiarios - Sincronizar CNS", _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    "ANS_SIB_ROTINA", _
                                    "SITUACAOCNS", _
                                    "", _
                                    "", _
                                    "P", _
                                    False, _
                                    vsMensagemRetorno, _
                                    Null)

 	 If viRetorno = 0 Then
  		bsShowMessage("Processo enviado para execução no servidor!", "I")
 	 Else
  		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemRetorno, "I")
  	 End If
  End If

  If Not WebMode Then
  	RefreshNodesWithTable("ANS_SIB_ROTINA")
  End If
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim Interface As Object
  Dim SQL As Object
  Dim vsMensagemRetorno As String
  Dim viRetorno As Long

  If CurrentQuery.State <>1 Then
	bsShowMessage("A tabela não pode estar em edição", "I")
	Exit Sub
  End If

  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT PROCESSADOUSUARIO  ")
  SQL.Add("  FROM SIS_ATUALIZACAO    ")
  SQL.Add(" WHERE HANDLE = :HANDLE   ")

  SQL.ParamByName("HANDLE").AsInteger = 1383
  SQL.Active = True

  If SQL.FieldByName("PROCESSADOUSUARIO").IsNull Then
  	bsShowMessage("A atualização 1383 deve ser processada antes do processamento da rotina SIB.", "I")
    Exit Sub
  End If
  Set SQL = Nothing

  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT PROCESSADOUSUARIO  ")
  SQL.Add("  FROM SIS_ATUALIZACAO    ")
  SQL.Add(" WHERE HANDLE = :HANDLE   ")

  SQL.ParamByName("HANDLE").AsInteger = 1384
  SQL.Active = True

  If SQL.FieldByName("PROCESSADOUSUARIO").IsNull Then
	bsShowMessage("A atualização 1384 deve ser processada antes do processamento da rotina SIB.", "I")
    Exit Sub
  End If
  Set SQL = Nothing

  If CurrentQuery.FieldByName("TABENVIOPONTUAL").AsInteger = 1 Then
    Set SQL = NewQuery
    SQL.Clear
    SQL.Add("SELECT COUNT(1) QTD                  ")
    SQL.Add("  FROM ANS_SIB_ENVIOPONTUAL ENV      ")
    SQL.Add(" WHERE ENV.ROTINA = :HANDLEROTINASIB ")
    SQL.ParamByName("HANDLEROTINASIB").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

    SQL.Active = True

    If SQL.FieldByName("QTD").AsInteger = 0 Then
      bsShowMessage("Não é possível processar uma rotina SIB de envio pontual sem parametrizar os beneficiários à serem enviados.", "I")
      Set SQL = Nothing
      Exit Sub
    End If
    Set SQL = Nothing
  End If

  If VisibleMode Then
	Set obj = CreateBennerObject("BSINTERFACE0056.Processar")
  	obj.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMensagemRetorno)
  	Set Interface = Nothing
  Else
  	Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = obj.ExecucaoImediata(CurrentSystem, _
                                    "BSANS002", _
                                    "Processar", _
                                    "SIB - Sistema de informações de Beneficiarios - Processar", _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    "ANS_SIB_ROTINA", _
                                    "SITUACAOPROCESSO", _
                                    "", _
                                    "", _
                                    "P", _
                                    False, _
                                    vsMensagemRetorno, _
                                    Null)

 	 If viRetorno = 0 Then
  		bsShowMessage("Processo enviado para execução no servidor!", "I")
 	 Else
  		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemRetorno, "I")
  	 End If
  End If

  Dim QUERY As Object
  Dim MSG   As String
  Dim AUX   As Object
  Set AUX   = NewQuery
  Set QUERY = NewQuery

  QUERY.Active = False
  QUERY.Clear
  QUERY.Add("SELECT OCORRENCIA FROM ANS_TSS_RESUMO")
  QUERY.Add("WHERE SIBROTINA = :HROTINA")
  QUERY.ParamByName("HROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QUERY.Active = True

  AUX.Active = False
  AUX.Clear
  AUX.Add("SELECT NOME FROM Z_GRUPOUSUARIOS WHERE HANDLE = :HUSUARIO")
  AUX.ParamByName("HUSUARIO").AsInteger = CurrentUser
  AUX.Active = True

  MSG = QUERY.FieldByName("OCORRENCIA").AsString + Chr(13) + "Processado em "+ CStr(Date) + " por " + AUX.FieldByName("NOME").AsString

  QUERY.Active = False
  QUERY.Clear
  QUERY.Add("UPDATE ANS_TSS_RESUMO SET OCORRENCIA = :TEXTO")
  QUERY.Add("WHERE SIBROTINA = :HROTINA")
  QUERY.ParamByName("TEXTO").AsString = MSG
  QUERY.ParamByName("HROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QUERY.ExecSQL

  Set QUERY = Nothing
  Set AUX   = Nothing

  If Not WebMode Then
  	RefreshNodesWithTable("ANS_SIB_ROTINA")
  End If
End Sub



Public Sub EXPORTAXML_OnClick()
  Dim Interface As Object
  Dim vsMensagemRetorno As String
  Dim viRetorno As Long

  If VisibleMode Then
	Set obj = CreateBennerObject("BSINTERFACE0056.ExportarXML")
  	obj.ExportarXML(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMensagemRetorno)
  	Set Interface = Nothing
  Else
  	Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = obj.ExecucaoImediata(CurrentSystem, _
                                    "BSANS002", _
                                    "ExportarXML", _
                                    "SIB - Sistema de informações de Beneficiarios - Preparar e exportar XML", _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    "ANS_SIB_ROTINA", _
                                    "SITUACAOXML", _
                                    "", _
                                    "", _
                                    "P", _
                                    False, _
                                    vsMensagemRetorno, _
                                    Null)

 	 If viRetorno = 0 Then
  		bsShowMessage("Processo enviado para execução no servidor!", "I")
 	 Else
  		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemRetorno, "I")
  	 End If
  End If

  If Not WebMode Then
  	RefreshNodesWithTable("ANS_SIB_ROTINA")
  End If
End Sub

Public Sub TABLE_AfterScroll()

  BOTAOEXPORTAR.Visible = False
  BOTAOAJUSTARBENEFICIARIO.Visible = False

  If CurrentQuery.FieldByName("TIPO").AsInteger = 3 Then
     ARQUIVO.Caption        = "Arquivo conferência"
  End If

  If CurrentQuery.FieldByName("SITUACAOPROCESSO").AsString = "1" Then
     BOTAOPROCESSAR.Visible = True
     BOTAOCANCELAR.Visible  = False
     PROCESSARTSS.Visible   = False
     CANCELARTSS.Visible    = False
     EXPORTAXML.Visible     = False
     BOTAOCCO.Visible       = False
     BOTAOSINCRONIZARCNS.Visible = False
     Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOPROCESSO").AsString = "2" Then
     BOTAOPROCESSAR.Visible = False
     BOTAOCANCELAR.Visible  = False
     PROCESSARTSS.Visible   = False
     CANCELARTSS.Visible    = False
     EXPORTAXML.Visible     = False
     BOTAOCCO.Visible       = False
     BOTAOSINCRONIZARCNS.Visible = False
     Exit Sub
  End If

  BOTAOPROCESSAR.Visible = False
  BOTAOCANCELAR.Visible  = True

  If CurrentQuery.FieldByName("SITUACAOPROCESSO").AsString = "5" And CurrentQuery.FieldByName("TIPO").AsInteger = 1 Then

     EXPORTAXML.Visible     = True
     BOTAOCCO.Visible       = False
     BOTAOSINCRONIZARCNS.Visible = False
     BOTAOAJUSTARBENEFICIARIO.Visible = True

    If CurrentQuery.FieldByName("SITUACAOPROCESSARTSS").AsString = "1" Then
       PROCESSARTSS.Visible   = True
       CANCELARTSS.Visible    = False
     Else
       PROCESSARTSS.Visible   = False
       CANCELARTSS.Visible    = True
     End If

     Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOPROCESSO").AsString = "5" And CurrentQuery.FieldByName("TIPO").AsInteger = 2 Then
     PROCESSARTSS.Visible   = False
     CANCELARTSS.Visible    = False
     EXPORTAXML.Visible     = False
     BOTAOCCO.Visible       = False
     BOTAOSINCRONIZARCNS.Visible = False
     Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOPROCESSO").AsString = "5" And CurrentQuery.FieldByName("TIPO").AsInteger = 3 Then
     PROCESSARTSS.Visible   = False
     CANCELARTSS.Visible    = False
     EXPORTAXML.Visible     = False
     BOTAOCCO.Visible       = True
     BOTAOSINCRONIZARCNS.Visible = True
     Exit Sub
  End If

End Sub

 Public Function UltimoDiaCompetencia(Data As Date)
 Dim Dia As Variant, Mes As Variant, Ano As Variant
   Dia = Day(Data)
   Mes = Month(Data)
   Ano = Year(Data)
   If Mes = 12 Then
     UltimoDiaCompetencia = FormatDateTime2("dd/mm/yyyy",DateSerial(Ano+1, 01, 01) - 1)
   Else
     UltimoDiaCompetencia = FormatDateTime2("dd/mm/yyyy", DateSerial(Ano, Mes+1, 01) - 1)
   End If
 End Function

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  If ((CurrentQuery.FieldByName("TIPO").AsInteger = 1) And (CurrentQuery.FieldByName("TABENVIOPONTUAL").AsInteger = 1)) Then
    Dim DeleteEnvioPontual As BPesquisa
    Set DeleteEnvioPontual = NewQuery
    DeleteEnvioPontual.Clear
    DeleteEnvioPontual.Add(" DELETE ANS_SIB_ENVIOPONTUAL WHERE ROTINA = :HANDLEROTINASIB ")
    DeleteEnvioPontual.ParamByName("HANDLEROTINASIB").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

    DeleteEnvioPontual.ExecSQL

    Set DeleteEnvioPontual = Nothing
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If Len(CurrentQuery.FieldByName("DIRETORIOXML").AsString) > 1 Then
    If Mid(CurrentQuery.FieldByName("DIRETORIOXML").AsString, Len(CurrentQuery.FieldByName("DIRETORIOXML").AsString),1) = "\" Then
     CurrentQuery.FieldByName("DIRETORIOXML").AsString = Mid(CurrentQuery.FieldByName("DIRETORIOXML").AsString, 1,Len(CurrentQuery.FieldByName("DIRETORIOXML").AsString) - 1)
    End If
  End If

  If CurrentQuery.FieldByName("OPERADORA").IsNull Then
    bsShowMessage("Preencha o campo 'OPERADORA'!", "E")
    CanContinue = False
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("TABTIPOREMESSA").AsString = "1") Then
	bsShowMessage("O tipo de arquivo para o envio não é mais o TXT!", "E")
	CanContinue = False
	Exit Sub
  End If

  Dim SQL As BPesquisa
  Dim SQL2 As BPesquisa
  Set SQL = NewQuery
  Set SQL2 = NewQuery

  If CurrentQuery.FieldByName("TIPO").AsInteger = 1 Then

    If (CurrentQuery.FieldByName("TABENVIOPONTUAL").AsInteger = 2) Then

      If (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime < CurrentQuery.FieldByName("COMPETENCIA").AsDateTime) Or _
         (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime >= DateAdd("m",1,CurrentQuery.FieldByName("COMPETENCIA").AsDateTime)) Then
        bsShowMessage("A data inicial deve estar dentro da competencia informada!", "E")
	    CanContinue = False
	    Exit Sub
      End If

      If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
        bsShowMessage("A data inicial deve ser menor que a data final!", "E")
	    CanContinue = False
	    Exit Sub
      End If

      SQL.Clear
      SQL.Add("SELECT COUNT(1) QTD                  ")
      SQL.Add("  FROM ANS_SIB_ENVIOPONTUAL ENV      ")
      SQL.Add(" WHERE ENV.ROTINA = :HANDLEROTINASIB ")
      SQL.ParamByName("HANDLEROTINASIB").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

      SQL.Active = True

      If SQL.FieldByName("QTD").AsInteger > 0 Then
        bsShowMessage("Não é possível parametrizar a rotina para não realizar envio pontual, pois existem beneficiários à serem enviados ainda cadastrados.", "E")
        CanContinue = False
        Set SQL = Nothing
        Exit Sub
      End If

    Else
      CurrentQuery.FieldByName("DATAINICIAL").AsDateTime = ServerDate
      CurrentQuery.FieldByName("DATAFINAL").AsDateTime = ServerDate
    End If
  End If

  If CurrentQuery.FieldByName("TIPO").AsInteger = 2 Then

	SQL.Active = False
	SQL.Clear
	SQL.Add("SELECT COUNT(1) QTDE                 ")
	SQL.Add("  FROM ANS_SIB_ROTINA                ")
	SQL.Add(" WHERE COMPETENCIA = :COMPETENCIA    ")
	SQL.Add("   AND OPERADORA   = :OPERADORA      ")
	SQL.Add("   AND TIPO        = 1               ")
	SQL.Add("   AND DATAPROCESSAMENTO Is Not Null ")
	SQL.ParamByName("COMPETENCIA").AsDateTime = CurrentQuery.FieldByName("COMPETENCIA").AsDateTime
	SQL.ParamByName("OPERADORA").AsInteger    = CurrentQuery.FieldByName("OPERADORA").AsInteger
	SQL.Active = True

	SQL2.Active = False
	SQL2.Clear
	SQL2.Add("SELECT COUNT(1) QTDE             ")
	SQL2.Add("  FROM ANS_SIB_ROTINA            ")
	SQL2.Add(" WHERE HANDLE     <> :HANDLE     ")
    SQL2.Add("   AND COMPETENCIA = :COMPETENCIA")
	SQL2.Add("   AND OPERADORA   = :OPERADORA  ")
	SQL2.Add("   AND TIPO        = :TIPO       ")
	SQL2.ParamByName("HANDLE").AsInteger       = CurrentQuery.FieldByName("HANDLE").AsInteger
	SQL2.ParamByName("COMPETENCIA").AsDateTime = CurrentQuery.FieldByName("COMPETENCIA").AsDateTime
	SQL2.ParamByName("OPERADORA").AsInteger    = CurrentQuery.FieldByName("OPERADORA").AsInteger
	SQL2.ParamByName("TIPO").AsInteger         = CurrentQuery.FieldByName("TIPO").AsInteger
	SQL2.Active = True

	If SQL.FieldByName("QTDE").AsInteger <= SQL2.FieldByName("QTDE").AsInteger Then
        bsShowMessage("Não existe nenhuma rotina de remessa processada ou já existe uma rotina de devolução para essa competência e operadora.", "E")
        CanContinue = False
        Set SQL  = Nothing
		Set SQL2 = Nothing
        Exit Sub
	End If
  End If

  Set SQL  = Nothing
  Set SQL2 = Nothing

End Sub

Function ExisteOutraNaCompetencia As Boolean

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT COUNT(1) QTDE             ")
  SQL.Add("  FROM ANS_SIB_ROTINA            ")
  SQL.Add(" WHERE HANDLE     <> :HANDLE     ")
  SQL.Add("   AND COMPETENCIA = :COMPETENCIA")
  SQL.Add("   AND OPERADORA   = :OPERADORA  ")
  SQL.Add("   AND TIPO        = :TIPO       ")

  SQL.ParamByName("HANDLE").AsInteger       = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("COMPETENCIA").AsDateTime = CurrentQuery.FieldByName("COMPETENCIA").AsDateTime
  SQL.ParamByName("OPERADORA").AsInteger    = CurrentQuery.FieldByName("OPERADORA").AsInteger
  SQL.ParamByName("TIPO").AsInteger         = CurrentQuery.FieldByName("TIPO").AsInteger

  SQL.Active = True

  ExisteOutraNaCompetencia = SQL.FieldByName("QTDE").AsInteger > 0

  SQL.Active = False
  Set SQL = Nothing

End Function

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
 	Select Case CommandID
 		Case "BOTAOCANCELAR"
 			BOTAOCANCELAR_OnClick
 		Case "BOTAOCCO"
 			BOTAOCCO_OnClick
 		Case "BOTAOPROCESSAR"
 			BOTAOPROCESSAR_OnClick
 	    Case "EXPORTAXML"
 	    	EXPORTAXML_OnClick
 	    Case "BOTAOSINCRONIZARCNS"
 	    	BOTAOSINCRONIZARCNS_OnClick
 	    Case "BOTAOAJUSTARBENEFICIARIO"
 	        BOTAOAJUSTARBENEFICIARIO_OnClick
 	    Case "PROCESSARTSS"
 	        PROCESSARTSS_OnClick
 	    Case "CANCELARTSS"
 	        CANCELARTSS_OnClick
	End Select
End Sub

Public Sub PROCESSARTSS_OnClick()
       Dim component As CSBusinessComponent
       Set component = BusinessComponent.CreateInstance("Benner.Saude.Ans.Business.AnsSibRotinaBLL, Benner.Saude.Ans.Business")
       component.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
       bsShowMessage(component.Execute("ProcessarTssClickMacro"),"I")
       Set component=Nothing
End Sub

Public Sub CANCELARTSS_OnClick()
       Dim component As CSBusinessComponent
	   Set component = BusinessComponent.CreateInstance("Benner.Saude.Ans.Business.AnsSibRotinaBLL, Benner.Saude.Ans.Business")
 	   component.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
 	   bsShowMessage(component.Execute("CancelarTssClickMacro"),"I")
 	   Set component=Nothing
End Sub

Public Sub TIPO_OnChange()

  Dim SQL As Object
  Set SQL = NewQuery

  CurrentQuery.FieldByName("DATAINICIAL").AsDateTime = CurrentQuery.FieldByName("COMPETENCIA").AsDateTime
  CurrentQuery.FieldByName("DATAFINAL").AsDateTime   = DateAdd("d",-1, DateSerial(DatePart("yyyy",CurrentQuery.FieldByName("COMPETENCIA").AsDateTime),DatePart("m",CurrentQuery.FieldByName("COMPETENCIA").AsDateTime) + 1,1))



  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT MAX(DATAFINAL) DATAFINAL  ")
  SQL.Add("  FROM ANS_SIB_ROTINA            ")
  SQL.Add(" WHERE TIPO      = 1             ")

  SQL.Add("   AND SITUACAO  = 'P'           ")
  SQL.Add("   AND OPERADORA = :OPERADORA    ")
  SQL.Add("   AND COMPETENCIA = :COMPETENCIA ")

  SQL.ParamByName("COMPETENCIA").AsDateTime    = CurrentQuery.FieldByName("COMPETENCIA").AsDateTime
  SQL.ParamByName("OPERADORA").AsInteger      = CurrentQuery.FieldByName("OPERADORA").AsInteger

  SQL.Active = True

  If Not SQL.FieldByName("DATAFINAL").IsNull Then

    CurrentQuery.FieldByName("DATAINICIAL").AsDateTime = DateAdd("d", 1, SQL.FieldByName("DATAFINAL").AsDateTime)
    CurrentQuery.FieldByName("DATAFINAL").AsDateTime   = DateAdd("d",-1, DateSerial(DatePart("yyyy",CurrentQuery.FieldByName("DATAINICIAL").AsDateTime),DatePart("m",CurrentQuery.FieldByName("DATAINICIAL").AsDateTime) + 1,1))
  End If

  SQL.Active = False
  Set SQL = Nothing

End Sub
