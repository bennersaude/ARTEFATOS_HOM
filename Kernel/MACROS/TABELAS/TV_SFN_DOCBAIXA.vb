'HASH: 43BEC0A69CE0086E5A0B18456DC044FF
 
'#Uses "*bsShowMessage

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim qDoc As Object
  Dim vHandleDoc As String

  Set qDoc = NewQuery

  With qDoc
    .Active = False
    .Clear
    .Add("SELECT HANDLE,")
    .Add("       TESOURARIA,")
    .Add("       NUMEROCHEQUE,")
    .Add("       BANCO,")
    .Add("       AGENCIA,")
    .Add("       BAIXAMOTIVO,")
    .Add("       NOME,")
    .Add("       NUMERO,")
    .Add("       DATAVENCIMENTO,")
    .Add("       DATAEMISSAO,")
    .Add("       BAIXADATA,")
    .Add("       CANCDATA,")
    .Add("       VALOR,")
    .Add("       VALORJURO,")
    .Add("       VALORMULTA,")
    .Add("       VALORCORRECAO,")
    .Add("       VALORDESCONTO,")
    .Add("       VALORTOTAL")
    .Add("  FROM SFN_DOCUMENTO")
    .Add(" WHERE HANDLE = :HANDLE")
  End With

  If Not(CurrentQuery.FieldByName("DOCUMENTO").IsNull) Then
    qDoc.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("DOCUMENTO").AsInteger
    qDoc.Active = True
  Else
    vHandleDoc = SessionVar("WebHandleDocumento")
    qDoc.ParamByName("HANDLE").AsString = vHandleDoc
    qDoc.Active = True
  End If

  '// Verifica se achou o documento \\'
  If qDoc.FieldByName("HANDLE").IsNull Then
    BsShowMessage("Documento não encontrado", "E")
	GoTo Fim_Erro
  End If

  '// Verifica se documento já está baixado \\'
  If Not(qDoc.FieldByName("BAIXADATA").IsNull) Then
	bsShowMessage("Este documento já foi baixado anteriormente. Processo cancelado.", "E")
	GoTo Fim_Erro
  End If

  '// Verifica se documento já está cancelado \\'
  If Not(qDoc.FieldByName("CANCDATA").IsNull) Then
	bsShowMessage("Este documento está cancelado. Impossível baixar.", "E")
	GoTo Fim_Erro
  End If

  GoTo Fim_Normal

  Fim_Erro:
    CanContinue = False

  Fim_Normal:
    Set qDoc = Nothing
End Sub

Public Sub TABLE_AfterInsert()
  Dim qDoc       As Object
  Dim vHandleDoc As String
  Dim Interface  As Object

  Dim vValor     As Double
  Dim vJuro      As Double
  Dim vMulta     As Double
  Dim vCorrecao  As Double
  Dim vDesconto  As Double

  Set qDoc = NewQuery

  With qDoc
    .Active = False
    .Clear
    .Add("SELECT HANDLE,")
    .Add("       TESOURARIA,")
    .Add("       NUMEROCHEQUE,")
    .Add("       BANCO,")
    .Add("       AGENCIA,")
    .Add("       BAIXAMOTIVO,")
    .Add("       NOME,")
    .Add("       NUMERO,")
    .Add("       DATAVENCIMENTO,")
    .Add("       DATAEMISSAO,")
    .Add("       BAIXADATA,")
    .Add("       CANCDATA,")
    .Add("       VALOR,")
    .Add("       VALORJURO,")
    .Add("       VALORMULTA,")
    .Add("       VALORCORRECAO,")
    .Add("       VALORDESCONTO,")
    .Add("       VALORTOTAL")
    .Add("  FROM SFN_DOCUMENTO")
    .Add(" WHERE HANDLE = :HANDLE")
  End With

  If Not(CurrentQuery.FieldByName("DOCUMENTO").IsNull) Then
    qDoc.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("DOCUMENTO").AsInteger
    qDoc.Active = True
  Else
    vHandleDoc = SessionVar("WebHandleDocumento")
    qDoc.ParamByName("HANDLE").AsString = vHandleDoc
    qDoc.Active = True
  End If

  CurrentQuery.FieldByName("DOCUMENTO").AsInteger       = qDoc.FieldByName("HANDLE").AsInteger
  CurrentQuery.FieldByName("TESOURARIA").AsInteger      = qDoc.FieldByName("TESOURARIA").AsInteger
  CurrentQuery.FieldByName("DATABAIXA").AsDateTime      = ServerNow
  CurrentQuery.FieldByName("DATACONTABIL").AsDateTime   = ServerNow
  CurrentQuery.FieldByName("NUMEROCHEQUE").AsString     = qDoc.FieldByName("NUMEROCHEQUE").AsString
  CurrentQuery.FieldByName("BANCO").AsInteger           = qDoc.FieldByName("BANCO").AsInteger
  CurrentQuery.FieldByName("AGENCIA").AsInteger         = qDoc.FieldByName("AGENCIA").AsInteger
  CurrentQuery.FieldByName("BAIXAMOTIVO").AsString      = qDoc.FieldByName("BAIXAMOTIVO").AsString
  CurrentQuery.FieldByName("NOME").AsString             = qDoc.FieldByName("NOME").AsString
  CurrentQuery.FieldByName("NUMERO").AsInteger          = qDoc.FieldByName("NUMERO").AsInteger
  CurrentQuery.FieldByName("DATAVENCIMENTO").AsDateTime = qDoc.FieldByName("DATAVENCIMENTO").AsDateTime
  CurrentQuery.FieldByName("DATAEMISSAO").AsDateTime    = qDoc.FieldByName("DATAEMISSAO").AsDateTime

  Set Interface = CreateBennerObject("SfnBaixa.Documento")
  Interface.BxCalcDocumento _
    (CurrentSystem, _
     qDoc.FieldByName("HANDLE").AsInteger, _
     ServerNow, _
     qDoc.FieldByName("TESOURARIA").AsInteger, _
     vValor, _
     vJuro, _
     vMulta, _
     vCorrecao, _
     vDesconto)
  Set Interface = Nothing

  CurrentQuery.FieldByName("VALOR").AsFloat         = vValor
  CurrentQuery.FieldByName("VALORJURO").AsFloat     = vJuro
  CurrentQuery.FieldByName("VALORMULTA").AsFloat    = vMulta
  CurrentQuery.FieldByName("VALORCORRECAO").AsFloat = vCorrecao
  CurrentQuery.FieldByName("VALORDESCONTO").AsFloat = vDesconto
  CurrentQuery.FieldByName("VALORTOTAL").AsFloat    = (vValor + vJuro + vMulta + vCorrecao - vDesconto)


End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim vMensagemErro As String
  Dim Interface     As Object
  Dim SQL           As Object

  On Error GoTo Erro

  '// Verifica documento está associado a uma rotina arquivo\\'
  Set SQL  = NewQuery
  With SQL
    .Clear
    .Active = False
    .Add("SELECT D.NUMERO                              ")
    .Add("  FROM SFN_DOCUMENTO D,                      ")
    .Add("       SFN_ROTINAARQUIVO_DOC RAD             ")
    .Add(" WHERE D.ULTIMAROTINAARQUIVODOC = RAD.HANDLE ")
    .Add("   AND RAD.TABENVIORETORNO      = 1          ")
    .Add("   AND D.HANDLE                 = :DOCUMENTO ")
    .ParamByName("DOCUMENTO").AsInteger = CurrentQuery.FieldByName("DOCUMENTO").AsInteger
    .Active = True
  End With

  If ((SQL.EOF) Or _
      ( Not(SQL.EOF) And _
	   (bsShowMessage("Documento com rotina arquivo de ENVIO" + Chr(13) + "Deseja continuar ?", "Q") = vbYes))) Then

      Set Interface = CreateBennerObject("SfnBaixa.Documento")
      vMensagemErro = Interface.BxDocWeb _
        (CurrentSystem, _
         CurrentQuery.FieldByName("DOCUMENTO").AsInteger, _
         CurrentQuery.FieldByName("TESOURARIA").AsInteger, _
         CurrentQuery.FieldByName("DATABAIXA").AsDateTime, _
         CurrentQuery.FieldByName("DATACONTABIL").AsDateTime, _
         CurrentQuery.FieldByName("NUMEROCHEQUE").AsString, _
         CurrentQuery.FieldByName("BANCO").AsInteger, _
         CurrentQuery.FieldByName("AGENCIA").AsInteger, _
         CurrentQuery.FieldByName("BAIXAMOTIVO").AsString, _
         CurrentQuery.FieldByName("VALOR").AsFloat, _
         CurrentQuery.FieldByName("VALORJURO").AsFloat, _
         CurrentQuery.FieldByName("VALORMULTA").AsFloat, _
         CurrentQuery.FieldByName("VALORCORRECAO").AsFloat, _
         CurrentQuery.FieldByName("VALORDESCONTO").AsFloat)
      Set Interface = Nothing

      If Len(Trim(vMensagemErro)) <= 0 Then
        bsShowMessage("Baixa concluída", "I")
        Exit Sub
      End If

      GoTo MostrarMsgErro
  End If
  Set SQL  = Nothing

  Exit Sub

  Erro:
    vMensagemErro = vMensagemErro + " - Um erro ocorreu ao chamar a DLL: " + Err.Description
    GoTo MostrarMsgErro

  MostrarMsgErro:
    bsShowMessage(vMensagemErro, "E")
    CanContinue = False
End Sub
