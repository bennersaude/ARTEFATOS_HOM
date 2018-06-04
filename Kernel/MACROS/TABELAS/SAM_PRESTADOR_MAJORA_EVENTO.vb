'HASH: 3F45DA6DFEA2F8EC09F30180AFBA4219
'Macro: SAM_PRESTADOR_MAJORA_EVENTO
Option Explicit
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Public Function VigenciaFechada As Boolean
  VigenciaFechada = False

  Dim vSQL As BPesquisa
  Set vSQL = NewQuery
  vSQL.Clear
  vSQL.Add("SELECT DATAFINAL                  ")
  vSQL.Add("  FROM SAM_PRESTADOR_MAJORA       ")
  vSQL.Add(" WHERE HANDLE = :HPRESTADORMAJORA ")
  vSQL.ParamByName("HPRESTADORMAJORA").Value = CurrentQuery.FieldByName("PRESTADORMAJORA").AsInteger
  vSQL.Active = True
  If Not vSQL.FieldByName("DATAFINAL").IsNull Then
    If vSQL.FieldByName("DATAFINAL").AsDateTime < ServerDate Then
  	  bsShowMessage("Majoração finalizada não permite manutenções", "E")
	  VigenciaFechada = True
    End If
  End If
  vSQL.Active = False
  Set vSQL = Nothing
End Function

Public Function EventosDuplicados As Boolean
	Dim vSQL As BPesquisa
	Set vSQL = NewQuery
	vSQL.Clear
	vSQL.Add("SELECT HANDLE                             ")
	vSQL.Add("  FROM SAM_PRESTADOR_MAJORA_EVENTO        ")
	vSQL.Add(" WHERE PRESTADORMAJORA = :PRESTADORMAJORA ")
	vSQL.Add("   And EVENTO = :EVENTO                   ")
	vSQL.Add("   And HANDLE <> :HANDLE                  ")
	vSQL.ParamByName("PRESTADORMAJORA").Value = CurrentQuery.FieldByName("PRESTADORMAJORA").AsString
	vSQL.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTO").AsString
	vSQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
	vSQL.Active = True
	EventosDuplicados = False
	If Not vSQL.FieldByName("HANDLE").IsNull Then
		EventosDuplicados = True
		bsShowMessage("Não é possível incluir um Evento duplicado", "E")
	End If
	vSQL.Active = False
	Set vSQL = Nothing
End Function
Public Function EventoUsado As Boolean
	Dim vSQL1 As BPesquisa
	Set vSQL1 = NewQuery
	vSQL1.Clear
	vSQL1.Add("SELECT DATAINICIAL,               ")
	vSQL1.Add("       PRESTADOR                   ")
	vSQL1.Add("  FROM SAM_PRESTADOR_MAJORA       ")
	vSQL1.Add(" WHERE HANDLE = :HPRESTADORMAJORA ")
    vSQL1.ParamByName("HPRESTADORMAJORA").Value = CurrentQuery.FieldByName("PRESTADORMAJORA").AsInteger
	vSQL1.Active = True

	Dim vSQL2 As BPesquisa
    Set vSQL2 = NewQuery
	vSQL2.Clear
	vSQL2.Add("SELECT COUNT(1) QTDEVENTOS                        ")
    vSQL2.Add("  FROM SAM_GUIA_EVENTOS EVENTO                    ")
    vSQL2.Add("  Join SAM_GUIA GUIA On GUIA.Handle = EVENTO.GUIA ")
    vSQL2.Add("  Join SAM_PEG PEG On PEG.Handle = GUIA.PEG       ")
    vSQL2.Add(" WHERE PEG.SITUACAO In (2,3,4,6)                  ")
    vSQL2.Add("   AND EVENTO.EVENTO = :HANDLE                    ")
    vSQL2.Add("   AND EVENTO.DATAATENDIMENTO >= :DATA            ")
    vSQL2.Add("   AND PEG.RECEBEDOR = :HRECEBEDOR                ")
 	vSQL2.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("EVENTO").AsInteger
    vSQL2.ParamByName("DATA").Value = vSQL1.FieldByName("DATAINICIAL").AsDateTime
	vSQL2.ParamByName("HRECEBEDOR").Value = vSQL1.FieldByName("PRESTADOR").AsInteger
	vSQL2.Active = True

	EventoUsado = False

	If vSQL2.FieldByName("QTDEVENTOS").AsInteger > 0 Then
	   EventoUsado = True
       bsShowMessage("Não é possível excluir um Evento utilizado", "E")
	End If

    vSQL1.Active = False
	Set vSQL1 = Nothing
	vSQL2.Active = False
	Set vSQL2 = Nothing
End Function


Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTO.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
 Dim Msg As String
 If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
 Else
   If VigenciaFechada Then
      CanContinue = False
   Else
     If EventoUsado Then
        CanContinue = False
     End If
   End If
 End If
Exit Sub
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
 Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
  If VigenciaFechada Then
     CanContinue = False
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)'
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

If VigenciaFechada Then
   CanContinue = False
Else
  If EventosDuplicados Then
     CanContinue = False
  End If
End If
End Sub
