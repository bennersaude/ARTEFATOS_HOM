'HASH: F02143EA082C3C07CA36AB2CC905706B
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"
Option Explicit

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTO.LocateText)

  If vHandle > 0 Then
	CurrentQuery.Edit
	CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
End Sub

Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTOINICIAL.LocateText)

  If vHandle > 0 Then
	CurrentQuery.Edit
	CurrentQuery.FieldByName("EVENTOINICIAL").Value = vHandle
  End If
End Sub

Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTOFINAL.LocateText)

  If vHandle > 0 Then
	CurrentQuery.Edit
	CurrentQuery.FieldByName("EVENTOFINAL").Value = vHandle
  End If
End Sub

Public Sub TABLE_NewRecord()
  Dim handleReajustePrecoSimp As Long
  Dim qReajustePrecoSimp As Object

  handleReajustePrecoSimp = RecordHandleOfTable("REAJUSTEPRECOSIMP")

  Set qReajustePrecoSimp = NewQuery
  qReajustePrecoSimp.Clear
  qReajustePrecoSimp.Active = False
  qReajustePrecoSimp.Add("SELECT TIPODEREAJUSTE")
  qReajustePrecoSimp.Add("FROM REAJUSTEPRECOSIMP")
  qReajustePrecoSimp.Add("WHERE HANDLE = :HANDLE")
  qReajustePrecoSimp.ParamByName("HANDLE").AsInteger = handleReajustePrecoSimp
  qReajustePrecoSimp.Active = True

  CurrentQuery.FieldByName("TIPODEREAJUSTE").AsInteger = qReajustePrecoSimp.FieldByName("TIPODEREAJUSTE").AsInteger
End Sub

Public Sub TABLE_OnInsertBtnClick(CanContinue As Boolean)
  Dim qReajustePrecoSimp As Object
  Set qReajustePrecoSimp = NewQuery
  qReajustePrecoSimp.Clear
  qReajustePrecoSimp.Active = False
  qReajustePrecoSimp.Add("SELECT TIPODEREAJUSTE")
  qReajustePrecoSimp.Add("FROM REAJUSTEPRECOSIMP")
  qReajustePrecoSimp.Add("WHERE HANDLE = :HANDLE")
  qReajustePrecoSimp.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("REAJUSTEPRECOSIMP").AsInteger
  qReajustePrecoSimp.Active = True

  CurrentQuery.FieldByName("TIPODEREAJUSTE").AsInteger = qReajustePrecoSimp.FieldByName("TIPODEREAJUSTE").AsInteger
End Sub
