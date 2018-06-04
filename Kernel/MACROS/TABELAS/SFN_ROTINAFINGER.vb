'HASH: 60E924AA7C6B67963EB92BF7E8E178B3

Option Explicit

Public Sub BOTAOCANCELAR_OnClick()
  Dim dll As Object
  Set dll = CreateBennerObject("SfnGeraFatura.Rotinas")
  dll.Cancelar(CurrentSystem)
  Set dll = Nothing
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim dll As Object
  Set dll = CreateBennerObject("SfnGeraFatura.Rotinas")
  dll.Processar(CurrentSystem)
  Set dll = Nothing
End Sub


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  VerificaSeProcessada(CanContinue)
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  VerificaSeProcessada(CanContinue)
End Sub

Public Sub VerificaSeProcessada(CanContinue As Boolean)
  Dim SQLRotFin As Object
  Set SQLRotFin = NewQuery
  SQLRotFin.Add("SELECT SITUACAO FROM SFN_ROTINAFIN WHERE HANDLE = :HANDLE")
  SQLRotFin.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("ROTINAFIN").Value
  SQLRotFin.Active = True
  If SQLRotFin.FieldByName("SITUACAO").Value = "P" Then
    CanContinue = False
    SQLRotFin.Active = False
    Set SQLRotFin = Nothing
    MsgBox("A Rotina já foi processada")
    Exit Sub
  End If
  SQLRotFin.Active = False
  Set SQLRotFin = Nothing
End Sub

