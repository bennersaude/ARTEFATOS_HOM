'HASH: B4E8E2E4E2F745C831E26F0B512B9264

'#Uses "*bsShowMessage"

Option Explicit

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  '#Uses "*ProcuraPrestador"
  '  If Len(PRESTADOR.Text) = 0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraPrestador("C", "T", PRESTADOR.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If
  '  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim sqlVerificaExistencia As Object
  Set sqlVerificaExistencia = NewQuery

  sqlVerificaExistencia.Clear
  sqlVerificaExistencia.Add("SELECT A.CONTRATOPRECO                                          ")
  sqlVerificaExistencia.Add("  FROM SAM_PLANO_PRECO_PRESTADOR A                              ")
  sqlVerificaExistencia.Add("  JOIN SAM_PLANO_PRECO B On B.HANDLE = A.CONTRATOPRECO          ")
  sqlVerificaExistencia.Add("WHERE B.PLANO = (SELECT PLANO                                   ")
  sqlVerificaExistencia.Add("                        FROM SAM_PLANO_PRECO                    ")
  sqlVerificaExistencia.Add("                       WHERE HANDLE = :CONTRATOPRECO)           ")
  sqlVerificaExistencia.Add("  AND A.PRESTADOR = :PRESTADOR                                  ")
  sqlVerificaExistencia.Add("  AND A.HANDLE <> :HANDLE                                       ")

  sqlVerificaExistencia.ParamByName("CONTRATOPRECO").Value = CurrentQuery.FieldByName("CONTRATOPRECO").AsString
  sqlVerificaExistencia.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").Value
  sqlVerificaExistencia.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value

  sqlVerificaExistencia.Active = True

  If Not sqlVerificaExistencia.EOF Then
    If sqlVerificaExistencia.FieldByName("CONTRATOPRECO").AsString = CurrentQuery.FieldByName("CONTRATOPRECO").AsString Then
      bsShowMessage("Este Prestador já está relacionado nesta tabela de preço", "E")
    Else
      bsShowMessage("Este Prestador já está relacionada em outra tabela de preço", "E")
    End If
    CanContinue = False
  End If
  Set sqlVerificaExistencia = Nothing
End Sub

