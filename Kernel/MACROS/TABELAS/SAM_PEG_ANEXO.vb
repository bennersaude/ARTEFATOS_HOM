'HASH: 3248B8DFDA2B111BDB81D53BC51C30FA
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qVerificaAnexoNfPeg As Object
  Set qVerificaAnexoNfPeg = NewQuery

  qVerificaAnexoNfPeg.Clear
  qVerificaAnexoNfPeg.Add(" SELECT HANDLE ")
  qVerificaAnexoNfPeg.Add("   FROM SAM_PEG_ANEXO ")
  qVerificaAnexoNfPeg.Add("  WHERE HANDLE = :HANDLE ")
  qVerificaAnexoNfPeg.Add("  AND NFANEXADAPORTAL = 'S' ")
  qVerificaAnexoNfPeg.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qVerificaAnexoNfPeg.Active = True

  If Not qVerificaAnexoNfPeg.FieldByName("HANDLE").IsNull Then
    bsShowMessage("Já existe uma Nota Fiscal anexada para este(s) PEG(s).", "i")
    CanContinue = False
  End If

End Sub
