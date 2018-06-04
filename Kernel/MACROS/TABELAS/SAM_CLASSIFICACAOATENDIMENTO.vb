'HASH: B9DF8FD290C189113E7B442C7951D532
 
'#Uses "*bsShowMessage"
Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qVerificaSeExisteClassificacaoPadrao As BPesquisa
  Set qVerificaSeExisteClassificacaoPadrao = NewQuery

  qVerificaSeExisteClassificacaoPadrao.Active = False
  qVerificaSeExisteClassificacaoPadrao.Clear
  qVerificaSeExisteClassificacaoPadrao.Add("SELECT HANDLE                                    ")
  qVerificaSeExisteClassificacaoPadrao.Add("  FROM SAM_CLASSIFICACAOATENDIMENTO              ")
  qVerificaSeExisteClassificacaoPadrao.Add(" WHERE HANDLE <> :HANDLE                         ")
  qVerificaSeExisteClassificacaoPadrao.Add("   AND TIPOAUTORIZACAO = :TIPOAUTORIZACAO        ")
  qVerificaSeExisteClassificacaoPadrao.Add("   AND CONDICAOATENDIMENTO = :CONDICAOATENDIMENTO")

  qVerificaSeExisteClassificacaoPadrao.ParamByName("TIPOAUTORIZACAO").Value = CurrentQuery.FieldByName("TIPOAUTORIZACAO").AsInteger
  qVerificaSeExisteClassificacaoPadrao.ParamByName("CONDICAOATENDIMENTO").Value = CurrentQuery.FieldByName("CONDICAOATENDIMENTO").AsInteger
  qVerificaSeExisteClassificacaoPadrao.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qVerificaSeExisteClassificacaoPadrao.Active = True

  If (Not qVerificaSeExisteClassificacaoPadrao.EOF) Then
    MsgBox("Já existe uma classificação de atendimento com estas configurações!")
    CanContinue = False
  End If

  Set qVerificaSeExisteClassificacaoPadrao = Nothing
End Sub
