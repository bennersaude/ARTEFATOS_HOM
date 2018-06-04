'HASH: 2B37E9BC8A9ED0FAF48F25EA997F1DBB
 
Option Explicit
Public Sub BOTAOBUSCARTEXTO_OnClick()
  Dim interface As Object
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTextoObservacao As String
  Dim vHandleTextoPadrao As Long
  Dim qSelecionaTextoPadrao As Object
  Set qSelecionaTextoPadrao = NewQuery

  Set interface=CreateBennerObject("Procura.Procurar")

  vTextoObservacao = CurrentQuery.FieldByName("OBSERVACAO").AsString

  vColunas = "CODIGO|RESUMO"
  vCampos = "Código|Resumo"

  vHandleTextoPadrao = interface.Exec(CurrentSystem, "SAM_TEXTOPADRAO", vColunas, 1, vCampos, vCriterio, "Texto padrão", True, "")

  qSelecionaTextoPadrao.Clear
  qSelecionaTextoPadrao.Active = False
  qSelecionaTextoPadrao.Add(" SELECT TEXTO            ")
  qSelecionaTextoPadrao.Add("   FROM SAM_TEXTOPADRAO  ")
  qSelecionaTextoPadrao.Add("  WHERE HANDLE = :HANDLE ")
  qSelecionaTextoPadrao.ParamByName("HANDLE").AsInteger = vHandleTextoPadrao
  qSelecionaTextoPadrao.Active = True

  If CurrentQuery.FieldByName("OBSERVACAO").IsNull Or CurrentQuery.FieldByName("OBSERVACAO").AsString = "" Then
    vTextoObservacao = qSelecionaTextoPadrao.FieldByName("TEXTO").AsString
  Else
    vTextoObservacao = vTextoObservacao + qSelecionaTextoPadrao.FieldByName("TEXTO").AsString
  End If

  CurrentQuery.FieldByName("OBSERVACAO").AsString = vTextoObservacao

  Set qSelecionaTextoPadrao = Nothing
  Set interface = Nothing

End Sub

Public Sub BOTAOLIMPATEXTO_OnClick()
  CurrentQuery.FieldByName("OBSERVACAO").AsString = ""
End Sub
