'HASH: E5E37073497E46D3DED5F4424806ED45
'Macro: SFN_CONTABPAG_REGRA


' Não executar o Procura para a classe gerencial,pois pode trazer classe gerencial que não é o ultimo nível.
'(07/03/2001 -by celso)

Public Sub XCLASSEGERENCIAL_OnPopup(ShowPopup As Boolean)
  Dim interface As Object

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String


  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_CLASSEGERENCIAL.ESTRUTURA|SFN_CLASSEGERENCIAL.DESCRICAO|SFN_CLASSEGERENCIAL.CODIGOREDUZIDO|SFN_CLASSEGERENCIAL.NATUREZA|SFN_CLASSEGERENCIAL.HISTORICO"

  vCriterio = "HANDLE>0"

  vCampos = "Estrutura|Descrição|Código|D/C|Historico"

  vHandle = interface.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vColunas, 1, vCampos, vCriterio, "Classe Gerencial", False, CLASSEGERENCIAL.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLASSEGERENCIAL").Value = vHandle
  End If
  Set INTERFACE = Nothing
End Sub

