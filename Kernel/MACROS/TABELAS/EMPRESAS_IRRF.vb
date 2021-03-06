﻿'HASH: F0AD40A930EADA01AF9337545C18B42A

'Macro: EMPRESAS_IRRF
'#Uses "*bsShowMessage"

Public Sub CENTROCUSTO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String


  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_CENTROCUSTO.ESTRUTURA|SFN_CENTROCUSTO.DESCRICAO|SFN_CENTROCUSTO.CODIGOREDUZIDO"

  vCriterio = "HANDLE>0"

  vCampos = "Estrutura|Descrição|Código"

  vHandle = interface.Exec(CurrentSystem, "SFN_CENTROCUSTO", vColunas, 1, vCampos, vCriterio, "Centro de Custo", False, CENTROCUSTO.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CENTROCUSTO").Value = vHandle
  End If
  Set interface = Nothing

End Sub

Public Sub CLASSEGERENCIAL_OnPopup(ShowPopup As Boolean)
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
  Set interface = Nothing

End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
    COMPETENCIAFINAL.ReadOnly = False
  Else
    COMPETENCIAFINAL.ReadOnly = True
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim interface As Object
  Dim Linha As String

  Set interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = interface.Vigencia(CurrentSystem, "EMPRESAS_IRRF", "COMPETENCIAINICIAL", "COMPETENCIAFINAL", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime, CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime, "EMPRESA", "")

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If
  Set interface = Nothing

End Sub

