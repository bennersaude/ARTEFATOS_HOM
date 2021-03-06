﻿'HASH: 89E8EFE5A35B2E46DFA617FA190420E7
'Macro: SAM_REAJUSTEPRC_PRESTADOR_PAC
'#Uses "*NegociacaoPrecos"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vFiltroAdicional As String
  Dim vAtedias As Integer
  Dim vDeDias As Integer
  Dim vAteAnos As Integer
  Dim vDeAnos As Integer

  If VisibleMode Then

    vFiltroAdicional = " AND PRESTADOR = " + CurrentQuery.FieldByName("PRESTADOR").AsString

    If Not CurrentQuery.FieldByName("GRAUAGERAR").IsNull Then
		vFiltroAdicional = vFiltroAdicional + " AND GRAUAGERAR = " + CurrentQuery.FieldByName("GRAUAGERAR").AsString
	End If

	If CurrentQuery.FieldByName("ATEDIAS").IsNull Then
    vAtedias = -1
  Else
    vAtedias = CurrentQuery.FieldByName("ATEDIAS").AsInteger
  End If

  If CurrentQuery.FieldByName("ATEANOS").IsNull Then
    vAteAnos = -1
  Else
  	vAteAnos = CurrentQuery.FieldByName("ATEANOS").AsInteger
  End If

  If CurrentQuery.FieldByName("DEDIAS").IsNull Then
    vDeDias = -1
  Else
    vDeDias = CurrentQuery.FieldByName("DEDIAS").AsInteger
  End If

  If CurrentQuery.FieldByName("DEANOS").IsNull Then
    vDeAnos = -1
  Else
    vDeAnos = CurrentQuery.FieldByName("DEANOS").AsInteger
  End If

	CanContinue = ValidacoesBeforePostNegociacaoPreco(CurrentQuery.FieldByName("HANDLE").AsInteger, _
	  "SAM_REAJUSTEPRC_PRESTADOR_PAC", "", "", "PRESTADOR", _
	  CurrentQuery.FieldByName("PRESTADOR").AsInteger, CurrentQuery.FieldByName("EVENTO").AsInteger, "", "-", _
	  vFiltroAdicional, vDeAnos, vDeDias, _
	  vAteAnos, vAtedias, _
	  CurrentQuery.FieldByName("TABNEGOCIACAO").AsInteger, 0, 0)

    If Not CanContinue Then
      Exit Sub
	End If
  End If
End Sub
