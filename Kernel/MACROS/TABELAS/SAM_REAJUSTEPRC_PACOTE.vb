'HASH: D6900552AEF064608458D633BB7A9295
Option Explicit
'#Uses "*bsShowMessage"
'#Uses "*NegociacaoPrecos"


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim validaNegociacao As String
  Dim vAtedias As Integer
  Dim vDeDias As Integer
  Dim vAteAnos As Integer
  Dim vDeAnos As Integer

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

  validaNegociacao = ValidarTipoNegociacao(vDeAnos, vDeDias, vAteAnos, vAtedias, CurrentQuery.FieldByName("TABNEGOCIACAO").AsInteger)

  If (validaNegociacao <> "") Then
	bsShowMessage(validaNegociacao, "E")
	CanContinue = False
	Exit Sub
  End If

  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > _
                              CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then

	bsShowMessage("DATA INICIAL não pode ser maior que a DATA FINAL", "E")

    CanContinue = False
  ElseIf CurrentQuery.FieldByName("NOVAVIGENCIA").AsDateTime <= _
                                    CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then


	bsShowMessage("NOVA VIGÊNCIA deve ser maior que a DATA FINAL", "E")

    CanContinue = False
  End If
End Sub
