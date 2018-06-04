'HASH: C11A71D7A0E96F583AC54780F87766BB
'#Uses "*bsShowMessage"

'Macro: LOGRADOUROS

Option Explicit

Dim oldCep       As String
Dim oldLatitude  As Double
Dim oldLongitude As Double

Public Sub CEP_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "CEP|LOGRADOURO|BAIRRO|COMPLEMENTO"
  vCampos = "Cep|Logradouro|Bairro|Complemento"
  vHandle = interface.Exec(CurrentSystem, "LOGRADOUROS", vColunas, 2, vCampos, "", "CEP", False, "")

  If vHandle <>0 Then
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Add("SELECT CEP,ESTADO,MUNICIPIO,BAIRRO,LOGRADOURO,COMPLEMENTO   ")
    SQL.Add("  FROM LOGRADOUROS      ")
    SQL.Add(" WHERE HANDLE = :HANDLE ")
    SQL.ParamByName("HANDLE").Value = vHandle
    SQL.Active = True

    CurrentQuery.Edit
    CurrentQuery.FieldByName("CEP").Value = SQL.FieldByName("CEP").AsString
    CurrentQuery.FieldByName("ESTADO").Value = SQL.FieldByName("ESTADO").AsString
    CurrentQuery.FieldByName("MUNICIPIO").Value = SQL.FieldByName("MUNICIPIO").AsString
    CurrentQuery.FieldByName("BAIRRO").Value = SQL.FieldByName("BAIRRO").AsString
    CurrentQuery.FieldByName("LOGRADOURO").Value = SQL.FieldByName("LOGRADOURO").AsString
    CurrentQuery.FieldByName("COMPLEMENTO").Value = SQL.FieldByName("COMPLEMENTO").AsString

  End If

  Set interface = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  oldCep       = CurrentQuery.FieldByName("CEP").AsString
  oldLatitude  = CurrentQuery.FieldByName("LATITUDE").AsFloat
  oldLongitude = CurrentQuery.FieldByName("LONGITUDE").AsFloat
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  If (CampoAlterado("CEP")) Then
    SQL.Clear
    SQL.Add("SELECT HANDLE")
  	SQL.Add("FROM LOGRADOUROS")
  	SQL.Add("WHERE CEP = :CEP")
  	SQL.Add("  AND LOGRADOURO = :LOGRADOURO   ")
  	SQL.Add("  AND (BAIRRO = :BAIRRO ")
  	SQL.Add("   OR (BAIRRO IS NULL AND :BAIRRO = ''))")
  	SQL.Add("  AND (COMPLEMENTO = :COMPLEMENTO ")
  	SQL.Add("   Or (COMPLEMENTO Is Null And :COMPLEMENTO = ''))")
  	SQL.Add("  AND HANDLE <> :HANDLE          ")
  	SQL.ParamByName("CEP").Value = CurrentQuery.FieldByName("CEP").AsString
  	SQL.ParamByName("LOGRADOURO").Value = CurrentQuery.FieldByName("LOGRADOURO").AsString
	SQL.ParamByName("BAIRRO").Value = CurrentQuery.FieldByName("BAIRRO").AsString
	SQL.ParamByName("COMPLEMENTO").Value = CurrentQuery.FieldByName("COMPLEMENTO").AsString
  	SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

    SQL.Active = True

    If Not SQL.EOF Then
      bsShowMessage("CEP já cadastrado", "E")
      CanContinue = False
      Set SQL = Nothing
      Exit Sub
    End If
  End If

  If (CampoAlterado("LATITUDE")) Or (CampoAlterado("LONGITUDE")) Then
    CurrentQuery.FieldByName("DTATUALIZACAOLATITUDELONGITUDE").AsDateTime = ServerDate
  End If

  Set SQL = Nothing
End Sub

Function CampoAlterado(nomeCampo As String) As Boolean
  Select Case nomeCampo
    Case "CEP"
      CampoAlterado = CurrentQuery.FieldByName(nomeCampo).AsString <> oldCep
    Case "LATITUDE"
      CampoAlterado = CurrentQuery.FieldByName(nomeCampo).AsFloat <> oldLatitude
    Case "LONGITUDE"
      CampoAlterado = CurrentQuery.FieldByName(nomeCampo).AsFloat <> oldLongitude
  End Select
End Function
