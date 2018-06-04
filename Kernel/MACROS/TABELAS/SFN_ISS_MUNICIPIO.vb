'HASH: 88C950870E35299464081139246BEEF7
'#Uses "*bsShowMessage"


Option Explicit
Dim vMunicipio As Integer, vVigencia As Date

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  vMunicipio = CurrentQuery.FieldByName("MUNICIPIO").AsInteger
  vVigencia = CurrentQuery.FieldByName("VIGENCIA").AsDateTime
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQLMUN As Object
  Set SQLMUN = NewQuery

  SQLMUN.Active = False
  SQLMUN.Clear
  SQLMUN.Add("SELECT HANDLE FROM SFN_ISS_MUNICIPIO WHERE MUNICIPIO = :PMUNICIPIO AND VIGENCIA = :PVIGENCIA AND ISS = :PISS")
  SQLMUN.ParamByName("PMUNICIPIO").AsInteger = CurrentQuery.FieldByName("MUNICIPIO").AsInteger
  SQLMUN.ParamByName("PVIGENCIA").AsDateTime = CurrentQuery.FieldByName("VIGENCIA").AsDateTime
  SQLMUN.ParamByName("PISS").AsInteger = CurrentQuery.FieldByName("ISS").AsInteger
  SQLMUN.Active = True

  If (Not SQLMUN.EOF) And (CurrentQuery.State <> 2) Then
    bsShowMessage("Este Município já está cadastrado para esta mesma vigência de ISS", "E")
    CanContinue = False
    Exit Sub
  End If

  If CurrentQuery.State = 2 Then
    If vVigencia <> CurrentQuery.FieldByName("VIGENCIA").AsDateTime Or vMunicipio <> CurrentQuery.FieldByName("MUNICIPIO").AsInteger Then
      SQLMUN.Active = False
      SQLMUN.Clear
      SQLMUN.Add("SELECT ESTADO, MUNICIPIO, VIGENCIA FROM SFN_ISS_MUNICIPIO WHERE ISS = :PISS")
      SQLMUN.ParamByName("PISS").AsInteger = CurrentQuery.FieldByName("ISS").AsInteger
      SQLMUN.Active = True
      While Not SQLMUN.EOF
        If (CurrentQuery.FieldByName("ESTADO").AsInteger = SQLMUN.FieldByName("ESTADO").AsInteger) And _
             (CurrentQuery.FieldByName("MUNICIPIO").AsInteger = SQLMUN.FieldByName("MUNICIPIO").AsInteger) And _
             (CurrentQuery.FieldByName("VIGENCIA").AsDateTime = SQLMUN.FieldByName("VIGENCIA").AsDateTime) Then
          bsShowMessage("Este Município já está cadastrado para esta mesma vigência de ISS", "E")
          CanContinue = False
          Exit Sub
        End If
        SQLMUN.Next
      Wend
    End If
  End If

End Sub

