'HASH: 8ADBA30A10AFAD1EBDDFD3C210CDE8F7
'Macro: SFN_FATURA_LANC_MOD


Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("TABTIPO").AsInteger = 3 Or _
                              CurrentQuery.FieldByName("TABTIPO").AsInteger = 4 Then

    Dim SQL As Object
    Set SQL = NewQuery

    SQL.Clear
    SQL.Add("SELECT EMP.TABTIPOGESTAO")
    SQL.Add("FROM SAM_BENEFICIARIO B, EMPRESAS EMP")
    SQL.Add("WHERE B.HANDLE = :HBENEFICIARIO")
    SQL.Add("  AND EMP.HANDLE = B.EMPRESA")

    SQL.ParamByName("HBENEFICIARIO").Value = RecordHandleOfTable("SAM_BENEFICIARIO")
    SQL.Active = True

    If SQL.FieldByName("TABTIPOGESTAO").AsInteger = 3 Then
      VALORCLASSEPRINCIPAL.Visible = True
      CLASSEGERENCIALAUXILIAR.Visible = True
      VALORCLASSEAUXILIAR.Visible = True
    Else
      VALORCLASSEPRINCIPAL.Visible = False
      CLASSEGERENCIALAUXILIAR.Visible = False
      VALORCLASSEAUXILIAR.Visible = False
    End If
  Else
    VALORCLASSEPRINCIPAL.Visible = False
    CLASSEGERENCIALAUXILIAR.Visible = False
    VALORCLASSEAUXILIAR.Visible = False
  End If
End Sub

Public Sub TABTIPO_OnChanging(AllowChange As Boolean)
  AllowChange = False
End Sub

