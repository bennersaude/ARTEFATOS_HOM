'HASH: 20472F0DA347D8EE22F1E53107E01518
'#Uses "*bsShowMessage"
'SFN_TIPOLANCFIN
Option Explicit

Public Sub CLASSEGERENCIAL_OnPopup(ShowPopup As Boolean)

  Dim qTipoLancFin As Object
  Set qTipoLancFin = NewQuery

  qTipoLancFin.Add("SELECT CODIGO                ")
  qTipoLancFin.Add("  FROM SIS_TIPOLANCFIN       ")
  qTipoLancFin.Add(" WHERE HANDLE = :TIPOLANCFIN ")
  qTipoLancFin.ParamByName("TIPOLANCFIN").Value = CurrentQuery.FieldByName("TIPOLANCFIN").AsInteger
  qTipoLancFin.Active = True

  If qTipoLancFin.FieldByName("CODIGO").AsInteger = 400 Then

    Dim qParam As Object
    Set qParam = NewQuery

    qParam.Clear
    qParam.Add("SELECT CONTABILIZA FROM SFN_PARAMETROSFIN")
    qParam.Active = True

    If qParam.FieldByName("CONTABILIZA").AsString = "S" Then
      CLASSEGERENCIAL.LocalWhere = "SFN_CLASSEGERENCIAL.IMPEDIRLANCAMENTOMANUAL = 'N' "
    Else
      CLASSEGERENCIAL.LocalWhere = ""
    End If

    Set qParam = Nothing

  End If
  Set qTipoLancFin = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'SMS: 6660  -   25/02/2002 Keila/Marcio
  Dim Sql As Object
  Set Sql = NewQuery

  Sql.Add("SELECT HANDLE, CLASSEGERENCIAL FROM SFN_TIPOLANCFIN")
  Sql.Add(" WHERE HANDLE IN (SELECT SF.HANDLE                 ")
  Sql.Add("                   FROM SFN_TIPOLANCFIN SF,        ")
  Sql.Add("                        SIS_TIPOLANCFIN SI         ")
  Sql.Add("                  WHERE SF.TIPOLANCFIN = SI.HANDLE ")
  Sql.Add("                    AND SI.CODIGO = 400)           ")
  Sql.Add("  AND HANDLE = :Tipo                               ")
  Sql.ParamByName("TIPO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Sql.Active = True

  If CurrentQuery.FieldByName("CLASSEGERENCIAL").IsNull And Sql.FieldByName("HANDLE").AsInteger >0 Then
    If CLASSEGERENCIAL.Visible = True Then
      bsShowMessage("Campo Classe Gerencial é obrigatório", "E")
      CanContinue = False
    End If
  End If

End Sub

