'HASH: 7F65D3F5A03ADB04DB1358944701DB85
'Macro: SFN_CONTABREC_REGRA
'#Uses "*bsShowMessage"

Option Explicit

Public Sub CLASSEGERENCIAL_OnPopup(ShowPopup As Boolean)

  Dim interface As Object

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_CLASSEGERENCIAL.ESTRUTURA|SFN_CLASSEGERENCIAL.DESCRICAO|SFN_CLASSEGERENCIAL.CONTACORPORATIVO|SFN_CLASSEGERENCIAL.CODIGOREDUZIDO|SFN_CLASSEGERENCIAL.NATUREZA|SFN_CLASSEGERENCIAL.HISTORICO"

  vCriterio = "HANDLE>0"

  vCampos = "Estrutura|Descrição|Conta Corporativo|Código|D/C|Historico"

  vHandle = interface.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vColunas, 1, vCampos, vCriterio, "Classe Gerencial", False, CLASSEGERENCIAL.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLASSEGERENCIAL").Value = vHandle
  End If
  Set interface = Nothing

End Sub



Public Sub CLASSEGERENCIALAUXILIAR_OnPopup(ShowPopup As Boolean)
  Dim interface As Object

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String


  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_CLASSEGERENCIAL.ESTRUTURA|SFN_CLASSEGERENCIAL.DESCRICAO|SFN_CLASSEGERENCIAL.CODIGOREDUZIDO|SFN_CLASSEGERENCIAL.NATUREZA|SFN_CLASSEGERENCIAL.HISTORICO"

  vCriterio = "HANDLE>0 AND TIPOATO = 'A'"

  vCampos = "Estrutura|Descrição|Código|D/C|Historico"

  vHandle = interface.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vColunas, 1, vCampos, vCriterio, "Classe Gerencial", False, CLASSEGERENCIAL.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLASSEGERENCIALAUXILIAR").Value = vHandle
  End If
  Set interface = Nothing
End Sub


Public Sub TABLE_AfterScroll()
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT SIS.CODIGO")
  SQL.Add("FROM SFN_TIPOLANCFIN SFN, SIS_TIPOLANCFIN SIS")
  SQL.Add("WHERE SFN.HANDLE = :HTIPOLANCFIN")
  SQL.Add("  AND SIS.HANDLE = SFN.TIPOLANCFIN")
  SQL.ParamByName("HTIPOLANCFIN").Value = CurrentQuery.FieldByName("TIPOLANCFIN").AsInteger
  SQL.Active = True

  If SQL.FieldByName("CODIGO").AsInteger = 40 Or _
                     SQL.FieldByName("CODIGO").AsInteger = 30 Then

    SQL.Clear
    SQL.Add("SELECT TABTIPOGESTAO")
    SQL.Add("FROM EMPRESAS")
    SQL.Add("WHERE HANDLE = :HEMPRESA")
    SQL.ParamByName("HEMPRESA").Value = CurrentCompany
    SQL.Active = True

    If SQL.FieldByName("TABTIPOGESTAO").AsInteger = 3 Then
      CLASSEGERENCIALAUXILIAR.Visible = True
    Else
      CLASSEGERENCIALAUXILIAR.Visible = False
    End If
  Else
    CLASSEGERENCIALAUXILIAR.Visible = False
  End If

  Set SQL = Nothing

  If WebMode Then
  	CLASSEGERENCIAL.WebLocalWhere = "HANDLE>0"
  	CLASSEGERENCIALAUXILIAR.WebLocalWhere = "HANDLE>0 And TIPOATO = 'A'"
  	TIPOLANCFIN.WebLocalWhere = "A.TIPOLANCFIN IN (SELECT SIS_TIPOLANCFIN.HANDLE FROM SIS_TIPOLANCFIN WHERE SIS_TIPOLANCFIN.HANDLE = A.TIPOLANCFIN AND SIS_TIPOLANCFIN.CLASSIFICACAORECEBIMENTO = 'S')"
  Else
  	TIPOLANCFIN.LocalWhere = "SFN_TIPOLANCFIN.TIPOLANCFIN IN (SELECT SIS_TIPOLANCFIN.HANDLE FROM SIS_TIPOLANCFIN WHERE SIS_TIPOLANCFIN.HANDLE=SFN_TIPOLANCFIN.TIPOLANCFIN AND SIS_TIPOLANCFIN.CLASSIFICACAORECEBIMENTO = 'S') "
  End If
End Sub
