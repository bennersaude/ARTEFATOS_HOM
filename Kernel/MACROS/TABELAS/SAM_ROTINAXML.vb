'HASH: DECD53B17052622C6AC48219AA2A7058
'Macro da tabela: SAM_ROTINAXML

Option Explicit

Public Sub TABLE_AfterScroll()

    Dim vbEditavel As Boolean

    vbEditavel = (CurrentQuery.State = 3) Or (CurrentQuery.FieldByName("SITUACAO").AsString = "A") 'Registro em inclusão ou rotina na situação 'Aberta'.

    DESCRICAO.ReadOnly  = Not vbEditavel
    DATAROTINA.ReadOnly = Not vbEditavel

End Sub
