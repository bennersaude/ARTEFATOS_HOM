﻿'HASH: 8092B7E2E61E14CAC8B6ECD1B36B12E6
'#Uses "*bsShowMessage"
Public Sub TABLE_AfterScroll()
  If VisibleMode Then
    TABLE.Pages("SIMPRO").Visible = CurrentQuery.FieldByName("CODIGO").AsString = "12"
	GRPNORMAL.Visible = (CurrentQuery.FieldByName("CODIGO").AsString <> "12")

    PRECOGENERICO.LocalWhere = "HANDLE NOT IN (@PRECOGENERICO2 , @PRECOGENERICO3 , @PRECOGENERICO4 , @PRECOGENERICOMG2 , @PRECOGENERICOMG3 , @PRECOGENERICOMG4 , @PRECOGENERICOFRACIONADO , @PRECOGENERICOFRACIONADO2 , @PRECOGENERICOFRACIONADO3 , @PRECOGENERICOFRACIONADO4 , @PRECOGENERICOMGFRACIONADO , @PRECOGENERICOMGFRACIONADO2 , @PRECOGENERICOMGFRACIONADO3 , @PRECOGENERICOMGFRACIONADO4)"
    PRECOGENERICO2.LocalWhere = "HANDLE NOT IN (@PRECOGENERICO , @PRECOGENERICO3 , @PRECOGENERICO4 , @PRECOGENERICOMG  , @PRECOGENERICOMG3 , @PRECOGENERICOMG4 , @PRECOGENERICOFRACIONADO , @PRECOGENERICOFRACIONADO2 , @PRECOGENERICOFRACIONADO3 , @PRECOGENERICOFRACIONADO4 , @PRECOGENERICOMGFRACIONADO , @PRECOGENERICOMGFRACIONADO2 , @PRECOGENERICOMGFRACIONADO3 , @PRECOGENERICOMGFRACIONADO4)"
    PRECOGENERICO3.LocalWhere = "HANDLE NOT IN (@PRECOGENERICO , @PRECOGENERICO2 , @PRECOGENERICO4 , @PRECOGENERICOMG  , @PRECOGENERICOMG2 , @PRECOGENERICOMG4 , @PRECOGENERICOFRACIONADO , @PRECOGENERICOFRACIONADO2 , @PRECOGENERICOFRACIONADO3 , @PRECOGENERICOFRACIONADO4 , @PRECOGENERICOMGFRACIONADO , @PRECOGENERICOMGFRACIONADO2 , @PRECOGENERICOMGFRACIONADO3 , @PRECOGENERICOMGFRACIONADO4)"
    PRECOGENERICO4.LocalWhere = "HANDLE NOT IN (@PRECOGENERICO , @PRECOGENERICO2 , @PRECOGENERICO3 , @PRECOGENERICOMG  , @PRECOGENERICOMG2 , @PRECOGENERICOMG3 , @PRECOGENERICOFRACIONADO , @PRECOGENERICOFRACIONADO2 , @PRECOGENERICOFRACIONADO3 , @PRECOGENERICOFRACIONADO4 , @PRECOGENERICOMGFRACIONADO , @PRECOGENERICOMGFRACIONADO2 , @PRECOGENERICOMGFRACIONADO3 , @PRECOGENERICOMGFRACIONADO4)"

    PRECOGENERICOFRACIONADO.LocalWhere = "HANDLE NOT IN (@PRECOGENERICOFRACIONADO2 , @PRECOGENERICOFRACIONADO3 , @PRECOGENERICOFRACIONADO4 , @PRECOGENERICOMG2 , @PRECOGENERICOMG3 , @PRECOGENERICOMG4 , @PRECOGENERICO , @PRECOGENERICO2 , @PRECOGENERICO3 , @PRECOGENERICO4 , @PRECOGENERICOMG , @PRECOGENERICOMGFRACIONADO2 , @PRECOGENERICOMGFRACIONADO3 , @PRECOGENERICOMGFRACIONADO4)"
    PRECOGENERICOFRACIONADO2.LocalWhere = "HANDLE NOT IN (@PRECOGENERICOFRACIONADO , @PRECOGENERICOFRACIONADO3 , @PRECOGENERICOFRACIONADO4 , @PRECOGENERICOMG2 , @PRECOGENERICOMG3 , @PRECOGENERICOMG4 , @PRECOGENERICO , @PRECOGENERICO2 , @PRECOGENERICO3 , @PRECOGENERICO4 , @PRECOGENERICOMG , @PRECOGENERICOMGFRACIONADO  , @PRECOGENERICOMGFRACIONADO3 , @PRECOGENERICOMGFRACIONADO4)"
    PRECOGENERICOFRACIONADO3.LocalWhere = "HANDLE NOT IN (@PRECOGENERICOFRACIONADO , @PRECOGENERICOFRACIONADO2 , @PRECOGENERICOFRACIONADO4 , @PRECOGENERICOMG2 , @PRECOGENERICOMG3 , @PRECOGENERICOMG4 , @PRECOGENERICO , @PRECOGENERICO2 , @PRECOGENERICO3 , @PRECOGENERICO4 , @PRECOGENERICOMG , @PRECOGENERICOMGFRACIONADO  , @PRECOGENERICOMGFRACIONADO2 , @PRECOGENERICOMGFRACIONADO4)"
    PRECOGENERICOFRACIONADO4.LocalWhere = "HANDLE NOT IN (@PRECOGENERICOFRACIONADO , @PRECOGENERICOFRACIONADO2 , @PRECOGENERICOFRACIONADO3 , @PRECOGENERICOMG2 , @PRECOGENERICOMG3 , @PRECOGENERICOMG4 , @PRECOGENERICO , @PRECOGENERICO2 , @PRECOGENERICO3 , @PRECOGENERICO4 , @PRECOGENERICOMG , @PRECOGENERICOMGFRACIONADO  , @PRECOGENERICOMGFRACIONADO2 , @PRECOGENERICOMGFRACIONADO3)"
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("PRECOGENERICO").IsNull Then
    bsShowMessage("O campo Tabela Genérica de Preço é obrigatório!", "I")
    CanContinue = False
    Exit Sub
  End If
End Sub
