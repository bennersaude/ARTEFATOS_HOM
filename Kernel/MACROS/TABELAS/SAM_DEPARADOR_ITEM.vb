'HASH: 6019960E57781CF07EFDC60A56A953A7
'Macro da tabela SAM_DEPARADOR_ITEM

Option Explicit

Dim giHandleEventoAnterior   As Long
Dim giHandleGrauAnterior     As Long
Dim giHandleUsuarioDeparador As Long
Dim gdDataDeparacao          As Date

Public Sub TABLE_AfterPost

    Dim qInsert                As Object
    Dim qUpdate                As Object
    Dim qSelect                As BPesquisa
    Dim vbFoiCorrigido         As Boolean
    Dim vbPossuiCodigoRestrito As Boolean
    Dim vbEventoDuplicado      As Boolean
    Dim viCodigoCliente        As Integer

    Set qInsert     = NewQuery
    Set qUpdate     = NewQuery
    Set qSelect     = NewQuery

    vbPossuiCodigoRestrito = False
    vbEventoDuplicado      = False
    vbFoiCorrigido         = False

    qSelect.Clear
    qSelect.Add("SELECT CODCLIENTE            ")
    qSelect.Add("  FROM GTO_CLIENTE           ")
    qSelect.Add(" WHERE PARAMETRO1S = :VIGENTE")

    qSelect.ParamByName("VIGENTE").AsString = "CLIENTE VIGENTE"

    qSelect.Active = True

    If (Not qSelect.EOF) Then
       viCodigoCliente = qSelect.FieldByName("CODCLIENTE").AsInteger
    Else
       InfoDescription = "Não existe um cliente parametrizado como 'CLIENTE VIGENTE' no sistema."
    End If

    qSelect.Active = False

    'Verifica se o evento é um código restrito...
    qSelect.Clear
    qSelect.Add("SELECT PARAMETRO2S,                                                ")
    qSelect.Add("       PARAMETRO3S                                                 ")
    qSelect.Add("  FROM GTO_CLIENTE                                                 ")
    qSelect.Add(" WHERE PARAMETRO1S = :CODIGORESTRITO                               ")
    qSelect.Add("   AND (CODCLIENTE = 0 OR CODCLIENTE = :CODIGOCLIENTE)             ")
    qSelect.Add("   AND (PARAMETRO2S = :CODIGOEVENTO OR PARAMETRO3S = :CODIGOEVENTO)")

    qSelect.ParamByName("CODIGORESTRITO").AsString  = "DEP_CODREST"
    qSelect.ParamByName("CODIGOCLIENTE" ).AsInteger = viCodigoCliente
    qSelect.ParamByName("CODIGOEVENTO"  ).AsString  = CurrentQuery.FieldByName("CODIGOEVENTOIMPORTADO").AsString

    qSelect.Active = True

    If (Not qSelect.EOF) Then
        vbPossuiCodigoRestrito = True
    Else

        qSelect.Active = False

        'Verifica se o evento é duplicado na TGE...
        qSelect.Clear
        qSelect.Add("SELECT COUNT(1) QTD                          ")
        qSelect.Add("  FROM SAM_TGE                               ")
        qSelect.Add(" WHERE ULTIMONIVEL       = :SIM              ")
        qSelect.Add("   AND ESTRUTURANUMERICA = :ESTRUTURANUMERICA")

        qSelect.ParamByName("SIM"              ).AsString = "S"
        qSelect.ParamByName("ESTRUTURANUMERICA").AsString = CurrentQuery.FieldByName("CODIGOEVENTOIMPORTADO").AsString

        qSelect.Active = True

        vbEventoDuplicado = qSelect.FieldByName("QTD").AsInteger > 1

    End If

    If (WebVisionCode = "ITENSVAL") Then 'Estamos na visão do validador de eventos, então...

        'Verifica se o evento foi corrigido pelo validador...
        If ((giHandleEventoAnterior <> CurrentQuery.FieldByName("EVENTO").AsInteger) Or (giHandleGrauAnterior <> CurrentQuery.FieldByName("GRAU").AsInteger)) Then
            vbFoiCorrigido = True
        End If

        If ((CurrentQuery.FieldByName("DEPARACAO").AsInteger = 2) And (Not vbPossuiCodigoRestrito) And (Not vbEventoDuplicado)) Then 'Deparação definitiva por prestador.

            Dim qTgePrestador As BPesquisa
            Set qTgePrestador = NewQuery

            qTgePrestador.Add("SELECT HANDLE                                                       ")
            qTgePrestador.Add("  FROM SAM_TGE_PRESTADOR                                            ")
            qTgePrestador.Add(" WHERE COALESCE(EVENTOPRESTADOR, ' ') = COALESCE(:CODIGOEVENTO, ' ')")
            qTgePrestador.Add("   AND PRESTADOR = (SELECT PRESTADOR                                ")
            qTgePrestador.Add("                      FROM SAM_DEPARADOR_PRESTADORES                ")
            qTgePrestador.Add("                     WHERE HANDLE = :HANDLEPRESTADOR)               ")

            qTgePrestador.ParamByName("CODIGOEVENTO"   ).AsString  = CurrentQuery.FieldByName("CODIGOEVENTOIMPORTADO").AsString
            qTgePrestador.ParamByName("HANDLEPRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger

            qTgePrestador.Active = True

            If (Not qTgePrestador.EOF) Then

                qUpdate.Clear
                qUpdate.Add("UPDATE SAM_TGE_PRESTADOR                                     ")
                qUpdate.Add("   SET DESCRICAO       = :DESCRICAO,                         ")
                qUpdate.Add("       EVENTOPRESTADOR = :CODIGOEVENTO,                      ")
                qUpdate.Add("       GRAU            = (SELECT GRAU                        ")
                qUpdate.Add("                            FROM SAM_TGE_GRAU                ")
                qUpdate.Add("                           WHERE HANDLE = :HANDLEGRAUVALIDO),")
                qUpdate.Add("       EVENTO          = :HANDLEEVENTO                       ")
                qUpdate.Add(" WHERE HANDLE = :HANDLE                                      ")

                qUpdate.ParamByName("DESCRICAO"       ).AsString  = Mid(CurrentQuery.FieldByName("DESCRICAOEVENTOIMPORTADO").AsString, 1, 50)
                qUpdate.ParamByName("CODIGOEVENTO"    ).AsString  = CurrentQuery.FieldByName("CODIGOEVENTOIMPORTADO").AsString
                qUpdate.ParamByName("HANDLEEVENTO"    ).AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
                qUpdate.ParamByName("HANDLEGRAUVALIDO").AsInteger = CurrentQuery.FieldByName("GRAU").AsInteger
                qUpdate.ParamByName("HANDLE"          ).AsInteger = qTgePrestador.FieldByName("HANDLE").AsInteger

                qUpdate.ExecSQL

            Else

                qInsert.Clear
                qInsert.Add("INSERT INTO SAM_TGE_PRESTADOR (                                         ")
                qInsert.Add("            HANDLE,                                                     ")
                qInsert.Add("            EVENTOPRESTADOR,                                            ")
                qInsert.Add("            EVENTO,                                                     ")
                qInsert.Add("            DESCRICAO,                                                  ")
                qInsert.Add("            PRESTADOR,                                                  ")
                qInsert.Add("            GRAU)                                                       ")
                qInsert.Add("(SELECT :NOVOHANDLE,                                                    ")
                qInsert.Add("        :CODIGOEVENTO,                                                  ")
                qInsert.Add("        :HANDLEEVENTO,                                                  ")
                qInsert.Add("        :DESCRICAOEVENTO,                                               ")
                qInsert.Add("        PRESTADOR,                                                      ")
                qInsert.Add("        (SELECT GRAU FROM SAM_TGE_GRAU WHERE HANDLE = :HANDLEGRAUVALIDO)")
                qInsert.Add("   FROM SAM_DEPARADOR_PRESTADORES                                       ")
                qInsert.Add("  WHERE HANDLE = :HANDLE)                                               ")

                qInsert.ParamByName("NOVOHANDLE"      ).AsInteger = NewHandle("SAM_TGE_PRESTADOR")
                qInsert.ParamByName("CODIGOEVENTO"    ).AsString = CurrentQuery.FieldByName("CODIGOEVENTOIMPORTADO").AsString
                qInsert.ParamByName("HANDLEEVENTO"    ).AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
                qInsert.ParamByName("DESCRICAOEVENTO" ).AsString = Mid(CurrentQuery.FieldByName("DESCRICAOEVENTOIMPORTADO").AsString,1,50)
                qInsert.ParamByName("HANDLEGRAUVALIDO").AsInteger = CurrentQuery.FieldByName("GRAU").AsInteger
                qInsert.ParamByName("HANDLE"          ).AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger

                qInsert.ExecSQL

            End If

            qTgePrestador.Active = False
            Set qTgePrestador = Nothing

        End If

    End If

    AtualizarDeparador _
        piHandleDeparador          := CurrentQuery.FieldByName("HANDLE").AsInteger, _
        piHandleDeparadorPrestador := CurrentQuery.FieldByName("PRESTADOR").AsInteger, _
        pbFoiCorrigido             := vbFoiCorrigido

    'Após a deparação do item, verifica se existe o mesmo evento para deparações futuras do mesmo prestador...
    If (CurrentQuery.FieldByName("DEPARACAO").AsInteger = 2) Then 'Deparação definitiva por prestador.

        Dim qItensGeral As BPesquisa
        Set qItensGeral = NewQuery

        qItensGeral.Clear
        qItensGeral.Add("SELECT A.HANDLE,                                             ")
        qItensGeral.Add("       A.PRESTADOR,                                          ")
        qItensGeral.Add("       A.DESCRICAOEVENTIMPORT,                               ")
        qItensGeral.Add("       A.EVENTOIMPORTADO,                                    ")
        qItensGeral.Add("       A.CODIGOTABELA                                        ")
        qItensGeral.Add("  FROM SAM_DEPARADOR_ITEM        A                           ")
        qItensGeral.Add("  JOIN SAM_DEPARADOR_PRESTADORES B ON (B.HANDLE = A.PRESTADOR")
        qItensGeral.Add("  JOIN SAM_PRESTADOR             C ON (C.HANDLE = B.PRESTADOR")

        If (vbFoiCorrigido) Then

            qItensGeral.Add(" WHERE A.EVENTO       = :HANDLEEVENTOANTERIOR  ")
            qItensGeral.Add("   AND A.USUARIO      = :HANDLEUSUARIODEPARADOR")
            qItensGeral.Add("   AND A.FALTAVALIDAR = :SIM                   ")


            qItensGeral.ParamByName("HANDLEEVENTOANTERIOR"  ).AsInteger = giHandleEventoAnterior
            qItensGeral.ParamByName("HANDLEUSUARIODEPARADOR").AsInteger = giHandleUsuarioDeparador
            qItensGeral.ParamByName("SIM"                   ).AsString  = "S"
        Else

            qItensGeral.Add(" WHERE A.EVENTO IS NULL")

        End If

        qItensGeral.Add("   AND COALESCE(A.DESCRICAOEVENTIMPORTADO, ' ') = COALESCE(:DESCRICAOEVENTO, ' ')")
        qItensGeral.Add("   AND COALESCE(A.CODIGOTABELAIMPORTADO, ' ')   = COALESCE(:CODIGOTABELA, ' ')   ")
        qItensGeral.Add("   AND COALESCE(A.CODIGOEVENTIMPORTADO, ' ')    = COALESCE(:CODIGOEVENTO, ' ')   ")
        qItensGeral.Add("   AND C.HANDLE = (SELECT PRESTADOR                                              ")
        qItensGeral.Add("                     FROM SAM_DEPARADOR_PRESTADORES                              ")
        qItensGeral.Add("                    WHERE HANDLE = :HANDLEPRESTADOR)                             ")

        qItensGeral.ParamByName("DESCRICAOEVENTO" ).AsString  = CurrentQuery.FieldByName("DESCRICAOEVENTOIMPORTADO").AsString
        qItensGeral.ParamByName("CODIGOTABELA"    ).AsString  = CurrentQuery.FieldByName("CODIGOTABELAIMPORTADO").AsString
        qItensGeral.ParamByName("CODIGOEVENTO"    ).AsString  = CurrentQuery.FieldByName("CODIGOEVENTOIMPORTADO").AsString
        qItensGeral.ParamByName("HANDLEPRESTADOR" ).AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger

        qItensGeral.Active = True

        While (Not qItensGeral.EOF)

            If (qItensGeral.FieldByName("HANDLE").AsInteger <> CurrentQuery.FieldByName("HANDLE").AsInteger) Then
                AtualizarDeparador _
                    piHandleDeparador          := qItensGeral.FieldByName("HANDLE").AsInteger, _
                    piHandleDeparadorPrestador := qItensGeral.FieldByName("PRESTADOR").AsInteger, _
                    pbFoiCorrigido             := vbFoiCorrigido
            End If

            qItensGeral.Next

        Wend

        qItensGeral.Active = False
        Set qItensGeral = Nothing

    End If

    If (CurrentQuery.FieldByName("DEPARACAO").AsInteger = 2) And (vbPossuiCodigoRestrito) Then 'Deparação definitiva por prestador.
        InfoDescription = InfoDescription + " A deparação foi feita, porém a REGRA NÃO FOI SALVA. O código '" + CurrentQuery.FieldByName("CODIGOEVENTOIMPORTADO").AsString + "' é restrito."
    End If

    If (CurrentQuery.FieldByName("DEPARACAO").AsInteger = 2) And (vbEventoDuplicado) Then 'Deparação definitiva por prestador.
        InfoDescription = InfoDescription + " A deparação foi feita, porém a REGRA NÃO FOI SALVA. O código '" + CurrentQuery.FieldByName("CODIGOEVENTOIMPORTADO").AsString + "' está duplicado na tabela de eventos."
    End If

    qSelect.Active = False

    Set qInsert = Nothing
    Set qUpdate = Nothing
    Set qSelect = Nothing

End Sub

Sub AtualizarDeparador(piHandleDeparador As Long, piHandleDeparadorPrestador As Long, pbFoiCorrigido As Boolean)

    Dim qUpdate As Object
    Set qUpdate = NewQuery

    qUpdate.Clear
    qUpdate.Add("UPDATE SAM_DEPARADOR_ITEM                  ")
    qUpdate.Add("   SET EVENTO              = :HANDLEEVENTO,")

    If (WebVisionCode = "ITENSVAL") Then 'Estamos na visão do validador de eventos, então...

        qUpdate.Add("       FALTAVALIDAR        = :NAO,                  ")
        qUpdate.Add("       USUARIO             = :HANDLEUSUARIO,        ")
        qUpdate.Add("       DATAHORAINCLUSAO    = :DATAHORADEPARACAO,    ")
        qUpdate.Add("       DATAHORAVALIDACAO   = :AGORA,                 ")
        qUpdate.Add("       USUARIOVALIDACAO    = :HANDLEUSUARIOCORRENTE,")

        If (Not pbFoiCorrigido) Then

            qUpdate.Add("       NIVELERROVALIDACAO  = NULL,")
            qUpdate.Add("       OBSERVACAOVALIDACAO = NULL ")

        Else

            qUpdate.Add("       NIVELERROVALIDACAO  = :NIVELERRO,          ")
            qUpdate.Add("       OBSERVACAOVALIDACAO = :OBSERVACOES,        ")
            qUpdate.Add("       EVENTOANTERIOR      = :HANDLEEVENTOANTERIOR")

            If ((giHandleGrauAnterior > 0) And (giHandleGrauAnterior <> CurrentQuery.FieldByName("GRAU").AsInteger)) Then
                qUpdate.Add("       ,GRAUANTERIOR      = :HANDLEGRAUANTERIOR")
                qUpdate.ParamByName("HANDLEGRAUANTERIOR").AsInteger = giHandleGrauAnterior
            End If

            qUpdate.ParamByName("NIVELERRO"           ).AsInteger = CurrentQuery.FieldByName("NIVELERROVALIDACAO").AsInteger
            qUpdate.ParamByName("OBSERVACOES"         ).AsString  = CurrentQuery.FieldByName("OBSERVACAOVALIDACAO").AsString
            qUpdate.ParamByName("HANDLEEVENTOANTERIOR").AsInteger = giHandleEventoAnterior

        End If

        qUpdate.Add(" WHERE HANDLE = :HANDLE")

        qUpdate.ParamByName("NAO"                  ).AsString   = "N"
        qUpdate.ParamByName("HANDLEUSUARIO"        ).AsInteger  = giHandleUsuarioDeparador
        qUpdate.ParamByName("DATAHORADEPARACAO"    ).AsDateTime = gdDataDeparacao
        qUpdate.ParamByName("HANDLEUSUARIOCORRENTE").AsInteger  = CurrentUser

    Else

        qUpdate.Add("       FALTAVALIDAR        = :SIM,                  ")
        qUpdate.Add("       USUARIO             = :HANDLEUSUARIOCORRENTE,")
        qUpdate.Add("       DATAHORAINCLUSAO    = :AGORA                 ")


        qUpdate.ParamByName("SIM"                  ).AsString  = "S"
        qUpdate.ParamByName("HANDLEUSUARIOCORRENTE").AsInteger = CurrentUser

    End If

    qUpdate.ParamByName("HANDLEEVENTO").AsInteger  = CurrentQuery.FieldByName("EVENTO").AsInteger
    qUpdate.ParamByName("AGORA"       ).AsDateTime = ServerNow
    qUpdate.ParamByName("HANDLE"      ).AsInteger  = piHandleDeparador

    qUpdate.ExecSQL

    If (WebVisionCode = "ITENSVAL") Then 'Estamos na visão do validador de eventos, então...

        AtualizarEventos piHandleDeparadorPrestador, _
                         CurrentQuery.FieldByName("EVENTO").AsInteger, _
                         piHandleDeparador, _
                         False, _
                         0

    End If

    Set qUpdate = Nothing

End Sub

Sub AtualizarEventos(piHandleDeparadorPrestador As Long, piHandleEvento As Long, piHandleDeparador As Long, pbEhPorEstado As Boolean, piHandleEstado As Long)

    Dim qUpdate As Object
    Dim qSelect As BPesquisa

    Set qUpdate = NewQuery
    Set qSelect = NewQuery

    qUpdate.Clear
    qUpdate.Add("UPDATE SAM_DEPARADOR_PRESTADORES                                                                                                           ")
    qUpdate.Add("   SET QUANTIDADEITENS     = (SELECT COUNT(1)                                                                                              ")
    qUpdate.Add("                                FROM SAM_DEPARADOR_ITEM B                                                                                  ")
    qUpdate.Add("                               WHERE B.PRESTADOR = HANDLE                                                                                  ")
    qUpdate.Add("                                 AND B.EVENTO IS NULL),                                                                                    ")
    qUpdate.Add("       VALORPAGAR          = (SELECT SUM(C.VALORINFORMADO)                                                                                 ")
    qUpdate.Add("                                FROM SAM_DEPARADOR_ITEM C                                                                                  ")
    qUpdate.Add("                               WHERE C.PRESTADOR = HANDLE                                                                                  ")
    qUpdate.Add("                                 AND C.EVENTO IS NULL),                                                                                    ")
    qUpdate.Add("       QUANTIDADEEVENTOS   = (SELECT SUM(D.QTDOCORRENCIAS)                                                                                 ")
    qUpdate.Add("                                FROM SAM_DEPARADOR_ITEM D                                                                                  ")
    qUpdate.Add("                               WHERE D.PRESTADOR = HANDLE                                                                                  ")
    qUpdate.Add("                                 AND D.EVENTO IS NULL),                                                                                    ")
    qUpdate.Add("       DATAMAXIMAPAGAMENTO = (SELECT MAX(J.DATAPAGAMENTO)                                                                                  ")
    qUpdate.Add("                                FROM SAM_GUIA_EVENTOS EVE                                                                                  ")
    qUpdate.Add("                                 JOIN SAM_GUIA         I ON (I.HANDLE = EVE.GUIA)                                                          ")
    qUpdate.Add("                                 JOIN SAM_PEG          J ON (J.HANDLE = I.PEG                                                              ")
    qUpdate.Add("                                  WHERE EVE.DEPARADOR IN (SELECT U.HANDLE                                                                  ")
    qUpdate.Add("                                                         FROM SAM_DEPARADOR_ITEM U                                                         ")
    qUpdate.Add("                                                        WHERE U.PRESTADOR = :HANDLEDEPARADORPRESTADOR)),                                   ")
    qUpdate.Add("       DATAMINIMAPAGAMENTO = (SELECT MIN(J.DATAPAGAMENTO)                                                                                  ")
    qUpdate.Add("                                FROM SAM_GUIA_EVENTOS EVE                                                                                  ")
    qUpdate.Add("                                 JOIN SAM_GUIA         I ON (I.HANDLE = EVE.GUIA)                                                          ")
    qUpdate.Add("                                 JOIN SAM_PEG          J ON (J.HANDLE = I.PEG)                                                             ")
    qUpdate.Add("                                  WHERE EVE.DEPARADOR IN (SELECT U.HANDLE                                                                  ")
    qUpdate.Add("                                                         FROM SAM_DEPARADOR_ITEM U                                                         ")
    qUpdate.Add("                                                        WHERE U.PRESTADOR = :HANDLEDEPARADORPRESTADOR)),                                   ")
    qUpdate.Add("       QUANTIDADEGUIAS     = (SELECT COUNT(1)                                                                                              ")
    qUpdate.Add("                                FROM SAM_GUIA E                                                                                            ")
    qUpdate.Add("                               WHERE E.HANDLE IN (SELECT F.GUIA                                                                            ")
    qUpdate.Add("                                                    FROM SAM_GUIA_EVENTOS F                                                                ")
    qUpdate.Add("                                                   WHERE F.DEPARADOR IN (SELECT G.HANDLE                                                   ")
    qUpdate.Add("                                                                           FROM SAM_DEPARADOR_ITEM G                                       ")
    qUpdate.Add("                                                                          WHERE G.PRESTADOR = :HANDLEDEPARADORPRESTADOR                    ")
    qUpdate.Add("                                                                            AND G.EVENTO IS NULL)                                          ")
    qUpdate.Add("                                                  GROUP BY F.GUIA)),                                                                       ")
    qUpdate.Add("       QUANTIDADEPEGS      = (SELECT COUNT(1)                                                                                              ")
    qUpdate.Add("                                FROM SAM_PEG X                                                                                             ")
    qUpdate.Add("                               WHERE X.HANDLE IN (SELECT Y.PEG                                                                             ")
    qUpdate.Add("                                                    FROM SAM_GUIA Y                                                                        ")
    qUpdate.Add("                                                   WHERE Y.HANDLE IN (SELECT Z.GUIA                                                        ")
    qUpdate.Add("                                                                        FROM SAM_GUIA_EVENTOS Z                                            ")
    qUpdate.Add("                                                                       WHERE Z.DEPARADOR IN (SELECT W.HANDLE                               ")
    qUpdate.Add("                                                                                               FROM SAM_DEPARADOR_ITEM W                   ")
    qUpdate.Add("                                                                                              WHERE W.PRESTADOR = :HANDLEDEPARADORPRESTADOR")
    qUpdate.Add("                                                                                                AND W.EVENTO IS NULL)                      ")
    qUpdate.Add("                                                                      GROUP BY Z.GUIA)                                                     ")
    qUpdate.Add("                                                  GROUP BY Y.PEG))                                                                         ")
    qUpdate.Add(" WHERE HANDLE = :HANDLEDEPARADORPRESTADOR                                                                                                  ")

    If (pbEhPorEstado) Then

        qUpdate.Add("   AND PRESTADOR IN (SELECT E.HANDLE                                            ")
        qUpdate.Add("                       FROM SAM_PRESTADOR E                                     ")
        qUpdate.Add("                      WHERE E.FILIALPADRAO IN (SELECT F.HANDLE                  ")
        qUpdate.Add("                                                 FROM FILIAIS F                 ")
        qUpdate.Add("                                                WHERE F.ESTADO = :HANDLEESTADO))")

        qUpdate.ParamByName("HANDLEESTADO").AsInteger = piHandleEstado

    End If

    qUpdate.ParamByName("HANDLEDEPARADORPRESTADOR").AsInteger = piHandleDeparadorPrestador

    qUpdate.ExecSQL

    qSelect.Clear
    qSelect.Add("SELECT P.HANDLE,                                          ")
    qSelect.Add("       S.GRAUPRINCIPAL                                    ")
    qSelect.Add("  FROM TIS_TABELAPRECO    P                               ")
    qSelect.Add("  JOIN SAM_TGE_TABELATISS T ON (T.TABELATISS = P.HANDLE)  ")
    qSelect.Add("  JOIN SAM_TGE            S ON (S.HANDLE = T.EVENTO)      ")
    qSelect.Add("  JOIN SAM_ORIGEMEVENTO   O ON (O.HANDLE = S.ORIGEMEVENTO)")
    qSelect.Add(" WHERE S.HANDLE = :HANDLEEVENTO                           ")
    qSelect.Add("   AND P.VERSAOTISS IN (SELECT MAX(V.HANDLE)              ")
    qSelect.Add("                          FROM TIS_VERSAO V               ")
    qSelect.Add("                         WHERE V.ATIVODESKTOP = 'S')      ")

    qSelect.ParamByName("HANDLEEVENTO").AsInteger = piHandleEvento

    qSelect.Active = True

    qUpdate.Clear
    qUpdate.Add("UPDATE SAM_GUIA_EVENTOS      ")
    qUpdate.Add("   SET EVENTO = :HANDLEEVENTO")

    If (Not qSelect.EOF) Then
        qUpdate.Add("       ,CODIGOTABELA = :HANDLETABELAPRECO")
        qUpdate.ParamByName("HANDLETABELAPRECO").AsInteger = qSelect.FieldByName("HANDLE").AsInteger
    End If

    If (Not CurrentQuery.FieldByName("GRAU").IsNull) Then
         qUpdate.Add("       ,GRAU = (SELECT E.GRAU                 ")
         qUpdate.Add("                  FROM SAM_TGE_GRAU E         ")
         qUpdate.Add("                 WHERE E.HANDLE = :HANDLEGRAU)")
         qUpdate.ParamByName("HANDLEGRAU").AsInteger = CurrentQuery.FieldByName("GRAU").AsInteger
    Else
         qUpdate.Add("       ,GRAU = :HANDLEGRAU")
         qUpdate.ParamByName("HANDLEGRAU").AsInteger = qSelect.FieldByName("GRAUPRINCIPAL").AsInteger
    End If

    qUpdate.Add(" WHERE DEPARADOR = :HANDLEDEPARADOR                     ")
    qUpdate.Add("   AND GUIA IN (SELECT B.HANDLE                         ")
    qUpdate.Add("                  FROM SAM_GUIA B                       ")
    qUpdate.Add("                 WHERE B.PEG IN (SELECT C.HANDLE        ")
    qUpdate.Add("                                   FROM SAM_PEG C       ")
    qUpdate.Add("                                  WHERE C.SITUACAO = 1))")

    qUpdate.ParamByName("HANDLEEVENTO"   ).AsInteger = piHandleEvento
    qUpdate.ParamByName("HANDLEDEPARADOR").AsInteger = piHandleDeparador

    qUpdate.ExecSQL

    AtualizarStatus(piHandleDeparadorPrestador)

    qSelect.Active = False

    Set qUpdate = Nothing
    Set qSelect = Nothing

End Sub

Sub AtualizarStatus(piHandleDeparadorPrestador As Long)

    Dim qSelect As BPesquisa
    Dim qUpdate As Object

    Set qSelect = NewQuery
    Set qUpdate = NewQuery

    qSelect.Clear
    qSelect.Add("SELECT A.HANDLE                                               ")
    qSelect.Add("  FROM SAM_DEPARADOR_PRESTADORES A                            ")
    qSelect.Add("  JOIN SAM_DEPARADOR_ITEM        B ON (B.PRESTADOR = A.HANDLE)")
    qSelect.Add("  WHERE A.HANDLE = :HANDLEDEPARADORPRESTADOR                  ")
    qSelect.Add("    AND B.EVENTO IS NULL                                      ")

    qSelect.ParamByName("HANDLEDEPARADORPRESTADOR").AsInteger = piHandleDeparadorPrestador

    qSelect.Active = True

    If (qSelect.EOF) Then

        qUpdate.Clear
        qUpdate.Add("UPDATE SAM_DEPARADOR_PRESTADORES")
        qUpdate.Add("   SET ANALISADO = :SIM         ")
        qUpdate.Add(" WHERE HANDLE = :HANDLE         ")

        qUpdate.ParamByName("SIM"   ).AsString  = "S"
        qUpdate.ParamByName("HANDLE").AsInteger = piHandleDeparadorPrestador

        qUpdate.ExecSQL

    End If

    qSelect.Active = False

    Set qSelect = Nothing
    Set qUpdate = Nothing
End Sub

Public Sub TABLE_AfterScroll

    If (WebVisionCode = "ITENSVAL") Then 'Estamos na visão do validador de eventos, então...

        giHandleEventoAnterior   = CurrentQuery.FieldByName("EVENTO").AsInteger
        giHandleGrauAnterior     = CurrentQuery.FieldByName("GRAU").AsInteger
        giHandleUsuarioDeparador = CurrentQuery.FieldByName("USUARIO").AsInteger
        gdDataDeparacao          = CurrentQuery.FieldByName("DATAHORAINCLUSAO").AsDateTime

    End If

End Sub


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

    If (WebVisionCode = "DEPARADORITENS") Then

        Dim qSelect As BPesquisa
        Set qSelect = NewQuery

        qSelect.Clear
        qSelect.Add("SELECT USUARIORESPONSAVEL       ")
        qSelect.Add("  FROM SAM_DEPARADOR_PRESTADORES")
        qSelect.Add(" WHERE HANDLE = :HANDLE         ")

        qSelect.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger

        qSelect.Active = True

        If (qSelect.FieldByName("USUARIORESPONSAVEL").AsInteger <> CurrentUser) Then

            CancelDescription = "Você não é o responsável por esta demanda."
            CanContinue = False

            qSelect.Active = False
            Set qSelect = Nothing

            Exit Sub

        End If

        qSelect.Active = False
        Set qSelect = Nothing

    End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

    Dim vfValorAtual          As Currency
    Dim vfValorDeparado       As Currency
    Dim vfDiferencaReal       As Currency
    Dim vfDiferencaPercentual As Currency
    Dim qSelect               As Object

    Set qSelect = NewQuery

    qSelect.Clear
    qSelect.Add("SELECT PRECODEPARACAO  ")
    qSelect.Add("  FROM SAM_TGE WHERE   ")
    qSelect.Add(" WHERE HANDLE = :HANDLE")

    qSelect.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger

    qSelect.Active = True

    vfValorAtual    = CurrentQuery.FieldByName("VALORMEDIO").AsCurrency
    vfValorDeparado = qSelect.FieldByName("PRECODEPARACAO").AsCurrency

    vfDiferencaReal = vfValorDeparado - vfValorAtual

    If (vfValorAtual = 0) Then
        vfDiferencaPercentual = 0
    Else
        vfDiferencaPercentual = Round(((vfDiferencaReal * 100) / vfValorAtual), 2)
    End If

    If (CurrentQuery.FieldByName("DEPARACAO").IsNull)  Then
        CurrentQuery.FieldByName("DEPARACAO").AsString = "1" 'Deparação somente neste caso.
    End If

    If ((vfDiferencaPercentual >= 50) Or (vfDiferencaPercentual <= -50)) Then 'Mais que 50%, só apresentar mensagem.

        If (RequestConfirmation ("O valor deparado tem diferença de " + CStr(vfDiferencaPercentual) +"%, Confirma a operação de deparação?")) Then
            'InfoDescription = "Deparado com sucesso!"
        End If

    End If

    If (WebVisionCode = "ITENSREP") Then 'Estamos na visão da edição para deparação, então...

        qSelect.Active = False

        qSelect.Clear
        qSelect.Add("SELECT P.PEG,                                   ")
        qSelect.Add("       P.SITUACAO                               ")
        qSelect.Add("  FROM SAM_GUIA_EVENTOS E                       ")
        qSelect.Add("  JOIN SAM_GUIA         G ON (G.HANDLE = E.GUIA)")
        qSelect.Add("  JOIN SAM_PEG          P ON (P.HANDLE = G.PEG) ")
        qSelect.Add(" WHERE E.DEPARADOR = :HANDLEDEPARADOR           ")

        qSelect.ParamByName("HANDLEDEPARADOR").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

        qSelect.Active = True

        If (qSelect.FieldByName("SITUACAO").AsString <> "1") Then 'PEG em digitação.

            CancelDescription = "Este evento não pode ser mais deparado, pois o PEG " + qSelect.FieldByName("PEG").AsString + " já não está mais em digitação."
            CanContinue = False

            qSelect.Active = False
            Set qSelect = Nothing

            Exit Sub

        End If

    End If

    qSelect.Active = False
    Set qSelect = Nothing

    If (WebVisionCode = "ITENSVAL") Then 'Estamos na visão do validador de eventos, então...

        If ((CurrentQuery.FieldByName("NIVELERROVALIDACAO").AsInteger < 1) And _
            ((giHandleEventoAnterior <> CurrentQuery.FieldByName("EVENTO").AsInteger) Or (giHandleGrauAnterior <> CurrentQuery.FieldByName("GRAU").AsInteger))) Then

            CancelDescription = "Quando corrigido é necessário informar um nível de erro."
            CanContinue = False

            Exit Sub

        End If

    End If

End Sub
