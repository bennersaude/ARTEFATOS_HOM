'HASH: 813F396A9B47C75B11FAFB3E70BB79E9
'Macro da tabela CA_TIPOSERVICO

Option Explicit

'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

    If (CurrentQuery.FieldByName("TIPOSERVICOMENSAGERIA").AsString = "S") Then

        Dim qSelect As Object
        Set qSelect = NewQuery

        qSelect.Clear
        qSelect.Add("SELECT DESCRICAO                      ")
        qSelect.Add("  FROM CA_TIPOSERVICO                 ")
        qSelect.Add(" WHERE TIPOSERVICOMENSAGERIA = :SIM   ")
        qSelect.Add("   AND HANDLE               <> :HANDLE")

        qSelect.ParamByName("SIM"   ).AsString  = "S"
        qSelect.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

        qSelect.Active = True

        Dim vbJaExisteUmTipoServicoMensageria As Boolean
        Dim vsDescricaoTipoServicoExistente   As String

        vbJaExisteUmTipoServicoMensageria = Not qSelect.EOF
        vsDescricaoTipoServicoExistente = qSelect.FieldByName("DESCRICAO").AsString

        qSelect.Active = True
        Set qSelect = Nothing

        If (vbJaExisteUmTipoServicoMensageria) Then
            bsShowMessage("Já existe o tipo de serviço '" + vsDescricaoTipoServicoExistente + "' para mensageria.", "I")
            CanContinue = False
            Exit Sub
        End If

    End If

End Sub
