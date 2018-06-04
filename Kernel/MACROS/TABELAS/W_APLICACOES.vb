'HASH: FF78672FAA242B2A7C6B9AF736E8BE46
Public Sub PUBLICARFARM_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebPublisher") 
  Obj.PublishFarm(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set Obj = Nothing 
End Sub 
 
Public Sub PUBLICARMENUS_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebPublisher") 
  Obj.PublishGroups(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set Obj = Nothing 
End Sub 
 
Public Sub PUBLICARQUESTIONARIOS_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebPublisher") 
  Obj.PublishQuestionnaires(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set Obj = Nothing 
End Sub 
 
Public Sub PUBLICARRELATORIOS_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebPublisher") 
  Obj.PublishReports(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set Obj = Nothing 
End Sub 
 
Public Sub PUBLICARSEGURANCA_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebPublisher") 
  Obj.PublishTables(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set Obj = Nothing 
End Sub 
 
Public Sub PUBLICARTRADUC_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebPublisher") 
  Obj.PublishDictionaries(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set Obj = Nothing 
End Sub 
 
Public Sub PUBLICARTUDO_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebPublisher") 
  Obj.PublishAll(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set Obj = Nothing 
End Sub 
 
Public Sub PUBLICARVISOES_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebPublisher") 
  Obj.PublishVisions(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set Obj = Nothing 
End Sub 
 
Public Sub PUBLICARWORKFLOW_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebPublisher") 
  Obj.PublishWFFieldsVisibility(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set Obj = Nothing 
End Sub 
 
Public Sub RECRIARINDICE_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebPublisher") 
  Obj.RebuildContentIndex(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set Obj = Nothing 
End Sub 
 
Public Sub GERARSERVICOS_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Benner.Tecnologia.Architect.CustomServicesForm") 
  Obj.PublishAll(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set Obj = Nothing 
End Sub 
 
