Sub Main
	Call JoinDatabase()	'Aufbau SAT Lieferungen
	Call AppendField()	'Berechnung Lieferdauer
	Call JoinDatabase1()	'Verbindung Lieferungen & Stammdaten

	Call ModifyField()		'Anpassung der Feldnamen Lieferungen 1
	Call ModifyField1()	'Anpassung der Feldnamen Lieferungen 2
	Call ModifyField2()	'Anpassung der Feldnamen Stammdaten 1
	Call ModifyField3()	'Anpassung der Feldnamen Stammdaten 2
	Call AppendField1()	'Hinzufügen Testfeld für Incoterms
	Call DirectExtraction()	'Extraktion der Abweichungen
	Client.RefreshFileExplorer
End Sub

' Verbindung Lieferungen Kopf & Position
Function JoinDatabase
	Set db = Client.OpenDatabase("LIPS.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "LIKP.IMD"
	task.AddPFieldToInc "VBELN"
	task.AddPFieldToInc "POSNR"
	task.AddPFieldToInc "PSTYV"
	task.AddPFieldToInc "ERNAM"
	task.AddPFieldToInc "ERDAT"
	task.AddPFieldToInc "MATNR"
	task.AddPFieldToInc "MATWA"
	task.AddPFieldToInc "MATKL"
	task.AddPFieldToInc "WERKS"
	task.AddPFieldToInc "LGORT"
	task.AddPFieldToInc "CHARG"
	task.AddPFieldToInc "LICHN"
	task.AddPFieldToInc "PRODH"
	task.AddPFieldToInc "LFIMG"
	task.AddPFieldToInc "MEINS"
	task.AddPFieldToInc "VRKME"
	task.AddPFieldToInc "UMVKZ"
	task.AddPFieldToInc "UMVKN"
	task.AddPFieldToInc "NTGEW"
	task.AddPFieldToInc "BRGEW"
	task.AddPFieldToInc "GEWEI"
	task.AddPFieldToInc "UEBTK"
	task.AddPFieldToInc "UEBTO"
	task.AddPFieldToInc "UNTTO"
	task.AddPFieldToInc "CHSPL"
	task.AddPFieldToInc "FAKSP"
	task.AddPFieldToInc "LGMNG"
	task.AddPFieldToInc "ARKTX"
	task.AddPFieldToInc "VGBEL"
	task.AddPFieldToInc "VGPOS"
	task.AddPFieldToInc "BWART"
	task.AddPFieldToInc "BWLVS"
	task.AddPFieldToInc "VGREF"
	task.AddPFieldToInc "VKGRP"
	task.AddPFieldToInc "VTWEG"
	task.AddPFieldToInc "SPART"
	task.AddSFieldToInc "BZIRK"
	task.AddSFieldToInc "VSTEL"
	task.AddSFieldToInc "VKORG"
	task.AddSFieldToInc "LFART"
	task.AddSFieldToInc "LDDAT"
	task.AddSFieldToInc "TDDAT"
	task.AddSFieldToInc "LFDAT"
	task.AddSFieldToInc "ABLAD"
	task.AddSFieldToInc "INCO1"
	task.AddSFieldToInc "INCO2"
	task.AddSFieldToInc "EXPKZ"
	task.AddSFieldToInc "ROUTE"
	task.AddSFieldToInc "FAKSK"
	task.AddSFieldToInc "LIFSK"
	task.AddSFieldToInc "VBTYP"
	task.AddSFieldToInc "VSBED"
	task.AddSFieldToInc "KUNNR"
	task.AddSFieldToInc "KUNAG"
	task.AddSFieldToInc "KDGRP"
	task.AddSFieldToInc "BTGEW"
	task.AddSFieldToInc "NTGEW"
	task.AddSFieldToInc "GEWEI"
	task.AddMatchKey "MANDT", "MANDT", "A"
	task.AddMatchKey "VBELN", "VBELN", "A"
	task.CreateVirtualDatabase = False
	dbName = "SAT Lieferungen.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Berechnung Lieferdauer
Function AppendField
	Set db = Client.OpenDatabase("SAT Lieferungen.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "LIEFERDAUER"
	field.Description = ""
	field.Type = WI_VIRT_NUM
	field.Equation = "@Age(LDDAT;LFDAT)"
	field.Decimals = 0
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Verbindung SAT LIeferungen mit Stammdaten (KNVV)
Function JoinDatabase1
	Set db = Client.OpenDatabase("SAT Lieferungen.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "KNVV.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "INCO1"
	task.AddSFieldToInc "INCO2"
	task.AddMatchKey "KUNNR", "KUNNR", "A"
	task.CreateVirtualDatabase = False
	dbName = "Lieferungen mit Stammdaten.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Anpassung der Incoterms Felder
Function ModifyField
	Set db = Client.OpenDatabase("Lieferungen mit Stammdaten.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "INCOTERMS_1_LIEFERUNG"
	field.Description = "Incoterms Teil 1"
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 3
	task.ReplaceField "INCO1", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Anpassung der Incoterms Felder
Function ModifyField1
	Set db = Client.OpenDatabase("Lieferungen mit Stammdaten.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "INCOTERMS_2_LIEFERUNG"
	field.Description = "Incoterms Teil 2"
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 28
	task.ReplaceField "INCO2", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Anpassung der Incoterms Felder
Function ModifyField2
	Set db = Client.OpenDatabase("Lieferungen mit Stammdaten.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "INCOTERMS_1_STAMMDATEN"
	field.Description = "Incoterms Teil 1"
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 3
	task.ReplaceField "INCO11", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Anpassung der Incoterms Felder
Function ModifyField3
	Set db = Client.OpenDatabase("Lieferungen mit Stammdaten.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "INCOTERMS_2_STAMMDATEN"
	field.Description = "Incoterms Teil 2"
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 28
	task.ReplaceField "INCO21", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

''Hinzufügen Testfeld für Incoterms
Function AppendField1
	Set db = Client.OpenDatabase("Lieferungen mit Stammdaten.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TEST_INCO"
	field.Description = ""
	field.Type = WI_VIRT_NUM
	field.Equation = "@if(INCOTERMS_1_LIEFERUNG <> INCOTERMS_1_STAMMDATEN .AND. INCOTERMS_1_LIEFERUNG <> """";1;0)"
	field.Decimals = 0
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Extraktion der Abweichungen
Function DirectExtraction
	Set db = Client.OpenDatabase("Lieferungen mit Stammdaten.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Abweichende Incoterms.IMD"
	task.AddExtraction dbName, "", "TEST_INCO = 1"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function