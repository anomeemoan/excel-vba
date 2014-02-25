'Sample data'
'Department		Region 		Profit
'1000			 5			4564645
'1222			 6			4565645
'5522			 7			8564645

Sub sbCreatePivot

'Declaration
Dim ws  as Worksheet
Dim pc  as PivotCache
Dim pt  as PivotTable

'Create Pivot cache
Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, "Sheet1!R1C1:R10C3")

'Create Pivot Table
Set pt = pc.CreatePivotTable(ws.Range("B3"))

'Setting Fields
With pt
	'Set Row field
	With .pivotFields("Department")
			.orientation = xlRowField
			.Position = 1
	End With
	
	'set Column Field
	With .pivotFields("Region")
			.orientation = xlColumnField
			.Position = 1
	End With
	
	'Set data field
	.AddDataField .PivotFieldS("Profit"), "Sum of Profit", xlSum

End With

	
	

