Sub sbPivotChartInNewSheet()

'Declaration
Dim pt as PivotTable, ptr as Range, cht as Chart

'If no pivots exit procedure
	If ActiveSheet.PivotTables.Count = 0 then Exit Sub
	
'Setting pivot table
	Set pt = ActiveSheet.PivotTables(1)
	set ptr = pt.TableRange1

'Add a new chart sheet
Set cht = Charts.Add
	With cht
		.SetSourceData ptr
		.ChartType = xlLine
	end With
End Sub


	


	