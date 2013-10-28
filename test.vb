'**********************************************************************
'Script Name:  Enter Filter Field
'Author:		Alfredo González
'Created Date:  11/18/2011
'Purpose:				Set the description of a filter's field.
'Prerequisites:	    Cost Repository - Any Table Query - table has been opened.
'Change history:
'**********************************************************************

Option Explicit

'**********************************************************************
'Variable Declaration
'**********************************************************************
Dim ClickX
Dim ClickY
Dim ClickYSafe
Dim CurrentRowCount
Dim FieldRow
Dim FilterDescription
Dim FilterName
Dim i
Dim IncludeCheckbox
Dim InitialRowCount
Dim LowerBound
Dim NotCheckbox
Dim Order
Dim ParentFilterRow
Dim RowCount
Dim RowCountWithExpander
Dim RowName
Dim SubfilterName
Dim SubfilterRowCount
Dim SubtotalCheckbox
Dim UpperBound
Dim VisibleRowCount

'**********************************************************************
'Variable Assignment
'**********************************************************************
FilterDescription = Parameter("FilterDescription")
FilterName = Parameter("FilterName")
IncludeCheckbox = Parameter("IncludeCheckbox")
NotCheckbox = Parameter("NotCheckbox")
Order = Parameter("Order")
SubfilterName = Parameter("SubfilterName")
SubtotalCheckbox = Parameter("SubtotalCheckbox")

'**********************************************************************
'Business Process
'**********************************************************************
If PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").Exist(30) Then
	PbWindow("PowerPlant").Maximize
	Reporter.ReportEvent micPass, "Filter Table pass.", "The filter table opened successfully."

	ParentFilterRow = 0
	
	'If there is a subfilter
	If SubfilterName <> "SKIP" Then
		'Get the number of rows as the upper bound.  The lower bound will be 1 and the visible row count at the moment will be a guess at the number in between.
		UpperBound = PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").RowCount
		VisibleRowCount=UpperBound/2
		LowerBound = 1
	
		'Setting initial condition that the scroll bar is to the top.
		PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").SelectCell 1, "description_edit"
		
		'Run until i=UpperBound so that this component scales well with multiple trials and has an end point.
		For i=1 to UpperBound
			'If our lower bound equals the guessed visible row count then we have found the true visible row count.
			If LowerBound = VisibleRowCount Then
				Exit For
			End If
			'Click on the middle number row.
			PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").SelectCell VisibleRowCount,"description_edit"
	
			'If the scroll position moves then move the upper bound and find the new middle number.  After this reset to the initial position.
			If PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").GetROProperty("vscrollposition") > 0 Then
				UpperBound = VisibleRowCount
				VisibleRowCount= Int((LowerBound + VisibleRowCount)/2)
				PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").SelectCell 1, "description_edit"
			'If the scroll position doesn't move then move the lower bound and update our middle number.
			Else
				LowerBound = VisibleRowCount
				VisibleRowCount = Int((UpperBound + VisibleRowCount)/2)
			End If
		Next
		'Start the page with the first item at the top of the window
		PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").SelectCell "#1", "description_edit"
		'Set the starting values for the search coordinates
		ClickX = 5
	
		'Set the RowCount vaiable to the number of rows in the filter table
		InitialRowCount = PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").RowCount
		
		'Loop through the rows to find the specified parent filter's row and...
		For i = 1 To InitialRowCount
			RowName = PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").GetCellData(i, "description")
			If i>VisibleRowCount Then
				PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").SelectCell i, "description_edit"
			End If
			'... if the Field name is found, click in the description field of that row
			If FilterName =  RowName Then
				'Capture row number
				ParentFilterRow = i
				Exit For
			End If
		Next
	
		'The Parent Filter Field should now be on screen, expand the parent field to show the subfields
		'If the field was on a row greater than 24, when clicked it will sit at the bottom of the viewable page
		If i > VisibleRowCount Then
			ClickY = VisibleRowCount * 28
			For ClickYSafe = ClickY-2 To ClickY+2
				PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").Click ClickX, ClickYSafe
				RowCountWithExpander = PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").RowCount
				If RowCountWithExpander > InitialRowCount Then
					Exit For
				End If
			Next
		Else
			'Otherwise, calculate the expander's position and click it.
			ClickY = i*28
			For ClickYSafe = ClickY-2 To ClickY+2
				PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").Click ClickX, ClickYSafe
				RowCountWithExpander = PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").RowCount
				If RowCountWithExpander > InitialRowCount Then
					Exit For
				End If
			Next
		End If

		'Make sure our filter expanded.
		If RowCountWithExpander > InitialRowCount Then
			Reporter.ReportEvent micPass, "Expander pass.", "The filter was successfully expanded to expose its subfields."
		Else
			Reporter.ReportEvent micFail, "Expander fail.", "The filter was unsuccessfully expanded.  The required subfield will not be visible."
		End If

		'Start the search for the subfilter field; since the filter will be beneath the Parent's position, start at parent
		SubfilterRowCount = RowCountWithExpander - InitialRowCount
		For i = ParentFilterRow To SubfilterRowCount+ParentFilterRow
			RowName = PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").GetCellData(i, "description")
	
			'... if the Filed name is found, Set the required information for that row
			If SubfilterName =  RowName Then
				'Capture row number
				FieldRow = i
				'Set the description for the field
				PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").SetCellData i, "description_edit", FilterDescription
				Reporter.ReportEvent micPass, "Subfilter pass.", "Specified subfilter ["& SubfilterName  &"] was found."
				'Set the checkbox column titled "Not"
				If NotCheckbox <> "SKIP" Then
					PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").SetCellData i, "order", NotCheckbox
				End If
				'Set the checkbox column titled "Include"
				If IncludeCheckbox <> "SKIP"  Then
					PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").SetCellData i, "include_in_select", IncludeCheckbox
				End If
				'Set the checkbox column titled "Subtotal"
				If SubtotalCheckbox <> "SKIP" Then
					PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").SetCellData i, "element_id" , SubtotalCheckbox
				End If
				'Set the Order Column value
				If Order <> "SKIP" Then
					PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").SetCellData i, "sort_order", Order
				End If
				Exit For
			'If we are in the last row and can't find the subfilter then report a fail.
			ElseIf i = SubfilterRowCount+ParentFilterRow Then
				Reporter.ReportEvent micFail, "Subfilter fail.", "Specified subfilter ["& SubfilterName  &"] could not be found."
			End If
		Next
		'Ends opening the Parent Filter's subfilter options and setting the parameter information
	
		'Close the subfilter again
		'To reposition the coordinates, start by clicking on the first entry
		PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").SelectCell 1, "description_edit"
		
		'Loop through the rows to find the specified parent filter's row and...
		For i = 1 To RowCountWithExpander
			RowName = PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").GetCellData(i, "description")
			If i>VisibleRowCount Then
				PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").SelectCell i, "description_edit"
			End If
			'... if the Field name is found, click in the description field of that row
			If FilterName =  RowName Then
				'Capture row number
				ParentFilterRow = i
				Exit For
			End If
		Next
	
		'The Parent Filter Field should now be on screen, expand the parent field to show the subfields
		'If the field was on a row greater than 24, when clicked it will sit at the bottom of the viewable page
		If i > VisibleRowCount Then
			ClickY = VisibleRowCount * 28
			For ClickYSafe = ClickY-2 To ClickY+2
				PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").Click ClickX, ClickYSafe
				CurrentRowCount = PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").RowCount
				If CurrentRowCount = InitialRowCount Then
					Exit For
				End If
			Next
		Else
			'Otherwise, calculate the expander's position and click it.
			ClickY = i*28
			For ClickYSafe = ClickY-2 To ClickY+2
				PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").Click ClickX, ClickYSafe
				CurrentRowCount = PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").RowCount
				If CurrentRowCount = InitialRowCount Then
					Exit For
				End If
			Next
		End If

		'Make sure that the row count is set back to its initial setting.
		If CurrentRowCount = InitialRowCount Then
			Reporter.ReportEvent micPass, "Collapse pass.", "The filter was successfully collapsed to hide its subfields."
		Else
			Reporter.ReportEvent micFail, "Collapse fail.", "The filter was unsuccessfully collapsed.  The parent filter's subfields are still visible."
		End If

	Else
		'If there is no Subfilter
		'Reset the RowCount vaiable to the number of rows in the filter table
		RowCount = PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").RowCount
		
		'Loop through the rows to find the specified row (starting on the parent filter's row) and...
		For i = 1 To RowCount
			RowName = PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").GetCellData("#"&i, "description")
		
			'... if the Filed name is found, Set the required information for that row
			If FilterName =  RowName Then
				'Capture row number
				FieldRow = i
				'Set the description for the field
				PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").SetCellData i, "description_edit", FilterDescription
				Reporter.ReportEvent micPass, "Filter pass.", "Specified filter ["& FilterName &"] was found."
				'Set the checkbox column titled "Not"
				If NotCheckbox <> "SKIP" Then
					PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").SetCellData i, "order", NotCheckbox
				End If
				'Set the checkbox column titled "Include"
				If IncludeCheckbox <> "SKIP"  Then
					PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").SetCellData i, "include_in_select", IncludeCheckbox
				End If
				'Set the checkbox column titled "Subtotal"
				If SubtotalCheckbox <> "SKIP" Then
					PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").SetCellData i, "element_id" , SubtotalCheckbox
				End If
				'Set the Order Column value
				If Order <> "SKIP" Then
					PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").SetCellData i, "sort_order", Order
				End If
		
				Exit For
			End If
		Next
	End If
	PbWindow("PowerPlant").PbWindow("Cost Repository - Any Table").PbDataWindow("Filter Table").SelectCell "#1", "description_edit"
Else
	Reporter.ReportEvent micFail, "Filter Table fail.", "The filter table did not open."
End If


'**********************************************************************
'Output Assignment
'**********************************************************************
'**********************************************************************
'End File
'**********************************************************************





