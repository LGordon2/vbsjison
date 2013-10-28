'pt Name:								Bronze October 2013
'Created Date: 					10/14/13
'Purpose:                            Calculates the factorial of a given integer
'********************************************************************************************
Option Explicit

'********************************************************
'Processes
'********************************************************

'This function returns the factorial and accepts the given integer
'Function UserName(factorial)

	'if the number given was below 1 it is invalid
	If factorial < 1 Then
		UserName = 1

	'If the number is a one or a two then return that value to stop the recursion
	ElseIf factorial = 2 OR factorial = 1Then
		UserName = factorial

	'Else calculate the factorial recursively decrementing by 1
	Else
		UserName = factorial * UserName(factorial-1)
	End If

'End Function

'********************************************************
'End File
'********************************************************
