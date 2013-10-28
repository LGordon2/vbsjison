'me: October Coding Challenge - Bronze
'Author: Jason Trogdon
'Created Date:	10/14/2013
'Purpose: To calculate the factorial of a integer
'Prerequisites:		
'Change history:	
'**********************************************************************
Option Explicit
'**********************************************************************
'Variable Declaration
'**********************************************************************

'**********************************************************************
'Variable Assignment
'**********************************************************************

'**********************************************************************
'VBScript Section � Calculations, Formatting, Internal Procedures
'**********************************************************************
Function JasonTrogdon(value1)
   Dim i
   Dim x

	'initialize x
	x = 1

	'Find the length of the value and enter loop
	For i = 1 to value1
		'Update x with new values during loop
		  x = x * i
	Next

	'Print out the factorial's result
	'MsgBox(x)
End Function

'**********************************************************************
'Business Process
'**********************************************************************
Call JasonTrogdon(5)

'**********************************************************************
'End File
'**********************************************************************
