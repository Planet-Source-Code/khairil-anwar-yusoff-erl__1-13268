
not so many vb programmer realize the existance of Erl (error line) and some of them who used it alway get the value of 0. 

here is an example, (makes me remember my old gwbasic). you can put the line number at the beginning of visual basic code, but not the function/sub header.

eg.

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Const strModuleName = "Form1.frm" '//module or form or class

Private Sub myButtonClick()
On Error Goto errClick
12345 '//NOTICE, THIS IS THE LINE NUMBER, 
      '// YOU CAN PUT WHATEVER NUMBER FORM  UP TO 65535(max)

Dim intX as Integer

	'//GENERATING RUNTIME ERROR
	intX = 12 \ 0 'Division by Zero

Exit Sub
errClick:
	'//Show what line number does Erl recently passed before error 
	'occurs	
	MsgBox Err.Number & " - " & Err.Decription, , _
       strModuleName & " - " &Erl 
End Sub

This will pop up a message box "11 - Division by Zero" with title Form1.frm - 12345. 

Numbering every code line is tedious thing to do, some suggestion, just put UNIQUE number ID(plus some constant of your module/form/class Name) at the beginning of you critical function (function which error is likely to happen). so when error comes, you will know which function it came from