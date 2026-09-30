
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Declare a variable

' There are 2 types of variables. 
' Primitive type (boolean, integer, real, string) and Object type (object, collection)

Dim nameStr
nameStr = "Fred"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim xInt: xInt = 6 ' You can put them in one line.

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim nameStrArr(3) ' Declare an array of variables
nameStrArr(0) = "Fred"
nameStrArr(1) = "Beth"
nameStrArr(2) = "Max"
nameStrArr(3) = "Janet" 

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Control Statements

If nameStr = "Fred" Then
 ' You can add zero or more statements here
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If nameStr = "Fred" Then
 ' You can add zero or more statements here
Else
 ' You can add zero or more statements here
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

For counterInt = 0 To 10 Step 2
 ' You can add zero or more statements here
Next

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

For counterInt = 0 To 10 Step 2
 ' Before: You can add zero or more statements here
If counterInt > 4 Then Exit For
 ' After: You can add zero or more statements here
Next 

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

For Each traceObj In traceColl
 ' You can add zero or more statements here
Next

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

While tempInt > 0
 ' You can add zero or more statements here
 tempInt = tempInt - 1
Wend 

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Select Case nameStr
 Case "Fred"
 ' You can add zero or more statements here
 Case "Bert"
 ' You can add zero or more statements here
 Case Else
 ' You can add zero or more statements here
End Select

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Function

Function AddTwoInts(num1Int, num2Int)
  Dim tempInt
  tempInt = num1Int + num2Int
  MsgBox( "The result is " & tempInt)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function AddTwoInts(num1Int, num2Int)
  Dim tempInt
  tempInt = num1Int + num2Int
  AddTwoInts = tempInt
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

totalInt = AddTwoInts(6,10) ' (a) Return value is used
Call AddTwoInts(7,11) ' (b) Return value is discarded


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Message Box
  
MsgBox("Hello World")
MsgBox("There are " & countInt & " vias.")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

nameStr = InputBox("What is your name?")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Call pcbAppObj.Gui.StatusBarText("Finding bottom side components...", epcbStatusField1)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  
