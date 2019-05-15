# All basic Syntax

## variable declaration
use dim for declaring simple variable. 
```vba
Dim x As Integer
Dim y as String
```
Array Declaration. 
These are example of integer and string array. 
```vba
Dim x() As Integer
Dim y() as String
```

## boolean operation  
use and /or for doing the boolean operation. 
```vba
If LWebsite = "TechOnTheNet.com" And LPages <= 10 Then
   LBandwidth = "Low"
ElseIf LWebsite = “google.com“ or LWebsite = “facebook.com“ 
   LBandwidth = "High"
End If

```
## String manipulation  
There are lot of ways to manipulate string here are some example commonly used. 
```vba
Left(“Hello”,2)  'He
Right(“hello”,2) ' lo
Mid(“hello”,2,3) ' ell
Split(“hello world”) ' {“hello” , “world”}
Split(“Hi hello-world”, “-”) ' {“Hi hello”,”world”}
Split(“A;B;C;D”, ”;” , 3) ' {“A”,”B”,”CD”}
dim x() as string
x = Split(“A;B;C;D”, ”;” , 3) 
' x(0), x(1), x(2) to access splited data. 

```

## If Else statement  
This is a simple example of if else statement in vba. we can have only if  or if else or if elseif else in the code. 

```vba
IF
If a=1 Then 
debug.print “a=1”
ElseIf a=2 Then
debug.print “a=2”
Else 
debug.print “a is something else”
End If
```



## for Loop  
There is a loop counter  
Few example shown here.   
1) This is for 1 increment. 
```vba
Dim LCounter As Integer
For LCounter = 1 To 5
   MsgBox (LCounter)
Next LCounter
```
2) This is for a custom increment here it is for 5 increment. 
```vba
For LCounter = 50 To 30 Step -5
     MsgBox LCounter
Next LCounter
```
3) This is for iterating over an array or group. 
```vba
For Each p In arrofpath
	Debug.print p
Next

```
4) Do while type loop  
```vba
	Do
        ' do something and wait for x = ""
    Loop Until x = “”
```

## File Handling

example of creating and writing data to a external file. 
```vba
Dim MyIndex, FileNumber
For MyIndex = 1 To 5 
FileNumber = FreeFile 
    Open "TEST" & MyIndex For Output As #FileNumber
    Write #FileNumber, "This is a sample." 
    Close #FileNumber    
Next MyIndex
```
