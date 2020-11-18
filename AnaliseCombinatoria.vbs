set fso=createobject("scripting.FileSystemObject")
dim x
x = "./teste.txt"
set y=fso.opentextfile(x,2,true)
for i= 1 to 5
for b= i + 1 to 5
y.write "[" & i & "," & b & "]" & vbCrLf
next
next