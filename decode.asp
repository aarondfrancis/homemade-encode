<!--#include file='../utility.asp'-->
<!--#include file='xmlFunctions.asp'-->
<%
code = session("manip")
'reverse it
code = reverse(code)
'----------
length = len(code)
'take off all config letters
	'get the number of config letters based on the first letter
	configAmount = alphaNumber(mid(code,1,1))
	'get the letters out
	configLetters = mid(code,1,configAmount)
	'change variable for shortness' sake
	cL = configLetters
	'take off the first config letter
	cL = mid(cL,2,len(cL)-1)
	'take all the configs out of the code
	code = mid(code,configAmount+1,len(code)-configAmount)
'---------------------------

'read second config letter :: extra chars
	extra = alphaNumber(mid(cL,1,1))
	'take it off
	cL = mid(cL,2,len(cL)-1)
'----------------------------------

'read third config letter :: skip
	skip = alphaNumber(mid(cL,1,1))
	'take it off
	cL = mid(cL,2,len(cL)-1)
'----------------------------------

'read fourth config letter :: toLengthen (number added to make it long enough)
	toLengthen = alphaNumber(mid(cL,1,1))
	'take it off
	cL = mid(cL,2,len(cL)-1)
'----------------------------------

'if theres a fifth, take it off
	if(len(cl) > 0)then
		shift = true
		shifter = alphaNumber(mid(cL,1,1))
		cL = mid(cL,2,len(cL)-1)
	end if
'----------------------------------

'take out the random letters
	for i=0 to len(code)/skip
		withOutRandom = withOutRandom & mid(code,(i*skip)+1+i,skip)
	next
	code = withOutRandom
'---------------------------

length = len(code)
'divide the string into skip number of strings
	dim rebuild()
	redim rebuild(0)
	for i=0 to skip-1
		rebuild(i) = mid(code,(length/skip)*i+1,length/skip)
		redim preserve rebuild(i+1)
	next
'---------------------------------------------

'write every first letter from the strings,
'then every second.. and so on.
	for m=1 to len(rebuild(0))
		for i=0 to skip-1
			decoded = decoded + mid(rebuild(i),m,1)
		next
	next
'---------------------------------------------

'remove extraenuous characters
	decoded = mid(decoded,1,len(decoded)-extra)
'-----------------------------



'take the shifter off all the ascii
	if(shift)then
	for i=1 to len(decoded)
		char = mid(decoded,i,1)
		asci = asc(char) - shifter
		char = chr(asci)
		newString = newString & char
	next
	decoded = newString
	end if
'----------------------------------

'take off the chars that were added to lengthen
	decoded = mid(decoded,1,len(decoded)-toLengthen)
'----------------------------------------------

decoded = replace(decoded," {%^!@} ",vbCrLf)

'store it
save = decoded
writeFromArray()

response.Write "Decoded:<br /><textarea rows=20 cols=70>" & decoded & "</textarea>"
session("manip") = ""
%>