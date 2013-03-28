<!--#include file='../utility.asp'-->
<!--#include file='xmlFunctions.asp'-->
<%
'take the variable that was set on the index page
toEncode = session("manip")
'store what people are encoding, because im nosy
save = toEncode
writeFromArray()
'initialize an array to hold the cut up strings
dim subStrings()
'replace the line breaks with " {%^!@} " it'll add some cool lookin chars
toEncode = replace(toEncode,vbCrLf," {%^!@} ")
'pick a skip pattern, between 3 and 9. this part is extremely crucial
skip=random(9,3)
'pick a shifter
shift=random(9,2)
'make sure its long enough
	if(len(toEncode)<20) then
		for i=1 to 25-len(toEncode)
			toEncode = toEncode & randomCase(alphabet(random(25,0)))
			toLengthen = i
		next
	end if
'-------------------------
'add the shifter to the ascii of the letter
	for i=1 to len(toEncode)
		char = mid(toEncode,i,1)
		asci = asc(char) + shift
		char = chr(asci)
		newString = newString & char
	next
	toEncode = newString
'-------------------------
'make the string length divisble by the skip
if(not factor(len(toEncode),skip))then
	for i=1 to skip-1
		'add random letters til it gets to the right length
		toEncode = toEncode + randomCase(alphabet(random(25,0)))
		'store the number of added extra so in decode, we can take 'em off.
		extra = i
		if(factor(len(toEncode),skip))then
			exit for
		end if
	next
end if
'take every skipth letter out
'ie: it'll go through and take every fourth letter and put them at the beginning
'then it'll take every fourth letter after that and place it after the first set
for m=1 to skip
	for i=0 to len(toEncode)/skip
		'change variable from "toEncode" to "code"
		code = code + mid(toEncode,(i*skip)+m,1)
	next
next

'split it up into skip number of strings, these will be the sets 
'that were taken out earlier (ie: every 4th)
	for i=0 to (len(code)/skip) - 1
		redim preserve subStrings(i)
		subStrings(i) = mid(code,(i*skip)+1,skip)
	next
	code = ""
'---------------------------------------------------------------
'make every skipth letter random
	for i=0 to ubound(subStrings)
		code = code & subStrings(i) & alphabet(random(25,0))
	next
'----------------------------

'add config letters:
	'first = number of config letters (including this one)
	config = config & randomCase(alphabet(5))
	'second = number of extra spaces
	config = config & randomCase(alphabet(extra))
	'third = skip
	config = config & randomCase(alphabet(skip))
	'fourth = number of extra characters added to lengthen it
	config = config & randomCase(alphabet(toLengthen))
	'fifth = shifter
	config = config & randomCase(alphabet(shift))
	code = config & code
'-------------------
'reverse it
	code = reverse(code)
'----------
response.Write "Encoded:<br /><textarea rows=20 cols=70>" & code & "</textarea><br />Copy this and place it back in the box on the main page to decode."
session("manip") = ""
%>