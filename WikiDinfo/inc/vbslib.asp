<%
' File open mode constants
const ForReading = 1
const ForWriting = 2
const ForAppending = 8


function min(a,b)
	if a < b then
		min = a
	else
		min = b
	end if
end function

function max(a,b)
	if a > b then
		max = a
	else
		max = b
	end if
end function

' An inverse to Server.URLEncode
function URLDecode(str)
	dim re
	set re = new RegExp

	str = Replace(str, "+", " ")
	
	re.Pattern = "%([0-9a-fA-F]{2})"
	re.Global = True
	URLDecode = re.Replace(str, GetRef("URLDecodeHex"))
end function

' Replacement function for the above
function URLDecodeHex(match, hex_digits, pos, source)
	URLDecodeHex = chr("&H" & hex_digits)
end function

function Contains(collection, key)
	Contains = (collection(key) <> "")
end function

function debugwrite(msg)
	response.write "[" & msg & "]"
end function
%>
