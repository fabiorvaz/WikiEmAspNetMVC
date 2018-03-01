<%
' Struct to hold title/lastmodified data for generating
' The recent changes list
class NodeInfo
	public title
	public lastModified
	public createdOn
	public usuario
	public criador

	public sub init(aTitle, aLastModified, aCreatedOn, aUsuario, aCreatedBy)
		title = aTitle
		lastModified = aLastModified
		createdOn = aCreatedOn
		usuario = aUsuario
		criador = aCreatedBy
	end sub
end class

' Compare function for sorting a list of nodes by date
function cmpNodeInfo(a, b)
	if a.lastModified < b.lastModified then
		cmpNodeInfo = CMP_LESS
	elseif a.lastModified > b.lastModified then
		cmpNodeInfo = CMP_GREATER
	else
		cmpNodeInfo = 0
	end if
end function

function GroupNodesByDate(nodes)
	' Maximum length for subgroups
	dim length : length = UBound(nodes) + 1

	dim groups
	redim groups(length)
	
	dim nGroups : nGroups = 0

	dim thisGroup
	dim iGroup : iGroup = 0
	
	dim node
	dim lastDate : lastDate = ""
	
	dim currentDate
	
	for each node in nodes
		currentDate = FormatDateTime(node.lastModified, vbLongDate)
		if currentDate <> lastDate then
			if iGroup <> 0 then
				redim preserve thisGroup(iGroup - 1)
				groups(nGroups) = thisGroup
				nGroups = nGroups + 1
			end if
			
			redim thisGroup(length)
			iGroup = 0
			
			lastDate = currentDate
		end if
		
		' And this node to this group
		set thisGroup(iGroup) = node
		iGroup = iGroup + 1
	next
	
	if iGroup <> 0 then
		redim preserve thisGroup(iGroup - 1)
		groups(nGroups) = thisGroup
		nGroups = nGroups + 1
		
		redim thisGroup(length)
		iGroup = 0
	end if

	redim preserve groups(nGroups - 1)
	GroupNodesByDate = groups
end function
%>
