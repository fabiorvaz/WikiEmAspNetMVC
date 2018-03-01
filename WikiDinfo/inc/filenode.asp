<%
dim gFolderPath
gFolderPath = Server.MapPath(".") & "\" & NODE_FOLDER

' The node manager
' Can provide a sorted list of recently updated pages
class WikiNodeManager

	public function RecentChanges
		const MAX_NODES = 100

		dim fso
		set fso = Server.CreateObject("Scripting.FileSystemObject")

		dim nodeFolder
		set nodeFolder = fso.GetFolder(gFolderPath)

		dim changelist
		redim changelist(nodeFolder.files.count-1)

		dim file
		dim i
		i = 0
		for each file in nodeFolder.files
			dim change
			set change = new NodeInfo
			change.init URLDecode(file.name), file.dateLastModified, file.dateCreated
			set changelist(i) = change
			i = i + 1
		next

		' Sort the list, newest access times first.
		dim sort
		set sort = new QSort
		set sort.order = ORDER_DESC
		set sort.compare = GetRef("cmpNodeInfo")
		sort.sort changelist

		dim titlecount
		titlecount = min(nodeFolder.files.count, MAX_NODES)

		' CHOMP!
		redim preserve changelist(titlecount -1)

		RecentChanges = changelist
	end function


	public function TitleSearch(byval pattern)
		dim fso
		set fso = Server.CreateObject("Scripting.FileSystemObject")

		dim nodeFolder
		set nodeFolder = fso.GetFolder(gFolderPath)

		dim linklist
		redim linklist(nodeFolder.files.count-1)

		dim i
		i = 0

		dim file
		for each file in nodeFolder.files
			dim nodeName
			nodeName = URLDecode(file.name)

			if instr(1, nodeName, pattern, vbTextCompare) then
				linklist(i) = nodeName
				i = i + 1
			end if
		next

		if 0 < i then
			redim preserve linklist(i-1)
			dim sort : set sort = new QSort
			sort.sort linklist
'			BubbleSort(linklist)
			TitleSearch = linklist
		else
			TitleSearch = Empty
		end if
	end function


	' Return nodes that contain the given text
	public function FullSearch(byval pattern)
		dim fso
		set fso = Server.CreateObject("Scripting.FileSystemObject")

		dim nodeFolder
		set nodeFolder = fso.GetFolder(gFolderPath)

		dim linklist
		redim linklist(nodeFolder.files.count-1)

		dim i
		i = 0

		dim file
		for each file in nodeFolder.files
			dim node
			set Node = new WikiNode
			Node.init URLDecode(file.Name)

			if _
				instr(1, node.Name, pattern, vbTextCompare) _
			or _
				instr(1, node.getText(), pattern, vbTextCompare) _
			then
				linklist(i) = Node.name
				i = i + 1
			end if
		next

		if 0 < i then
			redim preserve linklist(i-1)
			dim sort : set sort = new QSort
			sort.sort linklist
'			BubbleSort(linklist)
			FullSearch = linklist
		else
			FullSearch = Empty
		end if
	end function


	' Return nodes that link to the given node
	public function BackLinks(byval node_name)
		BackLinks = FullSearch( "[" & node_name & "]" )
	end function
end class


' One text node in a wiki
' Uses a folder of files as the datastore

class WikiNode
	private fso

	private nodeName
	private fileName

	private cacheContents
	private cacheLastModified
	private bNew
	
	private sub class_initialize
		set fso = Server.CreateObject("Scripting.FileSystemObject")
	end sub
	
	
	sub init(theNodeName)
		nodeName = theNodeName
		fileName = gFolderPath & "\" & Server.URLEncode(nodeName)		

		dim file
		dim tstream

		if fso.FileExists(fileName) then
			set file = fso.GetFile(fileName)
			set tstream = file.OpenAsTextStream(ForReading)
			
			if tstream.AtEndOfStream then
				cacheContents = ""
			else
				cacheContents = tstream.Readall
			end if
			
			cacheLastModified = file.DateLastModified
			tstream.Close
		else
			cacheContents = ""
			cacheLastModified = "(Novo t&oacute;pico)"
		end if

		if cacheContents = "" then
			bNew = True
		else
			bNew = False
		end if
	end sub


	public property get name
		name = nodeName
	end property

	public property get isNew
		isNew = bNew
	end property

	public property get lastModified
		lastModified = cacheLastModified
	end property

	public property get canEdit
		canEdit = True
	end property


	function getText()
		getText = cacheContents
	end function
	
	
	sub setText(text)
		dim file
		if not fso.FolderExists(gFolderPath) then fso.CreateFolder(gFolderPath)
	
		Application.Lock
		if text = "" then
			fso.DeleteFile(fileName)
		else
			set file = fso.OpenTextFile(fileName, ForWriting, True)
			file.Write text
			file.Close
		end if
		Application.Unlock
	end sub
	
end class
%>
