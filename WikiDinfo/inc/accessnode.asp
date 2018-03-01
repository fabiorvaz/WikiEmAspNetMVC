<%
' Access 2000 / MS Jet Engine 4 backend for this Wiki

dim gConnectionString
gConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & GetDBPath()

' 2x Single quotes
function SQLescape(byval sqlstr)
	SQLescape = Replace(sqlstr, "'", "''")
end function

' Fix characters that cause a problem with LIKE
' Assumes sqlstr is literal text
function SQLescapeWild(byval sqlstr)
	dim value
	value = sqlstr

	value = Replace(value, "[", "[[]")
	value = Replace(value, "%", "%%")
	value = Replace(value, "'", "''")
	SQLescapeWild = value
end function


function GetDBPath()
	GetDBPath = Server.MapPath(".") & "\" & NODE_FOLDER & ".mdb"
end function

sub TryCreateDatabase()
	dim fso
	set fso = Server.CreateObject("Scripting.FileSystemObject")

	' If the database already exists then bail; nothing to do
	if fso.FileExists(GetDBPath()) then exit sub

	' Create the MDB file
	dim dbcat
	set dbcat = Server.CreateObject("ADOX.Catalog")

	dbcat.Create gConnectionString

	dim conn
	set conn = Server.CreateObject("ADODB.Connection")
	conn.Open gConnectionString

	' Wish there was a better way to format long text strings in VBS,
	' akin to triple quotes in Python or here documents in Perl.
	dim strSQL
	strSQL = "create table wiki (" & _
		"title varchar(255) not null, " & _
		"contents text, " & _
		"last_modified datetime, " & _
		"created_on datetime, " & _
		"usuario varchar(20), " & _
		"created_by varchar(20), " & _
		"PRIMARY KEY(title))"

	' Create the (only 1) tables
	conn.Execute strSQL

	' Index the last modified date for recent changes
	strSQL = "create index last_modified_index on wiki (last_modified)"
	conn.Execute strSQL
	
	
	strSQL = "create table Usuarios (" & _
		"nome varchar(20) not null, " & _
		"senha varchar(20) not null, " & _
		"PRIMARY KEY(nome))"
	conn.Execute strSQL
	
	strSQL = "create index nome_index on Usuarios (nome)"
	conn.Execute strSQL

	' Allow zero-length contents
	' NOTE: saving a node with zero-length contents currently deletes it
	' from the table

	dbcat.ActiveConnection = conn
	dbcat.tables("wiki").columns("contents").properties("Jet OLEDB:Allow Zero Length") = True
	
	conn.close()
end sub


function OpenDatabase()
	TryCreateDatabase()

	dim conn
	set conn = Server.CreateObject("ADODB.Connection")
	conn.Open gConnectionString
	set OpenDatabase = conn
end function



' The node manager
' Provides various kinds of node searches.
' There's no class state, so these COULD be separate functions.
' The class is only used to try to group them, since VBScript has no modules
class WikiNodeManager

	' Return the MAX_NODES most recently edited nodes
	public function RecentChanges
		const MAX_NODES = 100

		dim conn
		set conn = OpenDatabase()

		dim strSQL
		strSQL = "select top " & MAX_NODES & " * from wiki order by last_modified desc"

		dim rs
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.ActiveConnection = conn
		'Const adOpenStatic = 3
		rs.CursorType = 3

		rs.open strSQL

		dim nodeCount
		nodeCount = min(MAX_NODES,rs.RecordCount)

		dim changelist
		redim changelist(nodeCount-1)
		
		dim i
		i = 0

		while (not rs.EOF) and (i < nodeCount)
			dim change
			set change = new NodeInfo
			change.init rs("title"), rs("last_modified"), rs("created_on"), rs("usuario"), rs("created_by")

			set changelist(i) = change

			rs.MoveNext()
			i = i + 1
		wend

		RecentChanges = changelist
	end function


	' Return nodes whose titles contain 'str'
	public function TitleSearch(byval str)
		dim conn
		set conn = OpenDatabase()

		dim strSQL
		strSQL = "select title from wiki where title like '%" & SQLescape(str) & "%' order by title"

		dim rs
		set rs = conn.Execute(strSQL)

		if rs.eof then
			TitleSearch = Empty
		else
			TitleSearch = rs.GetRows(-1, 0, "title")
		end if
	end function


	' Return nodes whose bodies contain 'str'
	public function FullSearch(byval str)
		dim conn
		set conn = OpenDatabase()

		dim pattern
		pattern = SQLescapeWild(str)

		dim strSQL
		strSQL = "select title from wiki where title like '%" & pattern & "%' or contents like '%" & pattern & "%'  order by title"

		dim rs
		set rs = conn.Execute(strSQL)

		if rs.eof then
			FullSearch = Empty
		else
			FullSearch = rs.GetRows(-1, 0, "title")
		end if
	end function


	' Return nodes that link to the given node
	public function BackLinks(byval node_name)
		BackLinks = FullSearch("[" & node_name & "]")
	end function
end class


' One text node in a wiki
class WikiNode
	private conn

	private nodeName

	private cacheContents
	private cacheIsNew
	private cacheLastModified
	private cacheCreatedOn
	private cacheUsuario
	private cacheCriador
	

	private sub class_initialize
		set conn = Server.CreateObject("ADODB.Connection")
	end sub

	private sub class_terminate
		set conn = Nothing
	end sub
	

	sub init(byval theNodeName)
		dim nodeRow

		set conn = OpenDatabase()

		set nodeRow = conn.execute("select * from wiki where title='" & SQLescape(theNodeName) & "'")

		if nodeRow.EOF then
			cacheContents = ""
			cacheLastModified = "(Novo nodo)"
			cacheCreatedOn = Now
			cacheUsuario = objLogin.NomeUsuario
			cacheCriador = objLogin.NomeUsuario
			nodeName = theNodeName
		else
			cacheContents = nodeRow("contents")
			cacheLastModified = cstr(nodeRow("last_modified"))
			cacheCreatedOn = nodeRow("created_on").Value
			cacheUsuario = cstr(nodeRow("usuario"))
			cacheCriador = cstr(nodeRow("created_by"))
			nodeName = nodeRow("title")
		end if

		if cacheContents = "" then
			cacheIsNew = True
		else
			cacheIsNew = False
		end if
	end sub


	public property get name
		name = nodeName
	end property

	public property get isNew
		isNew = cacheIsNew
	end property

	public property get lastModified
		lastModified = cacheLastModified
	end property
	
	public property get createdOn
		createdOn = cacheCreatedOn
	end property

	public property get canEdit
		canEdit = True
	end property
	
	public property get Usuario
		Usuario = cacheUsuario
	end property

	public property get Criador
		Criador = cacheCriador
	end property

	function getText()
		getText = cacheContents
	end function

	
	sub setText(byval text)
		set conn = OpenDatabase()

		if text <> "" then
			dim nodeRow
			set nodeRow = conn.execute("select title from wiki where title='" & SQLescape(nodeName) & "'")

			if nodeRow.EOF then
				conn.Execute "insert into wiki values ('" & SQLescape(nodeName) & "', '" & SQLescape(text) & "', now(), now(), '" & SQLescape(objLogin.NomeUsuario) & "', '" & SQLescape(objLogin.NomeUsuario) & "')"
			else
				conn.Execute "update wiki set " & _ 
				"contents='" & SQLescape(text) & "', " & _
				"usuario='" & SQLescape(objLogin.NomeUsuario) & "', " & _
				"last_modified = now() " & _ 
				"where title='" & SQLescape(nodeName) & "'"
			end if
		else ' if text = ""
				conn.Execute "delete from wiki where title='" & SQLescape(nodeName) & "'"
		end if

	end sub
	
end class
%>
