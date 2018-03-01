<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="inc/config.asp" -->
<!--#include file="inc/vbslib.asp" -->
<!--#include file="inc/qsort.asp" -->
<!--#include file="inc/NodeInfo.asp" -->
<!--#include file="inc/login.asp" -->
<%
function RenderDateSimples(data)
	RenderDateSimples = FormatDateTime(data, vbShortDate)
end function

function RenderDate(node)
	dim d
	if IsObject(node) then
		d = node.lastModified
	else
		d = node
	end if
	if IsDate(d) then
		RenderDate = FormatDateTime(d, vbLongDate)
	else
		RenderDate = cstr(d)
	end if
end function

' This class is used to show system pages, like search results
class SystemNode
	public name

	public sub init(aName)
		name = aName
	end sub

	public property get canEdit
		canEdit = False
	end property

	public property get isNew
		isNew = False
	end property
	
	public property get Usuario
		usuario = ""
	end property

	public property get Criador
		criador = ""
	end property
end class


' * A wiki system
' Manages wiki nodes, page display, etc.
class Wiki
	private editing
	
	Public Property Get LastVisited()
		LastVisited = Session("LAST_VISITED_WIKI")
	End Property
	
	Public Property Let LastVisited(valor)
		Session("LAST_VISITED_WIKI") = valor
	End Property

	sub run
		dim nodeName
		nodeName = Request("page")
		if nodeName = "" then nodeName = DEFAULT_NODE

		dim node
		set node = GetNode(nodeName)

		if Contains(Request, "button-search") then
			dim searchtype : searchtype = Request("search-type")
			if searchtype = "title" then
				TitleSearch()
			else
				FullSearch()
			end if
			
			Response.End
		end if
		
		If Contains(Request, "button-login") Then
			dim usuario, senha
			usuario = Request("usuario")
			senha = Request("senha")
			If objLogin.Login(usuario, senha) Then
				If Not Request.Form("action") > "" Then
					UserWelcome("")
					Response.End					
				End If
			Else
				Login("Não foi possível efetuar o login")
				Response.End
			End If
		End If
		
		If Contains(Request, "button-alterarsenha") Then
			dim senhaatual, novasenha
			senhaatual = Request("senhaatual")
			novasenha = Request("novasenha")
			If objLogin.AlterarSenha(senhaatual, novasenha) Then
				UserWelcome("Senha alterada")
			Else
				UserWelcome("Falha na alteração da senha")
			End If
			Response.End
		End If
	
		dim op, action
		op = Request("op")
		action = Request("action")
		editing = (op = "edit")
		
		If action > "" Then
			SaveNode(Node)
			exit sub
		End If

		If objLogin.IsLogado() Then
			if op = "userwelcome" then
				UserWelcome("")
				exit sub
			elseif op = "logoff" then
				objLogin.Logoff()
				Login("")
				exit sub
            elseif op = "users" then
                Users("")
                exit sub
            elseif op = "newuser" then
                NovoUsuario("")
                exit sub
			end if
		End If

		if op = "recent" then
			RecentChanges()
			exit sub
		elseif op = "backlinks" then
			BackLinks(Node)
			exit sub
		elseif op = "login" then
			Login("")
			exit sub
		else
			DisplayNode(Node)
			exit sub
		end if
	end sub


	function GetNode(nodeName)
		dim Node
		set Node = new WikiNode
		Node.init nodeName

		set GetNode = node
	end function


	sub pageHeader(node)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title><%= Server.HTMLEncode(node.name) %>-<%= Server.HTMLEncode(WIKI_TITLE) %></title>
<link rel="shortcut icon" href="imagens/documento_wiki.bmp" type="image/gif" />
<link rel="stylesheet" href="css/noodle.css" type="text/css" />
<link rel="stylesheet" href="css/user.css" type="text/css" />
<script src="script/noodle.js" type="text/javascript"></script>
</head>
<body onunload="return _onunload()" onload="return _onload()">
<div id="wraper">
   <div id="corpo">
      <% pageSidebar(node) %>
      <% pageConteudoOpen(node) %>
      <%
	end sub

	
	sub pageSidebar(node)
%>
      <div id="panelsidebar">
         <div class="menu-navegacao">
            <h1 class="menuprocura">Procura</h1>
            <label for="searchtext">Procurar:</label>
            <form action="<%= ASP_PAGE %>" method="get">
               <input type="text" name="q" id="searchtext" title="Search Box" />
               <br />
               <input type="radio" name="search-type" value="title" />
               T&iacute;tulo
               <input type="radio" name="search-type" value="full" checked="checked" />
               Conte&uacute;do<br />
               <input type="submit" value="Procurar" name="button-search" />
               <br />
            </form>
         </div>
<% If Request("q") > "" Then %>		 
<script type="text/javascript">
<!--
function procurar_setfocus()
{
	try
	{
		var a = document.getElementById("searchtext");
		a.focus();
	} catch(e) {}
}
procurar_setfocus();
-->
</script>		 
<% End If %>
         <div class="menu-navegacao">
            <h1 class="menunavegacao">Navega&ccedil;&atilde;o</h1>
            <ul class="lista">
               <li><a class="homelink" href="<%= ASP_PAGE %>" title="T&oacute;pico inicial">In&iacute;cio</a></li>
               <li><a href="<%= ASP_PAGE %>?op=recent" title="&Uacute;ltimos t&oacute;picos modificados">Mudan&ccedil;as recentes</a></li>
               <% If objLogin.IsLogado() Then %>
               <li><a href="<%= ASP_PAGE %>?op=userwelcome" title="P&aacute;gina principal do usu&aacute;rio">Informa&ccedil;&otilde;s da conta</a></li>
               <li><a href="<%= ASP_PAGE %>?op=users" title="Cadastro de editores">Editores</a></li>
               <li><a href="<%= ASP_PAGE %>?op=newuser" title="Criar novo login de editor">Novo editor</a></li>
               <li><a href="<%= ASP_PAGE %>?op=logoff" title="Encerrar a sess&atilde;o de <%=Server.HTMLEncode(objLogin.NomeUsuario)%>">Logoff</a></li>
               <% Else %>
               <li><a href="<%= ASP_PAGE %>?op=login" title="Utilize o login para poder alterar t&oacute;picos">Login</a></li>
               <% End If %>
            </ul>
         </div>
      </div>
      <%
	end sub

	sub pageConteudoOpen(node)
%>
      <div id="panelconteudo">
         <div id="conteudo">
            <div id="cabecalhoconteudo">
               <h1 class="titulonodo"><%= Server.HTMLEncode(node.name) %></h1>
               <div>
                  <% if node.canEdit then %>
                  <br />
                  <% if (not editing) and (not node.isNew) then %>
                  <span class="itemmenunodo"><a href="<%= ASP_PAGE %>?page=<%= Server.URLEncode(node.name) %>&op=edit" title="Aterar o conte&uacute;do do t&oacute;pico">Editar este t&oacute;pico</a></span>
                  <% end if %>
                  <span class="itemmenunodo"><a href="<%= ASP_PAGE %>?page=<%= Server.URLEncode(node.name) %>&op=backlinks" title="Procurar links para este t&oacute;pico">Encontrar Refer&ecirc;ncias</a></span>
                  <%
	if not node.isNew then
%>
                  <span>&laquo; Criado por <%=node.Criador%> em <%=RenderDateSimples(node.CreatedOn)%>. Modificado por <%=node.Usuario%> em <%= RenderDate(node) %>.</span>
                  <% end if
   end if %>
               </div>
            </div>
            <div id="corpoconteudo">
               <%
	end sub
	
	sub pageFooter(node)
		pageConteudoClose(node)
		pageRodape(node)	
		pageFinish(node)
	end sub
	
	sub pageConteudoClose(node)
%>
               &nbsp; </div>
         </div>
      </div>
   </div>
   <%
	end sub
	
	sub pageRodape(node)
%>
   <div id="footer">
      <%
		dim conn
		set conn = OpenDatabase()

		dim strSQL
		dim rs

		strSQL = "select count(*) As Contagem from wiki"
		set rs = conn.Execute(strSQL)
		
		response.Write("<span>Existem " & rs.Fields("Contagem") & " t&oacute;pico(s) no banco de dados.</span>")

		strSQL = "select count(*) As Contagem from Usuarios"
		set rs = conn.Execute(strSQL)
		
		response.Write("<span>&nbsp;" & rs.Fields("Contagem") & " usu&aacute;rio(s) cadastrado(s).</span>")
%>
   </div>
   <%
	end sub
	
	sub pageFinish(node)
%>
</div>
</body>
</html>
<%
	end sub

	sub DisplayNode(node)
		If (editing Or node.isNew) And Not objLogin.IsLogado() Then
			Login("")
			Response.End()
		End If
		
		pageHeader(node)
		
		if editing or (node.isNew) then
			dim prompt
			if node.isNew then
				prompt = "Entre com o novo conte&uacute;do para '" & Server.HTMLEncode(node.Name) & "':"
			else
				prompt = "Editar conte&uacute;do para '" & Server.HTMLEncode(node.Name) & "':"
			end if
%>
<br />
<span><%=prompt%></span><br />
<form method="post" action="<%= ASP_PAGE %>" onsubmit="_confirm=false;" name="editor" id="editor">
   <input type="hidden" name="page" value="<%= node.Name %>" />
   <textarea rows="20" cols="40" name="contents" id="contents" class="node-editor"><%= Server.HTMLEncode(node.getText()) %></textarea>
   <br />
   <input type="submit" name="action" value="Salvar" />
   <input type="submit" name="action" value="Cancelar" />
   <input type="submit" name="action" value="Delete" />
</form>
<%
		else
			Response.Write "<div class='node-contents'>"
			Response.Write ContentsToHtml(node.getText())
			Response.Write "</div>"

			theWiki.LastVisited = node.Name
		end if
		
		pageFooter(node)
	end sub


	sub SaveNode(node)

		dim whichButton, proximoNodo
		whichButton = Request.Form("action")
		proximoNodo = node.Name
		
		If whichButton = "Cancelar" Then
			If node.isNew Then
				proximoNodo = theWiki.LastVisited
			End If
		
		Else
			If Not objLogin.IsLogado() Then
				Login("")
				Response.End()
			End If
	
			if whichButton = "Salvar" then
				dim t
				t = Request("contents")
				node.setText(t)
				
			elseif whichButton = "Delete" then
				node.setText("")
				
			end if
		End If
		
		Response.Redirect ASP_PAGE & "?page=" & Server.URLEncode(proximoNodo)
		Response.End
	end sub


	sub RecentChanges()
		' For how many days is a node new?
		const NEW_CUTOFF = 2

		dim node
		set node = new SystemNode
		node.init "Mudanças Recentes"
	
		dim nodeManager
		set nodeManager = new WikiNodeManager

		dim changes
		changes = nodeManager.RecentChanges()

		pageHeader(node)

		' Group edits by day, and show the dates
		dim lastDate
		lastDate = ""
		
		dim currentDate
		
		dim dateNow : dateNow = now()
		
		dim group
		dim change
		Response.Write "<div class='recent-changes'>"
		for each group in GroupNodesByDate(changes)
			currentDate = FormatDateTime(group(0).lastModified, vbLongDate)
			Response.Write "<div class='recent-changes-day'>"
			Response.Write "<div class='recent-changes-date'>" & currentDate & "</div>"
			Response.Write "<ul>"
			
			for each change in group
				dim howOld2
				howOld2 = DateDiff("d", change.createdOn, dateNow)
	
				Response.Write "<li>"
				Response.Write RegularWikiLink(change.title)
				if howOld2 <= NEW_CUTOFF then
					Response.Write " (novo)"
				end if
				
				Response.Write " - " & Server.HTMLEncode(change.usuario)
	
				Response.Write "</li>"
			next
			
			Response.Write "</ul></div>"
		next
		Response.Write "</div>"

		pageFooter(node)
	end sub
	
	
	sub Login(mensagem)
		dim node
		set node = new SystemNode
		node.init "Login"
	
		pageHeader(node)
%>
<div class='login'><br />
   <form action="<%= ASP_PAGE %>" method="post" class="form-login">
      <% If Not IsEmpty(Request("page")) Then %>
			<input type="hidden" id="page" name="page" value="<%=Request("page")%>" />
			<% If Not IsEmpty(Request("op")) Then %>
				<input type="hidden" id="op" name="op" value="<%=Request("op")%>" />
			<% End If %>
			<% If Not IsEmpty(Request("action")) Then %>
				<input type="hidden" name="action" value="<%=Request("action")%>" />
			<% End If %>
			<% If Not IsEmpty(Request("contents")) Then %>
				<input type="hidden" name="contents" value="<%=Server.HTMLEncode(Request("contents"))%>" />
			<% End If %>
      <% End If %>
      <label for="usuario">Usu&aacute;rio:<br />
      </label>
      <input type="text" name="usuario" id="usuario" class="form.form-login" value="<%=objLogin.NomeUsuario%>" />
      <br />
      <label for="senha">Senha:<br />
      </label>
      <input type="password" name="senha" id="senha">
      <br />
      <br />
      <input type="submit" value="Entrar" name="button-login" class="form-login" >
      <br />
   </form>
   <% If not IsEmpty(mensagem) Then %>
   <br />
   <span class="mensagemlogin"><%=Server.HTMLEncode(mensagem)%></span><br />
   <% End If %>
</div>
<script type="text/javascript">
<!--
function login_setfocus()
{
	try
	{
		var a = document.getElementById("usuario");
		a.focus();
	} catch(e) {}
}
login_setfocus();
-->
</script>
<%		
		pageFooter(node)
	end sub
	
	
	sub UserWelcome(mensagem)
		If Not IsEmpty(Request("page")) Then
			dim strUrl
			strUrl = ASP_PAGE & "?page=" & Server.URLEncode(Request("page"))
			If Not IsEmpty(Request("op")) Then
				strUrl = strUrl & "&op=" & Server.URLEncode(Request("op"))
			End If
			Response.Redirect strUrl
			Response.End()
		End If
		
		dim node
		set node = new SystemNode
		node.init "Bem vindo, " + Server.HTMLEncode(objLogin.NomeUsuario)	
		pageHeader(node)
%>
<div class='userwelcome'><br />
   <form action="<%= ASP_PAGE %>" method="get" class="form-welcome">
      <label for="senhaatual">Senha atual:<br />
      </label>
      <input type="password" name="senhaatual" id="senhaatual">
      <br />
      <label for="novasenha">Nova senha:<br />
      </label>
      <input type="password" name="novasenha" id="novasenha">
      <br />
      <br />
      <input type="submit" value="Alterar senha" name="button-alterarsenha" class="form-welcome" />
      <br />
   </form>
   <% If not IsEmpty(mensagem) Then %>
   <br />
   <span class="mensagemlogin"><%=Server.HTMLEncode(mensagem)%></span><br />
   <% End If %>
</div>
<%		
		pageFooter(node)
	end sub

	
	sub FormatSearchResults(props)
		dim node
		set node = new SystemNode
		node.init props("page_header")

		dim titles, title
		titles = props("results")

		pageHeader(node)

		Response.Write "<div class='search-results'>"
		Response.Write "<div class='search-header'>"
		Response.Write props("search_header")
		Response.Write "</div>"

		if IsEmpty(titles) then
			Response.Write "<i>Nada foi encontrado.</i>"
		else
			Response.Write "<ol>"
			for each title in titles
				Response.Write "<li>"
				Response.Write RegularWikiLink(title)
				Response.Write "</li>"
			next
			Response.Write "</ol>"
		end if
		
		Response.Write "</div>"

		pageFooter(node)
	end sub


	sub TitleSearch()
		dim nodeManager
		set nodeManager = new WikiNodeManager

		dim q
		q = Request("q")

		dim props
		set props = Server.CreateObject("Scripting.Dictionary")

		props.Add "page_header", "Procurar por Título"
		props.Add "search_header", "Procurar no t&igrave;tulo '" & Server.HTMLEncode(q) & "':"
		props.Add "results", nodeManager.TitleSearch(q)

		FormatSearchResults props
	end sub


	sub FullSearch()
		dim nodeManager
		set nodeManager = new WikiNodeManager

		dim q
		q = Request("q")

		dim props
		set props = Server.CreateObject("Scripting.Dictionary")

		props.Add "page_header", "Procurar por Conteúdo"
		props.Add "search_header", "Procurar no conte&uacute;do '" & Server.HTMLEncode(q) & "':"
		props.Add "results", nodeManager.FullSearch(q)

		FormatSearchResults props
	end sub


    Sub NovoUsuario(mensagem)
        dim strOp
        strOp = Request("button-newuser")
        if strOp > "" then
            mensagem = objLogin.CriarNovoLogin(Request("nome"), Request("senha"))
            If IsEmpty(mensagem) Then
			    dim strUrl
			    strUrl = ASP_PAGE & "?op=users"
			    Response.Redirect strUrl
			    Response.End()
            End If
        end if
		
		dim node
		set node = new SystemNode
		node.init "Criar novo editor"
		pageHeader(node)
%>
<div class='newuser'><br />
   <form action="<%= ASP_PAGE %>" method="get" class="form-newuser">
      <input type="hidden" name="op" value="newuser" />      
       <label for="nome">Novo login:</label><br />      
       <input type="text" name="nome" id="nome" value="<%=Request("nome")%>" />
       <br />             
       <label for="senha">Senha:</label><br />
       <input type="password" name="senha" id="senha" value="<%=Request("senha")%>" />
       <br />
       <br />
       <br />
       <input type="submit" value="Criar novo editor" name="button-newuser" class="form-newuser" />
       <br />
   </form>
   <% If not IsEmpty(mensagem) Then %>
   <br />
   <span class="mensagemlogin"><%=Server.HTMLEncode(mensagem)%></span><br />
   <% End If %>
</div>
<%		
		pageFooter(node)
    End Sub


    Sub Users(mensagem)
	    'If Not IsEmpty(Request("page")) Then
		'    dim strUrl
		'    strUrl = ASP_PAGE & "?page=" & Server.URLEncode(Request("page"))
		'    If Not IsEmpty(Request("op")) Then
		'	    strUrl = strUrl & "&op=" & Server.URLEncode(Request("op"))
		'    End If
		'    Response.Redirect strUrl
		'    Response.End()
	    'End If

		dim conn
		set conn = OpenDatabase()

		dim strSQL
		dim rs

		strSQL = "select Usuarios.nome, count(Wiki.Created_by) as Cont, max(created_on) as Data from Usuarios Left Join wiki on Usuarios.nome = wiki.Created_by group by Usuarios.nome order by count(Wiki.Created_by)"
		set rs = conn.Execute(strSQL)
		
	    dim node
	    set node = new SystemNode
	    node.init "Editores"
	    pageHeader(node)
%>
<div class='users'><br />
   <table border='1'>
       <tr><th>Login</th><th>Artigos</th><th>Ultimo criado</th></tr>
        <% do while not rs.eof
				Response.Write "<tr>"
				Response.Write "<td>" & (Server.HtmlEncode(rs("nome").Value & "")) & "</td>"
                Response.Write "<td>" & (Server.HtmlEncode(rs("cont").Value & "")) & "</td>"
                Response.Write "<td>" & (Server.HtmlEncode(rs("data").Value & "")) & "</td>"
				Response.Write "</tr>"
		    rs.MoveNext
            loop %>
   </table>
</div>
<%	
		pageFooter(node)
	End Sub


	sub BackLinks(node)
		dim nodeManager
		set nodeManager = new WikiNodeManager

		dim props
		set props = Server.CreateObject("Scripting.Dictionary")

		props.Add "page_header", node.Name
		props.Add "search_header", "<b>T&oacute;picos que faze refer&ecirc;ncia a '" & Server.HTMLEncode(node.name) & "':</b>"
		props.Add "results", nodeManager.BackLinks(node.name)

		FormatSearchResults props
	end sub


	' Turn a node's content into displayable HTML
	function ContentsToHTML(byval text)
			dim re
			set re = new RegExp
			re.Global = true

			re.Pattern = vbCRLF & "<"
			text = re.Replace(text,"<")

			re.Pattern = ">" & vbCRLF
			text = re.Replace(text,">")

			' ' Convert =content= in header (h1, h2...)
			' re.Pattern = "\=([^]\n<>]*)\="
			' text = re.Replace(text,GetRef("reMakeHeader"))

			' Convert *content* in bold text
			re.Pattern = "\*([^] \n<>]*)\*"
			text = re.Replace(text,GetRef("reMakeNegrito"))
			
			' Convert wiki words to HTML links
			re.Pattern = "\[([^]\n<>]*)]"
			text = re.Replace(text,GetRef("reMakeWikiLink"))

			' Convert *http URLs into hrefs
			re.Pattern = "\*(http\://\S+)"
			text = re.Replace(text,GetRef("reMakeWebLink"))

			' Turn linefeeds into HTML breaks.
			text = Replace(text, vbCRLF, "<br />" & vbCRLF)			

			ContentsToHTML = text
	end function
end class

' function reMakeHeader(match, content, pos, source)
' 	reMakeHeader = "<h1>" & content & "</h1>"
' end function

function reMakeNegrito(match, content, pos, source)
	reMakeNegrito = "<strong>" & content & "</strong>"
end function

function reMakeWebLink(match, URL, pos, source)
	reMakeWebLink = "<a href=""" & URL & """>" & URL & "</a>"
end function

function reMakeWikiLink(match,linked_page,pos,source)
	' If the linked page doesn't exist yet then we add a ? to the link text
	dim linkedNode
	set linkedNode = theWiki.GetNode(linked_page)

	if linkedNode.isNew then
		reMakeWikiLink = Server.HTMLEncode(linked_page) & "<a href=""" & ASP_PAGE & "?page=" & Server.URLEncode(linked_page) & """ title=""Criar o t&oacute;pico &quot;" & Server.HTMLEncode(linked_page) & "&quot;"">?</a>"
	else
		reMakeWikiLink = RegularWikiLink(linked_page)
	end if
end function

' HTML for a link to an existing Wiki node.
function RegularWikiLink(node_name)
	RegularWikiLink = "<a href=""" & ASP_PAGE & "?page=" & Server.URLEncode(node_name) & """>" & Server.HTMLEncode(node_name) & "</a>"
end function





' Global object for our Wiki
' Unfortunately we use it as a global in MakeLink,
' which can't be both a member of the class Wiki
' and a regex replacement function.
dim theWiki : set theWiki = new Wiki
theWiki.Run()
%>
