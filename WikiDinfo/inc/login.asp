<%

Dim objLogin
Set objLogin = New Login

Class Login
	Public Function IsLogado()
		isLogado = (Session("SessionID") = Session.SessionID)
	End Function
	
	Public Function Login(nomeUsuario, senha)	   
		dim conn
		set conn = OpenDatabase()

		dim sqlUsuario, sqlSenha
		sqlUsuario = SQLescape(nomeUsuario)
		sqlSenha = SQLescape(senha)

		dim strSQL
		strSQL = "select * from Usuarios where nome='" & sqlUsuario & "' and senha='" & sqlSenha & "'"
		
		dim rs
		set rs = conn.Execute(strSQL)

		if rs.eof then
			Session("SessionID") = Empty
			Login = False
		else
			Session("SessionID") = Session.SessionID
			Session("NomeUsuario") = rs.Fields("Nome")
			Login = True
		end if
	End Function
	
	Public Function AlterarSenha(senhaAtual, novaSenha)
		AlterarSenha = False
		If Not IsLogado() Then Exit Function

		dim conn
		set conn = OpenDatabase()

		dim sqlUsuario, sqlSenha, sqlNovaSenha
		sqlUsuario = SQLescape(NomeUsuario)
		sqlSenha = SQLescape(senhaAtual)
		sqlNovaSenha = SQLescape(novaSenha)

		dim strSQL
		strSQL = "select * from Usuarios where nome='" & sqlUsuario & "' and Senha='" & sqlSenha & "'"
		
		dim rs
		set rs = conn.Execute(strSQL)

		if rs.eof then
			Exit Function
		else
			strSQL = "update Usuarios set Senha='" & sqlNovaSenha & "' where nome='" & sqlUsuario & "'"
			set rs = conn.Execute(strSQL)
			AlterarSenha = True
		end if
	End Function
	
	Public Sub Logoff()
		Session("SessionID") = Empty
	End Sub	
	
	Public Property Get NomeUsuario()
		NomeUsuario = Session("NomeUsuario")
	End Property	

    Public Function ValidarStringLogin(strLogin)
        ValidarStringLogin = False
        Dim tam : tam = Len(strLogin)
        Dim pos
        If tam > 16 Or tam < 3 Then Exit Function 
        For pos = 1 To tam
            Dim c : c = Asc(Mid(strLogin, pos, 1))
            Dim ok : ok = False
            If pos > 1 And (c >= 48 And c <= 57) Then ok = True   ' 0-9
            If pos > 1 And (c = 95)              Then ok = True   ' _
            If (c >= 65 And c <= 90)             Then ok = True   ' A-Z
            If (c >= 97 And c <= 122)            Then ok = True   ' a-z
            If Not ok Then Exit Function
        Next
        ValidarStringLogin = True
    End Function

    Public Function CriarNovoLogin(strLogin, strSenha)
        If Not ValidarStringLogin(strLogin) Then
            CriarNovoLogin = "Login inválido. Permitido letras, números e '_'. Deve começar por letra."
            Exit Function
        End If

        If Len(strSenha) < 4 Then
            CriarNovoLogin = "Senha muito curta. Digite pelo menos 4 caracteres."
            Exit Function
        End If
 
        dim conn
		set conn = OpenDatabase()

		dim sqlLogin, sqlSenha
		sqlLogin = SQLescape(strLogin)
		sqlSenha = SQLescape(strSenha)

		dim strSQL
		strSQL = "select * from Usuarios where nome='" & sqlLogin & "'"
		
		dim rs
		set rs = conn.Execute(strSQL)

		If Not rs.eof then
            CriarNovoLogin = "Já existe o login informado."
            Exit Function			
		End If

        strSQL = "insert into Usuarios (nome, senha) values ('" & sqlLogin & "','" & sqlSenha & "')"
		set rs = conn.Execute(strSQL)
   End Function
End Class
%>