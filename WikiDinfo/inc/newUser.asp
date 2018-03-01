<%

Dim objCreate
Set objCreate = New newUser

Class newUser
	
	Public Function newUser(nomeUsuario, senha, confirmaSenha)	   
		dim conn
		set conn = OpenDatabase()

		dim sqlUsuario, sqlSenha, sqlConfirmaSenha
		sqlUsuario = SQLescape(nomeUsuario)
		sqlSenha = SQLescape(senha)
        sqlConfirmaSenha = SQLescape(confirmaSenha)
        

        if sqlSenha = sqlConfirmaSenha then
        
		    dim strSQL
		    strSQL = "select * from Usuarios where nome='" & sqlUsuario & "' and senha='" & sqlSenha & "'"
		
		    dim rs
		    set rs = conn.Execute(strSQL)

		    if rs.eof then
			    strSQL = "insert into Usuarios (nome, senha) values ('"&sqlUsuario&"','"&sqlSenha&"')"
                set rs = conn.Execute(strSQL)
                newUser = True
		    else
			    newUser = false
		    end if
        end if
	End Function
End Class
%>