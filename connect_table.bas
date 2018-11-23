'Access

'Setando a tabela em uma variável (ORM)
Dim <var> As DAO.Recordset
Set <var> = CurrentDb.OpenRecordset("<table>")

'Adicionando registros a uma tabela
<var>.AddNew
<var>!<field> = <info>
<var>.Update

'Editando registro na tabela
<var>.Edit
<var>!<field> = <info>
<var>.Update

'Percorrendo registros de uma tabela
<var>.MoveFirst
Do Until <var>.EOF
'...
'...
<var>.MoveNext
Loop
