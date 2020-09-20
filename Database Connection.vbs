'Database connection

Set con=CreateObject("ADODB.Connection")
Set rs=CreateObject("ADODB.RecordSet")
dbQuery="Select * from emp"

con.Open ("Driver={Microsoft Access Driver(.*mdb,.*accdb);dbq=c:/test.mdb;username=sumit;password=}")
rs.Open dbQuery,con

do while rs.EOF<>True
user=rs.Fields("username")
pwd=rs.Fields("password")
msgbox user
msgbox pwd
rs.MoveNext
loop
con.close
rs.close