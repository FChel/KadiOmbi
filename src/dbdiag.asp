<%@ Language="VBScript" %>
<%
Option Explicit

Response.Buffer = True
Response.ContentType = "text/html"
Server.ScriptTimeout = 180

Const UDL_RELATIVE_PATH = "Database/CAPS.udl"

Dim udlPath
udlPath = Server.MapPath(UDL_RELATIVE_PATH)

Dim selectedDb, actionName
selectedDb = Trim(Request.Form("databaseName"))
actionName = Trim(Request.Form("actionName"))

Call WriteHeader()

Call WriteIntro()

Call WriteDatabaseSelector(selectedDb)

If actionName <> "" Then
    Response.Write "<hr>"
    Response.Write "<h2>Diagnostic Result</h2>"
    Response.Write "<p><strong>Selected database:</strong> <code>" & Html(selectedDb) & "</code></p>"
    Response.Write "<p><strong>Test:</strong> <code>" & Html(actionName) & "</code></p>"

    If selectedDb = "" Then
        Response.Write "<p class='bad'>No database selected.</p>"
    Else
        Select Case LCase(actionName)
            Case "connection"
                Call TestConnection(selectedDb)

            Case "simplequery"
                Call TestSimpleQuery(selectedDb)

            Case "permissions"
                Call TestPermissions(selectedDb)

            Case "blocking"
                Call TestBlocking(selectedDb)

            Case "activerequests"
                Call TestActiveRequests(selectedDb)

            Case "tables"
                Call TestTables(selectedDb)

            Case "rowcounts"
                Call TestRowCounts(selectedDb)

            Case "slowquery"
                Call TestArtificialSlowQuery(selectedDb)

            Case "allbasic"
                Call TestConnection(selectedDb)
                Call TestSimpleQuery(selectedDb)
                Call TestPermissions(selectedDb)
                Call TestBlocking(selectedDb)
                Call TestActiveRequests(selectedDb)
                Call TestTables(selectedDb)

            Case Else
                Response.Write "<p class='bad'>Unknown action.</p>"
        End Select
    End If
End If

Call WriteFooter()


Sub WriteHeader()
    Response.Write "<!doctype html>"
    Response.Write "<html>"
    Response.Write "<head>"
    Response.Write "<meta charset='utf-8'>"
    Response.Write "<title>SQL Database Diagnostic Page</title>"
    Response.Write "<style>"
    Response.Write "body{font-family:Arial, sans-serif;font-size:14px;margin:30px;color:#222;}"
    Response.Write "h1{margin-bottom:6px;}"
    Response.Write "h2{margin-top:24px;margin-bottom:8px;}"
    Response.Write "h3{margin-top:18px;margin-bottom:8px;}"
    Response.Write ".small{font-size:12px;color:#666;}"
    Response.Write ".ok{color:green;font-weight:bold;}"
    Response.Write ".bad{color:red;font-weight:bold;}"
    Response.Write ".warn{color:#9a6500;font-weight:bold;}"
    Response.Write ".box{border:1px solid #ccc;background:#fafafa;padding:14px;margin:14px 0;}"
    Response.Write ".buttons{margin-top:12px;display:flex;flex-wrap:wrap;gap:8px;}"
    Response.Write "button{padding:8px 12px;cursor:pointer;}"
    Response.Write "select{padding:6px;min-width:280px;}"
    Response.Write "table{border-collapse:collapse;margin-top:10px;margin-bottom:18px;}"
    Response.Write "td,th{border:1px solid #ccc;padding:6px 10px;text-align:left;vertical-align:top;}"
    Response.Write "th{background:#f0f0f0;}"
    Response.Write "code{background:#eee;padding:2px 4px;}"
    Response.Write ".mono{font-family:Consolas, monospace;white-space:pre-wrap;}"
    Response.Write "</style>"
    Response.Write "</head>"
    Response.Write "<body>"
End Sub


Sub WriteIntro()
    Response.Write "<h1>SQL Database Diagnostic Page</h1>"
    Response.Write "<p>UDL relative path: <code>" & Html(UDL_RELATIVE_PATH) & "</code></p>"
    Response.Write "<p>UDL physical path: <code>" & Html(udlPath) & "</code></p>"
    Response.Write "<p class='warn'><strong>Security note:</strong> remove this page after use. It exposes server, login, database and SQL diagnostic information.</p>"
End Sub


Sub WriteDatabaseSelector(currentDb)
    Dim conn, rs, sql

    Response.Write "<div class='box'>"
    Response.Write "<form method='post'>"
    Response.Write "<h2>Select Database</h2>"

    On Error Resume Next

    Set conn = Server.CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = 15
    conn.CommandTimeout = 60
    conn.Open "File Name=" & udlPath

    If Err.Number <> 0 Then
        Response.Write "<p class='bad'>Could not open UDL connection.</p>"
        Response.Write "<p><strong>Error:</strong> " & Html(Err.Description) & "</p>"
        Err.Clear
    Else
        sql = "SELECT name FROM sys.databases WHERE HAS_DBACCESS(name) = 1 ORDER BY name"
        Set rs = conn.Execute(sql)

        If Err.Number <> 0 Then
            Response.Write "<p class='bad'>Could not list databases.</p>"
            Response.Write "<p><strong>Error:</strong> " & Html(Err.Description) & "</p>"
            Err.Clear
        Else
            Response.Write "<select name='databaseName'>"

            Do Until rs.EOF
                Dim dbName, selectedText
                dbName = CStr(rs("name"))

                selectedText = ""
                If LCase(dbName) = LCase(currentDb) Then
                    selectedText = " selected"
                End If

                Response.Write "<option value='" & HtmlAttr(dbName) & "'" & selectedText & ">" & Html(dbName) & "</option>"
                rs.MoveNext
            Loop

            Response.Write "</select>"
        End If

        If Not rs Is Nothing Then
            If rs.State = 1 Then rs.Close
        End If

        conn.Close
    End If

    Set rs = Nothing
    Set conn = Nothing

    On Error GoTo 0

    Response.Write "<div class='buttons'>"
    Response.Write "<button type='submit' name='actionName' value='connection'>Test connection timing</button>"
    Response.Write "<button type='submit' name='actionName' value='simplequery'>Test simple query</button>"
    Response.Write "<button type='submit' name='actionName' value='permissions'>Check permissions</button>"
    Response.Write "<button type='submit' name='actionName' value='blocking'>Check blocking</button>"
    Response.Write "<button type='submit' name='actionName' value='activerequests'>Check active SQL requests</button>"
    Response.Write "<button type='submit' name='actionName' value='tables'>List tables/views</button>"
    Response.Write "<button type='submit' name='actionName' value='rowcounts'>Estimate row counts</button>"
    Response.Write "<button type='submit' name='actionName' value='slowquery'>Test 35-sec command timeout</button>"
    Response.Write "<button type='submit' name='actionName' value='allbasic'>Run basic checks</button>"
    Response.Write "</div>"

    Response.Write "</form>"
    Response.Write "</div>"
End Sub


Sub TestConnection(dbName)
    Dim conn, t0, t1, sql

    Response.Write "<h3>Connection Timing</h3>"

    On Error Resume Next

    t0 = Timer()

    Set conn = Server.CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = 15
    conn.CommandTimeout = 60
    conn.Open "File Name=" & udlPath

    If Err.Number <> 0 Then
        t1 = Timer()
        Response.Write "<p class='bad'>Connection failed.</p>"
        Response.Write "<p><strong>Open time:</strong> " & FormatSeconds(ElapsedSeconds(t0, t1)) & "</p>"
        Response.Write "<p><strong>Error:</strong> " & Html(Err.Description) & "</p>"
        Err.Clear
    Else
        t1 = Timer()
        Response.Write "<p class='ok'>Connection opened successfully.</p>"
        Response.Write "<p><strong>Open time:</strong> " & FormatSeconds(ElapsedSeconds(t0, t1)) & "</p>"

        sql = "USE [" & EscapeDbName(dbName) & "]"
        conn.Execute sql

        If Err.Number <> 0 Then
            Response.Write "<p class='bad'>Connected to SQL Server, but could not switch to selected database.</p>"
            Response.Write "<p><strong>Error:</strong> " & Html(Err.Description) & "</p>"
            Err.Clear
        Else
            Response.Write "<p class='ok'>Selected database is accessible.</p>"
        End If

        conn.Close
    End If

    Set conn = Nothing

    On Error GoTo 0
End Sub


Sub TestSimpleQuery(dbName)
    Dim conn, rs, t0, t1, sql

    Response.Write "<h3>Simple Query Timing</h3>"

    Set conn = OpenDbConnection(dbName)

    If conn Is Nothing Then Exit Sub

    On Error Resume Next

    sql = "SELECT " & _
          "DB_NAME() AS CurrentDatabase, " & _
          "SUSER_SNAME() AS LoginName, " & _
          "USER_NAME() AS DatabaseUser, " & _
          "GETDATE() AS SqlServerTime"

    t0 = Timer()
    Set rs = conn.Execute(sql)
    t1 = Timer()

    If Err.Number <> 0 Then
        Response.Write "<p class='bad'>Simple query failed.</p>"
        Response.Write "<p><strong>Query time:</strong> " & FormatSeconds(ElapsedSeconds(t0, t1)) & "</p>"
        Response.Write "<p><strong>Error:</strong> " & Html(Err.Description) & "</p>"
        Err.Clear
    Else
        Response.Write "<p class='ok'>Simple query succeeded.</p>"
        Response.Write "<p><strong>Query time:</strong> " & FormatSeconds(ElapsedSeconds(t0, t1)) & "</p>"

        Response.Write "<table>"
        Response.Write "<tr><th>Current database</th><th>Login</th><th>Database user</th><th>SQL Server time</th></tr>"

        If Not rs.EOF Then
            Response.Write "<tr>"
            Response.Write "<td>" & Html(Nz(rs("CurrentDatabase"))) & "</td>"
            Response.Write "<td>" & Html(Nz(rs("LoginName"))) & "</td>"
            Response.Write "<td>" & Html(Nz(rs("DatabaseUser"))) & "</td>"
            Response.Write "<td>" & Html(Nz(rs("SqlServerTime"))) & "</td>"
            Response.Write "</tr>"
        End If

        Response.Write "</table>"
    End If

    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
    End If

    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    On Error GoTo 0
End Sub


Sub TestPermissions(dbName)
    Dim conn, rs, sql

    Response.Write "<h3>Permissions Check</h3>"

    Set conn = OpenDbConnection(dbName)

    If conn Is Nothing Then Exit Sub

    On Error Resume Next

    sql = "SELECT " & _
          "DB_NAME() AS CurrentDatabase, " & _
          "SUSER_SNAME() AS LoginName, " & _
          "USER_NAME() AS DatabaseUser, " & _
          "IS_MEMBER('db_owner') AS IsDbOwner, " & _
          "IS_MEMBER('db_datareader') AS IsDataReader, " & _
          "IS_MEMBER('db_datawriter') AS IsDataWriter, " & _
          "HAS_PERMS_BY_NAME(DB_NAME(), 'DATABASE', 'SELECT') AS HasDatabaseSelect, " & _
          "HAS_PERMS_BY_NAME(DB_NAME(), 'DATABASE', 'INSERT') AS HasDatabaseInsert, " & _
          "HAS_PERMS_BY_NAME(DB_NAME(), 'DATABASE', 'UPDATE') AS HasDatabaseUpdate, " & _
          "HAS_PERMS_BY_NAME(DB_NAME(), 'DATABASE', 'DELETE') AS HasDatabaseDelete, " & _
          "HAS_PERMS_BY_NAME(DB_NAME(), 'DATABASE', 'EXECUTE') AS HasDatabaseExecute"

    Set rs = conn.Execute(sql)

    If Err.Number <> 0 Then
        Response.Write "<p class='bad'>Permissions query failed.</p>"
        Response.Write "<p><strong>Error:</strong> " & Html(Err.Description) & "</p>"
        Err.Clear
    Else
        Response.Write "<table>"
        Response.Write "<tr><th>Permission / Role</th><th>Result</th></tr>"

        If Not rs.EOF Then
            Response.Write "<tr><td>Current database</td><td>" & Html(Nz(rs("CurrentDatabase"))) & "</td></tr>"
            Response.Write "<tr><td>SQL login</td><td>" & Html(Nz(rs("LoginName"))) & "</td></tr>"
            Response.Write "<tr><td>Database user</td><td>" & Html(Nz(rs("DatabaseUser"))) & "</td></tr>"
            Response.Write "<tr><td>db_owner</td><td>" & YesNo(rs("IsDbOwner")) & "</td></tr>"
            Response.Write "<tr><td>db_datareader</td><td>" & YesNo(rs("IsDataReader")) & "</td></tr>"
            Response.Write "<tr><td>db_datawriter</td><td>" & YesNo(rs("IsDataWriter")) & "</td></tr>"
            Response.Write "<tr><td>Database SELECT</td><td>" & YesNo(rs("HasDatabaseSelect")) & "</td></tr>"
            Response.Write "<tr><td>Database INSERT</td><td>" & YesNo(rs("HasDatabaseInsert")) & "</td></tr>"
            Response.Write "<tr><td>Database UPDATE</td><td>" & YesNo(rs("HasDatabaseUpdate")) & "</td></tr>"
            Response.Write "<tr><td>Database DELETE</td><td>" & YesNo(rs("HasDatabaseDelete")) & "</td></tr>"
            Response.Write "<tr><td>Database EXECUTE</td><td>" & YesNo(rs("HasDatabaseExecute")) & "</td></tr>"
        End If

        Response.Write "</table>"
    End If

    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
    End If

    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    On Error GoTo 0
End Sub


Sub TestBlocking(dbName)
    Dim conn, rs, sql

    Response.Write "<h3>Blocking Check</h3>"
    Response.Write "<p class='small'>This checks SQL Server sessions currently blocked in the selected database. Requires permission to view server state, otherwise it may return limited results or fail.</p>"

    Set conn = OpenDbConnection(dbName)

    If conn Is Nothing Then Exit Sub

    On Error Resume Next

    sql = "SELECT " & _
          "r.session_id, " & _
          "r.blocking_session_id, " & _
          "r.status, " & _
          "r.wait_type, " & _
          "r.wait_time, " & _
          "r.command, " & _
          "DB_NAME(r.database_id) AS database_name, " & _
          "LEFT(t.text, 2000) AS sql_text " & _
          "FROM sys.dm_exec_requests r " & _
          "OUTER APPLY sys.dm_exec_sql_text(r.sql_handle) t " & _
          "WHERE r.database_id = DB_ID() " & _
          "AND r.blocking_session_id <> 0 " & _
          "ORDER BY r.wait_time DESC"

    Set rs = conn.Execute(sql)

    If Err.Number <> 0 Then
        Response.Write "<p class='bad'>Blocking query failed.</p>"
        Response.Write "<p><strong>Error:</strong> " & Html(Err.Description) & "</p>"
        Response.Write "<p class='small'>This is usually a SQL permission issue, not necessarily an application issue.</p>"
        Err.Clear
    Else
        If rs.EOF Then
            Response.Write "<p class='ok'>No currently blocked requests found for this database.</p>"
        Else
            Response.Write "<table>"
            Response.Write "<tr><th>Session</th><th>Blocked by</th><th>Status</th><th>Wait type</th><th>Wait ms</th><th>Command</th><th>SQL text</th></tr>"

            Do Until rs.EOF
                Response.Write "<tr>"
                Response.Write "<td>" & Html(Nz(rs("session_id"))) & "</td>"
                Response.Write "<td class='bad'>" & Html(Nz(rs("blocking_session_id"))) & "</td>"
                Response.Write "<td>" & Html(Nz(rs("status"))) & "</td>"
                Response.Write "<td>" & Html(Nz(rs("wait_type"))) & "</td>"
                Response.Write "<td>" & Html(Nz(rs("wait_time"))) & "</td>"
                Response.Write "<td>" & Html(Nz(rs("command"))) & "</td>"
                Response.Write "<td class='mono'>" & Html(Nz(rs("sql_text"))) & "</td>"
                Response.Write "</tr>"
                rs.MoveNext
            Loop

            Response.Write "</table>"
        End If
    End If

    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
    End If

    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    On Error GoTo 0
End Sub


Sub TestActiveRequests(dbName)
    Dim conn, rs, sql

    Response.Write "<h3>Active SQL Requests</h3>"
    Response.Write "<p class='small'>Shows active requests currently running in the selected database. Requires permission to view server state for full detail.</p>"

    Set conn = OpenDbConnection(dbName)

    If conn Is Nothing Then Exit Sub

    On Error Resume Next

    sql = "SELECT TOP 50 " & _
          "r.session_id, " & _
          "r.status, " & _
          "r.command, " & _
          "r.cpu_time, " & _
          "r.total_elapsed_time, " & _
          "r.reads, " & _
          "r.writes, " & _
          "r.logical_reads, " & _
          "r.wait_type, " & _
          "r.blocking_session_id, " & _
          "LEFT(t.text, 2000) AS sql_text " & _
          "FROM sys.dm_exec_requests r " & _
          "OUTER APPLY sys.dm_exec_sql_text(r.sql_handle) t " & _
          "WHERE r.database_id = DB_ID() " & _
          "AND r.session_id <> @@SPID " & _
          "ORDER BY r.total_elapsed_time DESC"

    Set rs = conn.Execute(sql)

    If Err.Number <> 0 Then
        Response.Write "<p class='bad'>Active request query failed.</p>"
        Response.Write "<p><strong>Error:</strong> " & Html(Err.Description) & "</p>"
        Response.Write "<p class='small'>This is usually a SQL permission issue.</p>"
        Err.Clear
    Else
        If rs.EOF Then
            Response.Write "<p class='ok'>No other active requests found for this database.</p>"
        Else
            Response.Write "<table>"
            Response.Write "<tr><th>Session</th><th>Status</th><th>Command</th><th>CPU ms</th><th>Elapsed ms</th><th>Reads</th><th>Writes</th><th>Logical reads</th><th>Wait</th><th>Blocked by</th><th>SQL text</th></tr>"

            Do Until rs.EOF
                Response.Write "<tr>"
                Response.Write "<td>" & Html(Nz(rs("session_id"))) & "</td>"
                Response.Write "<td>" & Html(Nz(rs("status"))) & "</td>"
                Response.Write "<td>" & Html(Nz(rs("command"))) & "</td>"
                Response.Write "<td>" & Html(Nz(rs("cpu_time"))) & "</td>"
                Response.Write "<td>" & Html(Nz(rs("total_elapsed_time"))) & "</td>"
                Response.Write "<td>" & Html(Nz(rs("reads"))) & "</td>"
                Response.Write "<td>" & Html(Nz(rs("writes"))) & "</td>"
                Response.Write "<td>" & Html(Nz(rs("logical_reads"))) & "</td>"
                Response.Write "<td>" & Html(Nz(rs("wait_type"))) & "</td>"
                Response.Write "<td>" & Html(Nz(rs("blocking_session_id"))) & "</td>"
                Response.Write "<td class='mono'>" & Html(Nz(rs("sql_text"))) & "</td>"
                Response.Write "</tr>"
                rs.MoveNext
            Loop

            Response.Write "</table>"
        End If
    End If

    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
    End If

    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    On Error GoTo 0
End Sub


Sub TestTables(dbName)
    Dim conn, rs, sql

    Response.Write "<h3>Tables and Views</h3>"

    Set conn = OpenDbConnection(dbName)

    If conn Is Nothing Then Exit Sub

    On Error Resume Next

    sql = "SELECT TOP 200 " & _
          "TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE " & _
          "FROM INFORMATION_SCHEMA.TABLES " & _
          "ORDER BY TABLE_SCHEMA, TABLE_TYPE, TABLE_NAME"

    Set rs = conn.Execute(sql)

    If Err.Number <> 0 Then
        Response.Write "<p class='bad'>Table/view listing failed.</p>"
        Response.Write "<p><strong>Error:</strong> " & Html(Err.Description) & "</p>"
        Err.Clear
    Else
        If rs.EOF Then
            Response.Write "<p class='warn'>No tables or views visible to this login.</p>"
        Else
            Response.Write "<table>"
            Response.Write "<tr><th>Schema</th><th>Name</th><th>Type</th></tr>"

            Do Until rs.EOF
                Response.Write "<tr>"
                Response.Write "<td>" & Html(Nz(rs("TABLE_SCHEMA"))) & "</td>"
                Response.Write "<td>" & Html(Nz(rs("TABLE_NAME"))) & "</td>"
                Response.Write "<td>" & Html(Nz(rs("TABLE_TYPE"))) & "</td>"
                Response.Write "</tr>"
                rs.MoveNext
            Loop

            Response.Write "</table>"
            Response.Write "<p class='small'>Limited to first 200 objects.</p>"
        End If
    End If

    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
    End If

    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    On Error GoTo 0
End Sub


Sub TestRowCounts(dbName)
    Dim conn, rs, sql

    Response.Write "<h3>Estimated Row Counts</h3>"
    Response.Write "<p class='small'>Uses SQL metadata. This avoids running SELECT COUNT(*) across every table.</p>"

    Set conn = OpenDbConnection(dbName)

    If conn Is Nothing Then Exit Sub

    On Error Resume Next

    sql = "SELECT TOP 100 " & _
          "s.name AS schema_name, " & _
          "t.name AS table_name, " & _
          "SUM(p.rows) AS estimated_rows " & _
          "FROM sys.tables t " & _
          "INNER JOIN sys.schemas s ON t.schema_id = s.schema_id " & _
          "INNER JOIN sys.partitions p ON t.object_id = p.object_id " & _
          "WHERE p.index_id IN (0,1) " & _
          "GROUP BY s.name, t.name " & _
          "ORDER BY SUM(p.rows) DESC"

    Set rs = conn.Execute(sql)

    If Err.Number <> 0 Then
        Response.Write "<p class='bad'>Estimated row count query failed.</p>"
        Response.Write "<p><strong>Error:</strong> " & Html(Err.Description) & "</p>"
        Err.Clear
    Else
        If rs.EOF Then
            Response.Write "<p class='warn'>No table row count metadata visible.</p>"
        Else
            Response.Write "<table>"
            Response.Write "<tr><th>Schema</th><th>Table</th><th>Estimated rows</th></tr>"

            Do Until rs.EOF
                Response.Write "<tr>"
                Response.Write "<td>" & Html(Nz(rs("schema_name"))) & "</td>"
                Response.Write "<td>" & Html(Nz(rs("table_name"))) & "</td>"
                Response.Write "<td>" & Html(Nz(rs("estimated_rows"))) & "</td>"
                Response.Write "</tr>"
                rs.MoveNext
            Loop

            Response.Write "</table>"
            Response.Write "<p class='small'>Limited to top 100 largest visible tables.</p>"
        End If
    End If

    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
    End If

    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    On Error GoTo 0
End Sub


Sub TestArtificialSlowQuery(dbName)
    Dim conn, rs, t0, t1, sql

    Response.Write "<h3>Artificial 35-Second Command Timeout Test</h3>"
    Response.Write "<p class='small'>This deliberately waits 35 seconds on SQL Server. The page sets CommandTimeout to 45 seconds, so it should complete. If it fails around 30 seconds, something else is enforcing a 30-second timeout.</p>"

    Set conn = OpenDbConnection(dbName)

    If conn Is Nothing Then Exit Sub

    On Error Resume Next

    conn.CommandTimeout = 45

    sql = "WAITFOR DELAY '00:00:35'; SELECT DB_NAME() AS CurrentDatabase, GETDATE() AS FinishedAt"

    t0 = Timer()
    Set rs = conn.Execute(sql)
    t1 = Timer()

    If Err.Number <> 0 Then
        Response.Write "<p class='bad'>Slow query failed or timed out.</p>"
        Response.Write "<p><strong>Elapsed time:</strong> " & FormatSeconds(ElapsedSeconds(t0, t1)) & "</p>"
        Response.Write "<p><strong>Error:</strong> " & Html(Err.Description) & "</p>"
        Err.Clear
    Else
        Response.Write "<p class='ok'>Slow query completed.</p>"
        Response.Write "<p><strong>Elapsed time:</strong> " & FormatSeconds(ElapsedSeconds(t0, t1)) & "</p>"

        If Not rs.EOF Then
            Response.Write "<table>"
            Response.Write "<tr><th>Current database</th><th>Finished at</th></tr>"
            Response.Write "<tr>"
            Response.Write "<td>" & Html(Nz(rs("CurrentDatabase"))) & "</td>"
            Response.Write "<td>" & Html(Nz(rs("FinishedAt"))) & "</td>"
            Response.Write "</tr>"
            Response.Write "</table>"
        End If
    End If

    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
    End If

    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    On Error GoTo 0
End Sub


Function OpenDbConnection(dbName)
    Dim conn, sql, t0, t1

    On Error Resume Next

    t0 = Timer()

    Set conn = Server.CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = 15
    conn.CommandTimeout = 120
    conn.Open "File Name=" & udlPath

    If Err.Number <> 0 Then
        t1 = Timer()
        Response.Write "<p class='bad'>Could not open SQL connection.</p>"
        Response.Write "<p><strong>Open time:</strong> " & FormatSeconds(ElapsedSeconds(t0, t1)) & "</p>"
        Response.Write "<p><strong>Error:</strong> " & Html(Err.Description) & "</p>"
        Err.Clear
        Set OpenDbConnection = Nothing
        Exit Function
    End If

    sql = "USE [" & EscapeDbName(dbName) & "]"
    conn.Execute sql

    If Err.Number <> 0 Then
        Response.Write "<p class='bad'>Connected to SQL Server, but could not switch to database <code>" & Html(dbName) & "</code>.</p>"
        Response.Write "<p><strong>Error:</strong> " & Html(Err.Description) & "</p>"
        Err.Clear

        If Not conn Is Nothing Then
            If conn.State = 1 Then conn.Close
        End If

        Set OpenDbConnection = Nothing
        Exit Function
    End If

    Set OpenDbConnection = conn

    On Error GoTo 0
End Function


Function EscapeDbName(value)
    EscapeDbName = Replace(CStr(value), "]", "]]")
End Function


Function Html(value)
    Html = Server.HTMLEncode(CStr(value))
End Function


Function HtmlAttr(value)
    Dim s
    s = Server.HTMLEncode(CStr(value))
    s = Replace(s, "'", "&#39;")
    HtmlAttr = s
End Function


Function Nz(value)
    If IsNull(value) Then
        Nz = ""
    Else
        Nz = CStr(value)
    End If
End Function


Function YesNo(value)
    If IsNull(value) Then
        YesNo = "<span class='warn'>Unknown</span>"
    ElseIf CInt(value) = 1 Then
        YesNo = "<span class='ok'>Yes</span>"
    Else
        YesNo = "<span class='bad'>No</span>"
    End If
End Function


Function ElapsedSeconds(startTimer, endTimer)
    If endTimer >= startTimer Then
        ElapsedSeconds = endTimer - startTimer
    Else
        ElapsedSeconds = (86400 - startTimer) + endTimer
    End If
End Function


Function FormatSeconds(value)
    FormatSeconds = FormatNumber(value, 3) & " seconds"
End Function


Sub WriteFooter()
    Response.Write "<hr>"
    Response.Write "<h2>How to read this</h2>"
    Response.Write "<table>"
    Response.Write "<tr><th>Check</th><th>What it tells you</th></tr>"
    Response.Write "<tr><td>Test connection timing</td><td>Whether the UDL opens quickly and whether the selected database can be selected.</td></tr>"
    Response.Write "<tr><td>Test simple query</td><td>Whether SQL can run a basic query quickly. If this is fast, the general connection is probably fine.</td></tr>"
    Response.Write "<tr><td>Check permissions</td><td>Shows the SQL login, database user and common database permissions.</td></tr>"
    Response.Write "<tr><td>Check blocking</td><td>Shows whether current sessions are being blocked by another SQL session.</td></tr>"
    Response.Write "<tr><td>Check active SQL requests</td><td>Shows other currently running SQL requests in that database, useful for seeing long-running queries.</td></tr>"
    Response.Write "<tr><td>List tables/views</td><td>Confirms what objects the app login can see.</td></tr>"
    Response.Write "<tr><td>Estimate row counts</td><td>Shows large tables without running expensive COUNT(*) queries.</td></tr>"
    Response.Write "<tr><td>Test 35-sec command timeout</td><td>Deliberately runs a 35-second SQL wait. Useful to confirm whether a 30-second timeout is coming from ASP, ADO, IIS or the application code.</td></tr>"
    Response.Write "</table>"

    Response.Write "<p class='warn'><strong>Remove this page after testing.</strong></p>"
    Response.Write "</body>"
    Response.Write "</html>"
End Sub
%>