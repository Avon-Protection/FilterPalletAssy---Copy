<%@ Application Language="VB" %>

<script runat="server">

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        'Application("sqlSvr") = "na-cad-sqlrs"
        'Application("sqlSvr2") = "na-cad-AS1"
        'Application("sqlCat") = "APS"
        Application("sqlCatFILTER") = "FILTER_ASSEMBLY"
        Application("sqlSvr3") = "na-cad-db1"
        Application("sqlStrAPS") = "data source=" & Application("sqlSvr") & ";initial catalog=" & Application("sqlCat") & ";packet size=4096;user id=APS-FILTER-PALLET;password=su03AAR$#;Connection Timeout=30;Max Pool Size=100"
        Application("sqlStrAS1APS") = "data source=" & Application("sqlSvr2") & ";initial catalog=" & Application("sqlCat") & ";packet size=4096;user id=APS-FILTER-PALLET;password=su03AAR$#;Connection Timeout=30;Max Pool Size=100"
        Application("sqlStrFILTER") = "data source=" & Application("sqlSvr3") & ";initial catalog=" & Application("sqlCatFILTER") & ";packet size=4096;user id=APS-FILTER-PALLET;password=su03AAR$#;Connection Timeout=30;Max Pool Size=100"
    End Sub
    
    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs on application shutdown
    End Sub
        
    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs when an unhandled error occurs
    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs when a new session is started
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs when a session ends. 
        ' Note: The Session_End event is raised only when the sessionstate mode
        ' is set to InProc in the Web.config file. If session mode is set to StateServer 
        ' or SQLServer, the event is not raised.
    End Sub
       
</script>