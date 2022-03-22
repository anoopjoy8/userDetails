component {

    this.name = "User Details";
    this.datasource = "userdetails";
    this.sessionManagement  = true;
    
    function onRequestStart(requestname){ 
        application.rootpath = expandPath('./');
        application.baseUrl='http#iif(CGI.SERVER_PORT_SECURE,"s","")#://#CGI.SERVER_NAME##getDirectoryFromPath(cgi.SCRIPT_NAME)#';
    }

/*     function onError(Exception,EventName){
        writeOutput('<center><h1>An error occurred</h1>
        <p>Please Contact the developer</p>
        <p>Error details: #Exception.message#</p></center>');
    } 
 */
    function onMissingTemplate(targetPage){
        writeOutput('<center><h1>This Page is not avilable.</h1>
        <p>Please go back:</p></center>');
    }

    function onSessionEnd(sessionScope, applicationScope){
        writeOutput('<center>
                     <h1>Your session expired. Please login again</h1>
                     <a href="pages/user/index.cfm">Click Here</a>
                     </center>');
    }
}