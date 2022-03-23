component {

    remote function upload(required string file_name){
        if ( len(trim(arguments.file_name)) ) {
            cffile( fileField="file_name", destination="F:\Coldfushion\cfusion\wwwroot\Userdetails\files", action="upload", result="fileUploadResult",nameConflict ="overwrite");
            //taking a copy for making result file
            cffile( fileField="file_name", destination="F:\Coldfushion\cfusion\wwwroot\Userdetails\files/result_"&fileUploadResult.clientFile, action="upload", result="fileUploadResult1",nameConflict ="overwrite");
            cfspreadsheet( action="read", src="../files/"&fileUploadResult.clientFile,query="importdata",excludeHeaderRow="true");

            //to list roles available in the system
            local.role_list = queryExecute("SELECT role FROM roles");
            local.arrayList = ValueArray(role_list,"role");
       
            local.i = 2;
            local.statusflag  = "";

            cfloop( query="importdata",startrow=2) {
                if ( (importdata.col_1 != "") || (importdata.col_2 != "") || (importdata.col_3 != "") || (importdata.col_4 != "") || (importdata.col_5 != "") || (importdata.col_6 != "") || (importdata.col_7 != "") ) {
                    local.statusflag  = "";
                    if ( (importdata.col_1 != "") && (importdata.col_2 != "") && (importdata.col_3 != "") && (importdata.col_4 != "") && (importdata.col_5 != "") && (importdata.col_6 != "") && (importdata.col_7 != "") ) {
                        local.statusflag  = "";
                        local.role_chk    = "";
                        local.rolearray1  = "";
                        rolearray1        = listToArray(importdata.col_7);
                        local.flag = 0;
                
                        for( itm in rolearray1 ) {

                            if(arrayFind(arrayList,itm) == 0)
                            {
                                flag = 1;
                            }
                
                        }
                        if(flag == 0)
                        {
                            local.statusflag  = "";
                            email_chk = queryExecute(
                                "SELECT id FROM users WHERE email = :email_id;", 
                                {
                                    email_id   : { cfsqltype: "cf_sql_varchar", value: importdata.col_4}
                                }
                            );

                            if(email_chk.RecordCount > 0){
                                statusflag  = "update";
                                queryExecute("
                                        UPDATE users
                                        SET firstname = :firstname, lastname= :lastname, address= :address, email= :email, phone= :phone, dob= :dob,role_id = :role_id
                                        WHERE id = :ID;",
                                    {
                                        firstname: { cfsqltype: "cf_sql_varchar", value: importdata.col_1 },
                                        lastname: { cfsqltype: "cf_sql_varchar", value: importdata.col_2 },
                                        address: { cfsqltype: "cf_sql_varchar", value: importdata.col_3 },
                                        email: { cfsqltype: "cf_sql_varchar", value: importdata.col_4 },
                                        phone: { cfsqltype: "cf_sql_varchar", value: importdata.col_5 },
                                        dob: { cfsqltype: "cf_sql_varchar", value: DateFormat(importdata.col_6,"yyy-mm-dd") },
                                        role_id: { cfsqltype: "cf_sql_varchar", value: importdata.col_7 },
                                        ID: { cfsqltype: "cf_sql_varchar", value: email_chk.id}

                                    }
                                );    
                            }else{
                                statusflag  = "add";
                                queryExecute("
                                    insert into users(firstname,lastname,address,email,phone,dob,role_id)
                                    values( :firstname, :lastname, :address, :email, :phone, :dob, :role_id )",
                                    {
                                        firstname: { cfsqltype: "cf_sql_varchar", value: importdata.col_1 },
                                        lastname: { cfsqltype: "cf_sql_varchar", value: importdata.col_2 },
                                        address: { cfsqltype: "cf_sql_varchar", value: importdata.col_3 },
                                        email: { cfsqltype: "cf_sql_varchar", value: importdata.col_4 },
                                        phone: { cfsqltype: "cf_sql_varchar", value: importdata.col_5 },
                                        dob: { cfsqltype: "cf_sql_varchar", value: DateFormat(importdata.col_6,"yyy-mm-dd") },
                                        role_id: { cfsqltype: "cf_sql_varchar", value: importdata.col_7 }
                                    }
                                );
                            }

                        }
                    }
                    local.failList="";
                    if(importdata.col_1 == ""){
                        failList = listAppend(failList, "First name missing");
                    }
                    if(importdata.col_2 == ""){
                        failList = listAppend(failList, "Last name missing");
                    }
                    if(importdata.col_3 == ""){
                        failList = listAppend(failList, "Address missing");
                    }
                    if(importdata.col_4 == ""){
                        failList = listAppend(failList, "Email missing");
                    }
                    if(importdata.col_5 == ""){
                        failList = listAppend(failList, "Phone missing");
                    }
                    if(importdata.col_6 == ""){
                        failList = listAppend(failList, "DOB missing");
                    }
                    if(importdata.col_7 == ""){
                        failList = listAppend(failList, "role missing");
                    }

                    local.rolearray = "";
                    rolearray       = listToArray(importdata.col_7);
                    for( itm1 in rolearray ) {

                        if(arrayFind(arrayList,itm1) == 0)
                        {
                            failList = listAppend(failList, itm1&" is not a valid role");
                        }

                    }
                    if(statusflag == "add")
                    {
                        failList = listAppend(failList, "successfully added");
                        
                    }
                    if(statusflag == "update")
                    {
                        failList = listAppend(failList, "successfully updated");
                    }

                    local.spObj          = spreadsheetread("F:\Coldfushion\cfusion\wwwroot\Userdetails\files/result_"&fileUploadResult.clientFile,"Sheet1");
                    local.myFormat       = StructNew();
                    local.myFormat.bold  ="true";
                    // populate Sheet 
                    SpreadsheetSetCellValue(spObj, "Result",1, 8);
                    SpreadsheetSetCellValue(spObj,failList,i, 8);
                    SpreadsheetFormatRow(spObj,myFormat,1);
                    SpreadsheetWrite(spObj,"F:\Coldfushion\cfusion\wwwroot\Userdetails\files/result_"&fileUploadResult.clientFile,"yes");
                }
                i++;
            }
            location("../pages/home.cfm?name=result_"&fileUploadResult.clientFile, "false")
        }
    }
    remote function downloadResult(required string file_name){
        local.spObj          = spreadsheetNew("Sheet1",true);
        spreadsheetAddRow(spObj, "First Name,Last name,Address,Email,Phone,DOB,Role,Result");
        local.myFormat       =StructNew();
        local.myFormat.bold  ="true";
        cfspreadsheet( action="read", src="F:\Coldfushion\cfusion\wwwroot\Userdetails\files/"&arguments.file_name,query="importdata",excludeHeaderRow="true");

        myStruct=structNew();
        i=1;  y=2; z=1;
        for ( j in importdata ) {
            if(i>1){
                if ( (j.col_1 != "") || (j.col_2 != "") || (j.col_3 != "") || (j.col_4 != "") || (j.col_5 != "") || (j.col_6 != "") || (j.col_7 != "") ) {    
                    //FILTER DATA WITH Errors
                    if ((j.col_8 != "successfully updated") &&  (j.col_8 != "successfully added")) {

                        SpreadsheetSetCellValue(spObj,j.col_1,y,1);
                        SpreadsheetSetCellValue(spObj,j.col_2,y,2);  
                        SpreadsheetSetCellValue(spObj,j.col_3,y,3);
                        SpreadsheetSetCellValue(spObj,j.col_4,y,4);
                        SpreadsheetSetCellValue(spObj,j.col_5,y,5);
                        SpreadsheetSetCellValue(spObj,j.col_6,y,6); 
                        SpreadsheetSetCellValue(spObj,j.col_7,y,7);
                        SpreadsheetSetCellValue(spObj,j.col_8,y,8);  
                    y++;
                    local.c=y;  
                    }
                }
               
            }
        i++;
        }
        local.v=c;
        for ( z in importdata ) {
            if(v>1){
                if ( (z.col_1 != "") || (z.col_2 != "") || (z.col_3 != "") || (z.col_4 != "") || (z.col_5 != "") || (z.col_6 != "") || (z.col_7 != "") ) {    
                    //FILTER DATA WITH NO ERRORS
                    if ((z.col_8 == "successfully updated") ||  (z.col_8 == "successfully added")) {    
                        SpreadsheetSetCellValue(spObj,z.col_1,v,1);
                        SpreadsheetSetCellValue(spObj,z.col_2,v,2);
                        SpreadsheetSetCellValue(spObj,z.col_3,v,3);
                        SpreadsheetSetCellValue(spObj,z.col_4,v,4);
                        SpreadsheetSetCellValue(spObj,z.col_5,v,5);
                        SpreadsheetSetCellValue(spObj,z.col_6,v,6);
                        SpreadsheetSetCellValue(spObj,z.col_7,v,7);
                        SpreadsheetSetCellValue(spObj,z.col_8,v,8);

                    v++; 
                    }
                }
               
            }
        i++;
        }

        SpreadsheetFormatRow(spObj,myFormat,1);
        SpreadsheetWrite(spObj,"F:\Coldfushion\cfusion\wwwroot\Userdetails\files/result_file.xlsx","yes");
        cfheader( name="Content-Disposition", value="inline; filename=F:\Coldfushion\cfusion\wwwroot\Userdetails\files/"&arguments.file_name );
        cfcontent( type="application/vnd.msexcel",variable=SpreadSheetReadBinary(spObj));  
    }

    public function listUsers(){
        local.lists = queryExecute("SELECT* FROM users");
        savecontent variable="usersList" {
            writeOutput("<table class=""table"">
            <thead>
                <tr>
                    <th scope=""col"">First Name</th>
                    <th scope=""col"">Last Name</th>
                    <th scope=""col"">Email</th>
                    <th scope=""col"">Phone</th>
                </tr>
            </thead>
            <tbody>");
            for ( x in lists ) {
                cfoutput(  ) {

                    writeOutput("<tr>
                        <td>#x.firstname#</td>
                        <td>#x.lastname#</td>
                        <td>#x.email#</td>
                        <td>#x.phone#</td>");
                    }
            }
            writeOutput("</tbody>
            </table>");
        }
        return usersList; 
    }

    remote function downloadData(){
        local.lists          = queryExecute("SELECT firstname,lastname,address,email,phone,dob,role_id FROM users");
        local.spObj          = spreadsheetNew("Sheet1",true);
        local.myFormat       = StructNew();
        local.myFormat.bold  ="true";
        spreadsheetAddrows(spObj,lists,1,1,true,[""],true);
        SpreadsheetFormatRow(spObj,myFormat,1);
        //SpreadsheetWrite(spObj,"F:\Coldfushion\cfusion\wwwroot\Userdetails\files/result_withdatafile.xlsx","yes");
        cfheader( name="Content-Disposition", value="inline; filename=result_withdatafile.xlsx");
        cfcontent( type="application/vnd.msexcel",variable=SpreadSheetReadBinary(spObj));  

    }
    remote function plainTemplate(){ 
        local.spObj         = spreadsheetNew("Sheet1",true);
        local.myFormat      = StructNew();
        local.myFormat.bold = "true";
        spreadsheetAddRow(spObj, "First Name,Last name,Address,Email,Phone,DOB,Role");
        SpreadsheetFormatRow(spObj,myFormat,1);
        cfheader( name="Content-Disposition", value="inline; filename=plain_templatefile.xlsx");
        cfcontent( type="application/vnd.msexcel",variable=SpreadSheetReadBinary(spObj));  
    }      
}
