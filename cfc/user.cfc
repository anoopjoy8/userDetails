component {

    remote function upload(required string file_name){

        if ( len(trim(arguments.file_name)) ) {
            cffile( fileField="file_name", destination="F:\Coldfushion\cfusion\wwwroot\Userdetails\files", action="upload", result="fileUploadResult",nameConflict ="overwrite");
            //taking a copy for result file
            cffile( fileField="file_name", destination="F:\Coldfushion\cfusion\wwwroot\Userdetails\files/result_"&fileUploadResult.clientFile, action="upload", result="fileUploadResult1",nameConflict ="overwrite");
            cfspreadsheet( action="read", src="../files/"&fileUploadResult.clientFile,query="importdata",excludeHeaderRow="true");

            //to list roles available in the system
            local.role_list = queryExecute("SELECT role FROM roles");
            local.arrayList = ValueArray(role_list,"role")
            //writeDump( var="#arrayList#" );
       
            local.i =2;
            local.emptystructs                = structnew();
            myArray=[];
            cfloop( query="importdata",startrow=2) {
                if ( (importdata.col_1 != "") || (importdata.col_2 != "") || (importdata.col_3 != "") || (importdata.col_4 != "") || (importdata.col_5 != "") || (importdata.col_6 != "") || (importdata.col_7 != "") ) {
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
                            myArray=importdata.col_1;
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
                    else
                    {
                        //emptystructs."empty"&i.firstname  = importdata.col_1;

                    }

                   
                    local.failList  = "";
                    
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
                        failList = listAppend(failList, "successfuly added");
                        writeDump( var="#failList#" );
                    }
                    if(statusflag == "update")
                    {
                        failList = listAppend(failList, "successfuly updated");
                    }

                    //writeDump( var="#failList#" );
                    spObj = spreadsheetread("F:\Coldfushion\cfusion\wwwroot\Userdetails\files/result_"&fileUploadResult.clientFile,"Sheet1");
                    myFormat       =StructNew();
                    myFormat.bold  ="true";
                    // populate Sheet 
                    SpreadsheetSetCellValue(spObj, "Result",1, 8);
                    SpreadsheetSetCellValue(spObj,failList,i, 8);
                    SpreadsheetFormatRow(spObj,myFormat,1);
                    SpreadsheetWrite(spObj,"F:\Coldfushion\cfusion\wwwroot\Userdetails\files/result_"&fileUploadResult.clientFile,"yes");
                }
                i++;
            }
            //location("../pages/home.cfm?name=result_"&fileUploadResult.clientFile, "false")
            //writeDump( var="#myArray#" );
        }
    }
    remote function downloadResult(required string file_name){
        spObj = spreadsheetread("F:\Coldfushion\cfusion\wwwroot\Userdetails\files/"&arguments.file_name,"Sheet1");
        cfheader( name="Content-Disposition", value="inline; filename=F:\Coldfushion\cfusion\wwwroot\Userdetails\files/"&arguments.file_name );
        cfcontent( type="application/vnd.msexcel",variable=SpreadSheetReadBinary(spObj));  
    }

    public function listUsers(){
        local.lists = queryExecute("SELECT* FROM users");
        //writeDump( var="#lists#" );   
    }
    
}