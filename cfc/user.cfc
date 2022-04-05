component {
    remote function upload(required string file_name){
        try {
            if ( len(trim(arguments.file_name)) ) {
                cffile( fileField="file_name", destination="F:\Coldfushion\cfusion\wwwroot\Userdetails\files", action="upload", result="fileUploadResult",nameConflict ="overwrite");
                //taking a copy of uploaded file for making result file
                cffile( fileField="file_name", destination="F:\Coldfushion\cfusion\wwwroot\Userdetails\files/result_"&fileUploadResult.clientFile, action="upload", result="fileUploadResult1",nameConflict ="overwrite");
                cfspreadsheet( action="read", src="../files/"&fileUploadResult.clientFile,query="importdata",excludeHeaderRow="true");

                //to list roles available in the system
                local.role_list = queryExecute("SELECT role FROM roles");
                local.arrayList = ValueArray(role_list,"role");
        
                local.i = 2;
                local.statusflag  = "";
                local.em="";
                local.em_chk = 0;
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
                                    failList = listAppend(failList, itm&" is not a valid role");
                                }
                            }

                            if(flag == 0)
                            { 
                                local.statusflag  = "";
                                email_chk = queryExecute(
                                    "SELECT id,firstname,email FROM users WHERE email = :email_id;", 
                                    {
                                        email_id   : { cfsqltype: "cf_sql_varchar", value: importdata.col_4}
                                    }
                                );
                                if(email_chk.RecordCount > 0){ 
                                    local.lfind= listFind(em,email_chk.id, ",") ;
                                    if(lfind == 0){
                                        statusflag  = "update";
                                        em_chk = 0;
                                        queryExecute("
                                                UPDATE users
                                                SET firstname = :firstname, lastname= :lastname, address= :address, email= :email, phone= :phone, dob= :dob
                                                WHERE id = :ID;",
                                            {
                                                firstname: { cfsqltype: "cf_sql_varchar", value: importdata.col_1 },
                                                lastname: { cfsqltype: "cf_sql_varchar", value: importdata.col_2 },
                                                address: { cfsqltype: "cf_sql_varchar", value: importdata.col_3 },
                                                email: { cfsqltype: "cf_sql_varchar", value: importdata.col_4 },
                                                phone: { cfsqltype: "cf_sql_varchar", value: importdata.col_5 },
                                                dob: { cfsqltype: "cf_sql_varchar", value: DateFormat(importdata.col_6,"yyy-mm-dd") },
                                                ID: { cfsqltype: "cf_sql_varchar", value: email_chk.id}

                                            },
                                            {result = "rslt"}
                                        );

                                        //delete roles table
                                 /*        queryExecute("
                                               DELETE FROM user_roles WHERE user_id= :ID;",
                                            {
                                                ID: { cfsqltype: "cf_sql_varchar", value: email_chk.id}
                                            }
                                        ); */
                                        local.r_lists =[];

                                        for( role_name in rolearray1 ) {

                                            rq = queryExecute(
                                                "SELECT id FROM roles 
                                                WHERE role = :rolename;", 
                                                {
                                                    rolename: { cfsqltype: "cf_sql_varchar", value: role_name}
                                                }
                                            );
                                            
                                            arrayAppend(r_lists,rq.id);
                                           
                                            
                                            queryExecute("
                                                INSERT INTO user_roles(user_id,role_id) 
                                                SELECT c.id,roles.id
                                                FROM  roles
                                                INNER JOIN users c ON c.id = :ID
                                                WHERE(roles.role = :rolename)
                                                AND (:ID NOT IN (SELECT user_id FROM user_roles where roles.id = user_roles.role_id ))
                                                ",

                                                {
                                                    rolename: { cfsqltype: "cf_sql_varchar", value: role_name},
                                                    ID: { cfsqltype: "cf_sql_varchar", value: email_chk.id}
                                                }                           
                                            );     
                                     
                                        }
                                        
                                    
                                    //rArray =listToArray(r_lists," ",true,true);
                               
                                        aa =r_lists.toList();
                                        df =queryExecute(
                                                "DELETE FROM user_roles 
                                                WHERE user_id = #email_chk.id# AND
                                                role_id NOT IN (#aa#);"
                                            );
                                         
                        
                                       
                     
                        //writeDump(df);

                                    }
                                    else
                                    {
                                        em_chk = 1;
                                    }
                                    em = listAppend(em,email_chk.id);
                                      
                                }else{
                                    statusflag  = "add";
                                    em_chk = 0;
                                    queryExecute("
                                        insert into users(firstname,lastname,address,email,phone,dob)
                                        values( :firstname, :lastname, :address, :email, :phone, :dob )",
                                        {
                                            firstname: { cfsqltype: "cf_sql_varchar", value: importdata.col_1 },
                                            lastname: { cfsqltype: "cf_sql_varchar", value: importdata.col_2 },
                                            address: { cfsqltype: "cf_sql_varchar", value: importdata.col_3 },
                                            email: { cfsqltype: "cf_sql_varchar", value: importdata.col_4 },
                                            phone: { cfsqltype: "cf_sql_varchar", value: importdata.col_5 },
                                            dob: { cfsqltype: "cf_sql_varchar", value: DateFormat(importdata.col_6,"yyy-mm-dd") }
                                        },
                                        {result =  "sResult"}                           
                                    );

                                    
                                    for( role_name in rolearray1 ) {
                                        role_id_res = queryExecute(
                                            "SELECT id FROM roles WHERE role = :rolename;", 
                                            {
                                                    rolename   : { cfsqltype: "cf_sql_varchar", value: role_name}
                                            }
                                        );
                                        queryExecute("
                                            insert into user_roles(user_id,role_id)
                                            values( :UserId, :roleId )",
                                            {
                                                UserId: { cfsqltype: "cf_sql_integer", value: sResult.GENERATEDKEY},
                                                roleId: { cfsqltype: "cf_sql_integer", value: role_id_res.id }
                                            }                           
                                        );
                                     
                                    }
                                    em = listAppend(em,sResult.GENERATEDKEY); 
                                }
                                
                            }
                        }
                       
                        local.failList="";

                        for( itm in rolearray1 ) {
                            if(arrayFind(arrayList,itm) == 0)
                            {
                                    failList = listAppend(failList, itm&" is not a valid role");
                            }
                        }



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
                        if(em_chk == 1){
                            failList = listAppend(failList, "email already exist");
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
                cffile(action="delete",file="F:\Coldfushion\cfusion\wwwroot\Userdetails\files/"&fileUploadResult.clientFile);
                location("../pages/home.cfm?name=result_"&fileUploadResult.clientFile, "false")
            }
        }
        catch(Exception e){
			return 'error';
		}
    }
    remote function downloadResult(required string file_name){
        try {
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
            cffile(action="delete",file="F:\Coldfushion\cfusion\wwwroot\Userdetails\files/"&arguments.file_name);

            cfheader( name="Content-Disposition", value="inline; filename=result_file.xlsx");
            cfcontent( type="application/vnd.msexcel",variable=SpreadSheetReadBinary(spObj));
 
        }
        catch(Exception e){
			return 'error';
		} 
    }

    public function listUsers(){
        try {
            local.lists = queryExecute("SELECT firstname,lastname,address,email,phone,dob,GROUP_CONCAT(roles.role)as role_lists FROM users
                                   INNER JOIN user_roles ON users.id = user_roles.user_id 
                                   INNER JOIN roles ON user_roles.role_id = roles.id
                                   GROUP BY users.id
                            "); 
            savecontent variable="usersList" {
                writeOutput("<table class=""table"">
                <thead>
                    <tr>
                        <th scope=""col"">First Name</th>
                        <th scope=""col"">Last Name</th>
                        <th scope=""col"">Email</th>
                        <th scope=""col"">Address</th>
                        <th scope=""col"">Phone</th>
                        <th scope=""col"">DOB</th>
                        <th scope=""col"">Role</th>
                    </tr>
                </thead>
                <tbody>");
                for ( x in lists ) {
                    cfoutput(  ) {

                        writeOutput("<tr>
                            <td>#x.firstname#</td>
                            <td>#x.lastname#</td>
                            <td>#x.email#</td>
                            <td>#x.address#</td>
                            <td>#x.phone#</td>
                            <td>#DateFormat(x.dob,"dd-mm-yyy")#</td>
                            <td>#x.role_lists#</td>");
                        }
                }
                writeOutput("</tbody>
                </table>");
            }
            return usersList;
        } 
        catch(Exception e){
			return 'error';
		}
    }

    remote function downloadData(){
        try {
            local.lists          = queryExecute("SELECT firstname,lastname,address,email,phone,dob,GROUP_CONCAT(roles.role)as role_lists FROM users
                                   INNER JOIN user_roles ON users.id = user_roles.user_id 
                                   INNER JOIN roles ON user_roles.role_id = roles.id
                                   GROUP BY users.id
                                   ");                      
            local.spObj          = spreadsheetNew("Sheet1",true);
            spreadsheetAddRow(spObj, "First Name,Last name,Address,Email,Phone,DOB,Role");
            local.myFormat       = StructNew();
            local.myFormat.bold  ="true";
            spreadsheetAddrows(spObj,lists,2,1,true,[""],false);
            SpreadsheetFormatRow(spObj,myFormat,1);
            //SpreadsheetWrite(spObj,"F:\Coldfushion\cfusion\wwwroot\Userdetails\files/result_withdatafile.xlsx","yes");
            cfheader( name="Content-Disposition", value="inline; filename=result_withdatafile.xlsx");
            cfcontent( type="application/vnd.msexcel",variable=SpreadSheetReadBinary(spObj));
        }
        catch(Exception e){
			return 'error';
		}   
    }
    remote function plainTemplate(){ 
        try {
            local.spObj         = spreadsheetNew("Sheet1",true);
            local.myFormat      = StructNew();
            local.myFormat.bold = "true";
            spreadsheetAddRow(spObj, "First Name,Last name,Address,Email,Phone,DOB,Role");
            SpreadsheetFormatRow(spObj,myFormat,1);
            cfheader( name="Content-Disposition", value="inline; filename=plain_templatefile.xlsx");
            cfcontent( type="application/vnd.msexcel",variable=SpreadSheetReadBinary(spObj));
        }
        catch(Exception e){
			return 'error';
		}  
    }      
}
