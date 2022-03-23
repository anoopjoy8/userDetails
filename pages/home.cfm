<cfset local.filename = ""/>
<cfif structKeyExists(url, 'name')>
      <cfset local.filename = url.name/>
</cfif>

<!--- get userdetails --->
<cfset user_list         = createObject("component","Userdetails/cfc/user")/>
<cfset users_list        = user_list.listUsers()/> 

<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="../assets/fontAwesome/font-awesome.min.css">
    <link rel="stylesheet" href="../assets/css/bootstrap.min.css">
    <link rel="stylesheet" href="../assets/developer.css">

    <title>User Details</title>
</head>
<header>
    <!-- As a heading -->
    <nav class="navbar navbar-light bg-light justify-content-center">
    <span class="navbar-brand mb-0 h2" id="fontsize">User Details</span>
    </nav>
</header>
<body>
    <cfoutput>
        <div class="container">
            <form action="../cfc/user.cfc?method=plainTemplate" enctype="multipart/form-data" method="post">
                <button type="submit" class="btn btn-primary l1 cm">Plain Template</button>
            </form>
            <form action="../cfc/user.cfc?method=downloadData" enctype="multipart/form-data" method="post">
                <button type="submit" class="btn btn-secondary l1">Template With  Data</button>
            </form>
            <form action="../cfc/user.cfc?method=upload" enctype="multipart/form-data" method="post">
                <button type="submit" class="btn btn-info r1 cm">Upload</button>
                <div class="mb-3 r1">
                    <input class="form-control" name="file_name" type="file" id="formFile">
                </div>
            </form>

            #users_list#
            
            <!-- Modal -->
            <div class="modal fade" id="exampleModalCenter" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
                <div class="modal-dialog modal-dialog-centered" role="document">
                    <div class="modal-content">
            
                    <div class="modal-body">
                        File Uploaded Successfully!!!
                        <br>
                        <form action="../cfc/user.cfc?method=downloadResult" method="post">
                            <input type="hidden" id="filename" name="file_name" value="#local.filename#"/>
                            <button type="submitt" class="btn btn-secondary">Click to download result file</button>
                        </form>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
                    </div>
                    </div>
                </div>
            </div>
        </div>
    </cfoutput>
</body>
<script src="../assets/js/jquery.min.js"></script>
<script src="../assets/js/bootstrap.bundle.min.js"></script>
<script>
$(document).ready(function(){
    var fname = $('#filename').val();
    if(fname !="")
    {
        $("#exampleModalCenter").modal();
    }
});
</script>