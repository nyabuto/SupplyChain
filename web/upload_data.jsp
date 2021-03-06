<%-- 
    Document   : upload_data
    Created on : May 4, 2018, 12:52:41 PM
    Author     : GNyabuto
--%>

<%@page import="org.json.simple.JSONArray"%>
<%@page import="org.json.simple.JSONObject"%>
<%@page contentType="text/html" pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <title>Upload Excel Data</title>
    <meta name="description" content="Sufee Admin - HTML5 Admin Template">
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <link rel="apple-touch-icon" href="images/logo.png">
    <link rel="shortcut icon" href="images/logo.png">

    <link rel="stylesheet" href="assets/css/normalize.css">
    <link rel="stylesheet" href="assets/css/bootstrap.min.css">
    <link rel="stylesheet" href="assets/css/font-awesome.min.css">
    <link rel="stylesheet" href="assets/css/themify-icons.css">
    <link rel="stylesheet" href="assets/css/flag-icon.min.css">
    <link rel="stylesheet" href="assets/css/cs-skin-elastic.css">
    <!-- <link rel="stylesheet" href="assets/css/bootstrap-select.less"> -->
    <link rel="stylesheet" href="assets/scss/style.css">

    <link href='https://fonts.googleapis.com/css?family=Open+Sans:400,600,700,800' rel='stylesheet' type='text/css'>

    <!-- <script type="text/javascript" src="https://cdn.jsdelivr.net/html5shiv/3.7.3/html5shiv.min.js"></script> -->

</head>
<body>
        <!-- Left Panel -->
    <%@include file="sidebar.jsp" %>

    <!-- Left Panel -->

    <!-- Right Panel -->

    <div id="right-panel" class="right-panel">

        <!-- Header-->
       <%@include file="header.jsp" %>
        <!-- Header-->

        <div class="content mt-3">
            <div class="animated fadeIn">


                <div class="row">

                  <div class="col-lg-12">
                    <div class="card">
                      <div class="card-header">
                        <strong>Data Upload Module</strong>
                      </div>
                      
                        <form action="upload_excels" method="post" enctype="multipart/form-data" class="form-horizontal">
                         <div class="card-body card-block"> 
                          <div class="row form-group">
                            <!--<div class="col col-md-3"><label for="file-multiple-input" class=" form-control-label">Select Excel files to upload</label></div>-->
                            <!--<div class="col-12 col-md-6"><input type="file" id="file" name="file" multiple="" class="form-control-file"></div>-->
                          </div>
                    
                      </div>
                        <div class="card-footer" style="text-align: right;">
                        
<!--                        <button type="reset" class="btn btn-danger btn-sm">
                          <i class="fa fa-ban"></i> Cancel
                        </button>-->
                        
                            <div type="submit" class="btn btn-danger btn-sm">
                          <i class="fa fa-dot-circle-o"></i> You are not allowed to upload data
                        </div>
                      </div>
                            </form>
                    </div>
                  </div>
                    <div style="margin-left: 15%;">
                    <div id="uploaded" class="thead-light" style="text-align: center;">
                    <!--uploaded-->
                    </div>
                    <br>
                    <div id="errors" class="thead-light">
                    </div >
                    </div>
                </div>


            </div><!-- .animated -->
        </div><!-- .content -->


    </div><!-- /#right-panel -->

    <!-- Right Panel -->


    <script src="assets/js/vendor/jquery-1.11.3.min.js"></script>
    <script src="assets/js/popper.min.js"></script>
    <script src="assets/js/plugins.js"></script>
    <script src="assets/js/main.js"></script>


    <%if(session.getAttribute("upload_errors")!=null){
        int uploaded=0; 
        JSONObject finalobj = new JSONObject();
        JSONArray jarray = new JSONArray();
        finalobj = (JSONObject)session.getAttribute("upload_errors");
//        out.println(finalobj);
        jarray = (JSONArray)finalobj.get("data");
        uploaded = (Integer)finalobj.get("uploaded");
        out.println(uploaded);
//        out.println(jarray);
String files_uploaded = "<b style=\"color:red;\">"+uploaded+"</b> <b style=\"color:black;\">Files added/updated successfully.</b>";
        String output="<div style=\"font-weight:900; text-align:center;\">The following files were not uploaded:</div><br>"
                + "<table  class=\"table\"><thead class=\"thead-dark\">"
                + "<tr><th>No.</th><th>File Name</th><th>Error Description</th></tr></thead><tbody>";
        int counter=0;
            for(int i=0;i<jarray.size();i++) {
                counter++;
                JSONObject obj = (JSONObject)jarray.get(i);
//                out.println(obj);
               String description = obj.get("description").toString();
               String file_name =  obj.get("file_name").toString();
               
              output+="<tr><td>"+(i+1)+"</td><td>"+file_name+"</td><td>"+description+"</td></tr>"; 
            }
        output+="</tbody></table>";
    %>
     <script type="text/javascript">
           jQuery(document).ready(function() {
               <%if(counter>0){%>
            jQuery("#errors").html('<%=output%>');
            <%}if(uploaded>0 ){%>
            jQuery("#uploaded").html('<%=files_uploaded%>');
            <%}%>
            });
    </script>
    <% 
        session.removeAttribute("upload_errors");
        }
    %>
    
</body>
</html>
