<%-- 
    Document   : raw_data
    Created on : May 8, 2018, 12:11:56 PM
    Author     : GNyabuto
--%>

<%@page contentType="text/html" pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <title>Raw Data</title>
    <meta name="description" content="Sufee Admin - HTML5 Admin Template">
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <link rel="apple-touch-icon" href="apple-icon.png">
    <link rel="shortcut icon" href="images/logo.png">

    <link rel="stylesheet" href="assets/css/normalize.css">
    <link rel="stylesheet" href="assets/css/bootstrap.min.css">
    <link rel="stylesheet" href="assets/css/font-awesome.min.css">
    <link rel="stylesheet" href="assets/css/themify-icons.css">
    <link rel="stylesheet" href="assets/css/flag-icon.min.css">
    <link rel="stylesheet" href="assets/css/cs-skin-elastic.css">
    <!-- <link rel="stylesheet" href="assets/css/bootstrap-select.less"> -->
    <link rel="stylesheet" href="assets/scss/style.css">
<link rel="stylesheet" href="assets/css/lib/chosen/chosen.min.css">
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
                        <strong>Raw Data Generation Module</strong>
                      </div>
                      <form action="raw_data">
                          <form action="upload_excels" method="post" enctype="multipart/form-data" class="form-horizontal">
                         <div class="card-body card-block"> 
                          <div class="row form-group">
                            <div class="col col-md-3"><label for="file-multiple-input" class=" form-control-label">Select County</label></div>
                            <div class="col-12 col-md-6">
                                
                                <select id="county" name="county"  data-placeholder="Choose counties..." multiple class="standardSelect form-control">
                                    <option value =""> Choose Counties</option>
                                    </select>
                            </div>
                          </div>
                          <div class="row form-group">
                            <div class="col col-md-3"><label for="file-multiple-input" class=" form-control-label">Select Sub County</label></div>
                            <div class="col-12 col-md-6">
                                
                                <select id="sub_county" name="sub_county"  data-placeholder="Choose sub counties..." multiple class="standardSelect form-control">
                                    <option value =""> Choose Sub Counties</option>
                                    </select>
                            </div>
                          </div>
                          <div class="row form-group">
                            <div class="col col-md-3"><label for="file-multiple-input" class=" form-control-label">Select Health Facility</label></div>
                            <div class="col-12 col-md-6">
                                
                                <select id="mfl_code" name="mfl_code" data-placeholder="Choose health facilities..." multiple class="standardSelect form-control" style="height: 200px;">
                                    <option value =""> Choose Health Facilities</option>
                                    </select>
                            </div>
                          </div>
                    
                      </div>
                        <div class="card-footer" style="text-align: right;">
                        
                            <button type="submit" class="btn btn-primary">
                          <i class="fa fa-dot-circle-o"></i> Generate Report
                        </button>
                      </div>
                            </form> 
                          </form>
                    </div>
                  </div>


                </div>


            </div><!-- .animated -->
        </div><!-- .content -->


    </div><!-- /#right-panel -->

    <!-- Right Panel -->


    <script src="assets/js/vendor/jquery-2.1.4.min.js"></script>
    <script src="assets/js/popper.min.js"></script>
    <script src="assets/js/plugins.js"></script>
    <script src="assets/js/main.js"></script>
    <script src="assets/js/lib/chosen/chosen.jquery.min.js"></script>

    <script>
        jQuery(document).ready(function() {
            jQuery(".standardSelect").chosen({
                disable_search_threshold: 10,
                no_results_text: "Oops, nothing found!",
                width: "100%"
            });
        });
    </script>
<script>
   jQuery(document).ready(function() {
      load_counties();
      load_sub_counties();
      load_facilities();
      
      jQuery("#county").change(function(){
      load_sub_counties();    
      });
      jQuery("#sub_county").change(function(){
      load_facilities();    
      });
   }); 
 function load_counties(){
       jQuery.ajax({
        url:'load_county',
        type:"post",
        dataType:"json",
        success:function(raw_data){
         var county,output="";
         var data = raw_data.data;
          county="";
             for (var i=0; i<data.length;i++){
            if( data[i].county!=null){county = data[i].county;}
            output+="<option value='"+county+"'>"+county+"</option>"; 
         }
         // ouput
         jQuery("#county").html(output);
         jQuery("#county").chosen("destroy");
         jQuery("#county").chosen({
                disable_search_threshold: 10,
                no_results_text: "Oops, no county found!",
                width: "100%"
            });
        }
  });   
        
 }   
 function load_sub_counties(){
     
     var county = jQuery("#county").val();
     if(county==null){county="";}
     
     var form_data = {"county":county};
                var url = "load_subcounty";
                   jQuery.post(url,form_data , function(raw_data) {
                        var sub_county,output="";
                        var data = raw_data.data;
                         sub_county="";
                            for (var i=0; i<data.length;i++){
                           if( data[i].sub_county!=null){sub_county = data[i].sub_county;}
                           output+="<option value='"+sub_county+"'>"+sub_county+"</option>"; 
                        }
                        // ouput
                        jQuery("#sub_county").html(output);
                        jQuery("#sub_county").chosen("destroy");
                        jQuery("#sub_county").chosen({
                               disable_search_threshold: 10,
                               no_results_text: "Oops, no sub county found!",
                               width: "100%"
                           });

               });
    
        
 }   
 function load_facilities(){
     alert("called");
     var county = jQuery("#county").val();
     var sub_county = jQuery("#sub_county").val();
     if(county==null){county=""}
     if(sub_county==null){sub_county=""}
      
     var form_data = {"county":county,"sub_county":sub_county};
                var url = "load_facilities";
                   jQuery.post(url,form_data , function(raw_data) {
                        var mfl_code,facility_name,output="";
                        var data = raw_data.data;
                         sub_county="";
                            for (var i=0; i<data.length;i++){
                                mfl_code = facility_name="";
                           if( data[i].mfl_code!=null){mfl_code = data[i].mfl_code;}
                           if( data[i].facility_name!=null){facility_name = data[i].facility_name;}
                           output+="<option value='"+mfl_code+"'>"+facility_name+"</option>"; 
                        }
                        // ouput
                        jQuery("#mfl_code").html(output);
                        jQuery("#mfl_code").chosen("destroy");
                        jQuery("#mfl_code").chosen({
                               disable_search_threshold: 10,
                               no_results_text: "Oops, no sub county found!",
                               width: "100%"
                           });     
                       
                       
                   });
     
 }   
    
</script>
</body>
</html>
