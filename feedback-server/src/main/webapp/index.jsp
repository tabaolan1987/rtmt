<%@ page import="com.cmg.hipspot.util.StringUtil" %>
<!DOCTYPE html>
<html lang="eng">
<head>
    <title>Feedback RTMT</title>
    <!-- Bootstrap CSS -->
    <link href="./css/bootstrap.min.css" rel="stylesheet" media="screen">
    <style>
        .btn-file {
            position: relative;
            overflow: hidden;
        }

        .btn-file input[type=file] {
            position: absolute;
            top: 0;
            right: 0;
            min-width: 100%;
            min-height: 100%;
            font-size: 999px;
            text-align: right;
            filter: alpha(opacity=0);
            opacity: 0;
            outline: none;
            background: white;
            cursor: inherit;
            display: block;
        }
    </style>
</head>
<body>
<%
    String result = (String) StringUtil.isNull(request.getAttribute("result"),"");
%>
<br/>
<!-- Contacts -->
<div id="contacts">
    <div class="row">
        <!-- Alignment -->
        <div class="col-sm-offset-3 col-sm-6">
            <!-- Form itself -->
            <%
                if(result!=""){
            %>
            <div class="alert alert-success">
                <button type="button" class="close" data-dismiss="alert">&times;</button>
                <%
                    if(result.equalsIgnoreCase("success")){
                %>
                    <h4>Special thanks to your feedback!</h4>
                        <h5>We will improve this as soon as possible</h5>
                <%}else{%>
                    <h4>Oops! There is error occur in server! </h4>
                       <h5> Don't worry about this, please try again in minute.</h5>
                <%}%>
            </div>
            <%}%>
            <form method="POST" name="sentMessage" class="well" id="contactForm"
                  enctype="multipart/form-data" action="/feedback/RtmtHandler">
                <legend>FeedBack</legend>

                <div class="control-group">
                    <div class="controls">
                        <input name="email" type="email" class="form-control" placeholder="Email (*)"
                               id="email" required
                               data-validation-required-message="Please enter your email"/>
                    </div>
                </div>


                <div class="control-group">
                    <div class="controls">
                        <textarea rows="5" cols="100" class="form-control"
                                  placeholder="ERROR DESCRIPTION (*) " name="description" id="description" required
                                  data-validation-required-message="Please enter your message" minlength="5"
                                  data-validation-minlength-message="Min 5 characters"
                                  maxlength="999" style="resize:none"></textarea>
                    </div>
                </div>
                <div class="control-group">
                    <div class="controls">
                        <input type="file" name="screenshot" id="screenshot" multiple accept='image/*' title="Attach Screen shot">
                    </div>
                </div>
                <div class="control-group">
                    <div class="controls">
                        <input type="text" class="form-control"
                               placeholder="Version of Tool" id="version" name="version"/>

                        <p class="help-block"></p>
                    </div>
                </div>
                <div class="control-group">
                    <div class="controls">
                        <input type="text" class="form-control"
                               placeholder="Computer type/OS information/installed software" name="os_information" id="os_information"/>

                        <p class="help-block"></p>
                    </div>
                </div>
                <div class="control-group">
                    <div class="controls">
                        <textarea rows="5" cols="100" class="form-control"
                                  placeholder="STEP TO GET ERROR " name="stepERROR" id="stepERROR"
                                  maxlength="999" style="resize:none"></textarea>
                    </div>
                </div>
                <div class="control-group">
                    <div class="controls">
                       <input type="file" id="testData" name="testData" title="Attach Test Data">
                    </div>
                </div>
                <button type="submit" class="btn btn-primary pull-right">Send</button>
                <br/>
            </form>
        </div>
    </div>
</div>


<!-- JS FILES -->
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.7.2/jquery.min.js"></script>
<script src="./js/bootstrap.min.js"></script>
<script src="./js/jqBootstrapValidation.js"></script>
<script src="./js/bootstrap.file-input.js"></script>
<script src="./js/contact_me.js"></script>

</body>
</html>