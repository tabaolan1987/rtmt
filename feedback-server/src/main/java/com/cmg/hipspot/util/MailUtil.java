package com.cmg.hipspot.util;

import com.cmg.hipspot.data.ContactModel;
import com.cmg.hipspot.data.jdo.FeedbackModel;
import org.apache.commons.io.IOUtils;
import org.apache.log4j.Logger;

import java.io.*;

/**
 * Created by lantb on 2014-04-21.
 */
public class MailUtil {
    private static final Logger logger = Logger.getLogger(MailUtil.class
            .getName());
    public String getBody (FeedbackModel model){
        StringBuffer temp = new StringBuffer();
        temp.append("<p style=\"color:#666; font-family: arial; font-size:10pt;\">Dear all,</p>\n");
        temp.append("<p style=\"color:#666; font-family: arial; font-size:10pt;\">There is  a new feedback for RTMT Tool from client :  </p>\n");
        temp.append("<p style=\"color:#666; font-family: arial; font-size:10pt;\">Email : "+model.getEmail()+" </p>\n");
        temp.append("<p style=\"color:#666; font-family: arial; font-size:10pt;\">Description : "+model.getDescription()+" </p>\n");
        temp.append("<p style=\"color:#666; font-family: arial; font-size:10pt;\">Version : "+model.getVersion()+" </p>\n");
        temp.append("<p style=\"color:#666; font-family: arial; font-size:10pt;\">OS Information : "+model.getOsInformation()+"</p>\n");
        temp.append("<p style=\"color:#666; font-family: arial; font-size:10pt;\">Step to get error : "+model.getStepError()+" </p>\n");
        temp.append("<p style=\"color:#666; font-family: arial; font-size:10pt;\">Kind regards,<br />\n" +
                "FeedBack Support</p>");
        String body = temp.toString();
        logger.info("body : " + body);
        return body;
    }

    public static String getBodyContactMail(ContactModel model) {
        StringBuffer temp = new StringBuffer();
        temp.append("<p style=\"color:#666; font-family: arial; font-size:10pt;\">Dear all,</p>\n");
        temp.append("<p style=\"color:#666; font-family: arial; font-size:10pt;\">There is a contact information from C-MG website :  </p>\n");
        temp.append("<p style=\"color:#666; font-family: arial; font-size:10pt;\">Email : "+model.getEmail()+" </p>\n");
        temp.append("<p style=\"color:#666; font-family: arial; font-size:10pt;\">First name : "+model.getFirstName()+" </p>\n");
        temp.append("<p style=\"color:#666; font-family: arial; font-size:10pt;\">Last name : "+model.getLastName()+" </p>\n");
        temp.append("<p style=\"color:#666; font-family: arial; font-size:10pt;\">Message : "+model.getMessage() +"</p>\n");
        temp.append("<p style=\"color:#666; font-family: arial; font-size:10pt;\">Happy to keep update with C-MG developments : "+ (model.isHappy() ? "Yes" : "No") +" </p>\n");
        temp.append("<p style=\"color:#666; font-family: arial; font-size:10pt;\">Interested in becoming a C-MG member : "+ (model.isMember() ? "Yes" : "No") +" </p>\n");

        temp.append("<p style=\"color:#666; font-family: arial; font-size:10pt;\">Kind regards,<br />\n" +
                "Contact Support</p>");
        String body = temp.toString();
        logger.info("body : " + body);
        return body;
    }

}
