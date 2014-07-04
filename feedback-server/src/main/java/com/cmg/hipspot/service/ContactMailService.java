package com.cmg.hipspot.service;

import com.cmg.hipspot.util.MailUtil;
import org.apache.log4j.Logger;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.io.File;
import java.util.Date;
import java.util.Properties;

/**
 * Created by Hai Lu on 05/06/2014.
 */
public class ContactMailService extends Thread {
    private static final Logger logger = Logger.getLogger(ContactMailService.class
            .getName());

    private final String body;
    public ContactMailService(String body) {
        this.body = body;
    }
    public boolean sendMail(String body) {
        MailUtil util = new MailUtil();
        try {
            logger.info("start sending mail");
            String subject = "New Contact Information";
            Properties mailProps = System.getProperties();
            mailProps = System.getProperties();
            mailProps.put("mail.smtp.host", "smtp.gmail.com");
            mailProps.put("mail.smtp.socketFactory.port", "465");
            mailProps.put("mail.smtp.socketFactory.class",
                    "javax.net.ssl.SSLSocketFactory");
            mailProps.put("mail.smtp.auth", "true");
            mailProps.put("mail.smtp.port", "465");
            mailProps.put("mail.smtp.auth", "true");
            Authenticator pa = new Authenticator() {
                @Override
                protected PasswordAuthentication getPasswordAuthentication() {
                    return new PasswordAuthentication("feedbackcmg@c-mg.com","W3lcom3123");
                }
            };

            Session session = Session.getInstance(mailProps, pa);

            MimeMessage message = new MimeMessage(session);
            message.setHeader("Content-Type", "text/html");
            message.setFrom(new InternetAddress("feedbackcmg@c-mg.com"));
            message.setRecipients(Message.RecipientType.TO, "hai.lu@c-mg.com");
            message.setSubject(subject);

            Multipart mp = new MimeMultipart("related");
            MimeBodyPart mbp1 = new MimeBodyPart();
            mbp1.setContent(new String(body.toString().getBytes(), "iso-8859-1"), "text/html; charset=\"iso-8859-1\"");
            mp.addBodyPart(mbp1);
            message.setContent(mp);
            message.setSentDate(new Date());
            message.saveChanges();
            Transport.send(message);
            logger.info("end sending mail");
            return true;
        }catch (Exception e){
            e.printStackTrace();
            logger.error(e.getMessage());
        }
        return false;
    }

    @Override
    public void run() {
        try {
            sendMail(body);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
