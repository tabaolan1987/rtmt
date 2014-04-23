package com.cmg.hipspot.service;

import com.cmg.hipspot.data.jdo.FeedbackModel;
import com.cmg.hipspot.util.MailUtil;
import org.apache.log4j.Logger;

import java.io.File;
import java.util.Date;
import java.util.Properties;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

/**
 * Created by lantb on 2014-04-21.
 */
public class MailServices extends Thread{
    private static final Logger logger = Logger.getLogger(MailServices.class
            .getName());

    private String error;
    private FeedbackModel model;
    public MailServices(FeedbackModel model){
        this.model = model;
    }

    public MailServices(String error){
        this.error = error;
    }

    public boolean sendMailError(String error){
        MailUtil util = new MailUtil();
        try {
            String body = "This is the error description : " + error;
            String subject = "Feedback server has been error";
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
            message.setFrom(new InternetAddress("feedback"));
            message.setRecipients(Message.RecipientType.TO, "lan.ta@c-mg.com");
            message.setSubject(subject);

            Multipart mp = new MimeMultipart("related");
            MimeBodyPart mbp1 = new MimeBodyPart();
            mbp1.setContent(new String(body.toString().getBytes(), "iso-8859-1"), "text/html; charset=\"iso-8859-1\"");
            mp.addBodyPart(mbp1);
            message.setContent(mp);
            message.setSentDate(new Date());
            message.saveChanges();
            Transport.send(message);
            return true;
        }catch (Exception e){
            e.printStackTrace();
            logger.error(e.getMessage());
        }
        return false;
    }
    public boolean sendMail(FeedbackModel model){
        MailUtil util = new MailUtil();
        try {
            logger.info("start sending mail");
            String body = util.getBody(model);
            String subject = "New FeedBack RTMT";
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
            message.setFrom(new InternetAddress("lan.ta@c-mg.com"));
            message.setRecipients(Message.RecipientType.TO, "lan.ta@c-mg.com");
            message.setSubject(subject);

            Multipart mp = new MimeMultipart("related");
            MimeBodyPart mbp1 = new MimeBodyPart();
            mbp1.setContent(new String(body.toString().getBytes(), "iso-8859-1"), "text/html; charset=\"iso-8859-1\"");
            mp.addBodyPart(mbp1);

            if(model.getPictureError()!=null){
                String test = model.getPictureError().substring(0,model.getPictureError().length()-1);
                logger.info("all file before :" + test );
                String[] allFiles = test.split("\\|");
                for(String temp : allFiles){
                    logger.info("temp file : " + temp);
                    File file = new File(temp);
                    if(file.exists()){
                        logger.info("File existed : " + file.getAbsolutePath());
                        MimeBodyPart mbp2 = new MimeBodyPart();
                        FileDataSource fds = new FileDataSource(file);
                        mbp2.setDataHandler(new DataHandler(fds));
                        mbp2.setFileName(fds.getName());
                        mp.addBodyPart(mbp2);
                    }
                }
            }

            if(model.getTestData()!=null){
                String test = model.getTestData().substring(0,model.getTestData().length()-1);
                logger.info("all test before :" + test );
                String[] allFiles = test.split("\\|");
                for(String temp : allFiles){
                    File file = new File(temp);
                    if(file.exists()){
                        MimeBodyPart mbp3 = new MimeBodyPart();
                        FileDataSource fds = new FileDataSource(file);
                        mbp3.setDataHandler(new DataHandler(fds));
                        mbp3.setFileName(fds.getName());
                        mp.addBodyPart(mbp3);
                    }
                }
            }

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
            if(model!=null){
                logger.info("start thread");
                sendMail(model);
            }
            if(error!=null){
                sendMailError(error);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

    }


}
