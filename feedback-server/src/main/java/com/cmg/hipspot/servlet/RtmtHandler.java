package com.cmg.hipspot.servlet;

import com.cmg.hipspot.data.dao.impl.FeedbackModelDAO;
import com.cmg.hipspot.data.jdo.FeedbackModel;
import com.cmg.hipspot.service.FileServices;
import com.cmg.hipspot.service.MailServices;
import com.cmg.hipspot.util.StringUtil;
import org.apache.commons.fileupload.FileItemIterator;
import org.apache.commons.fileupload.FileItemStream;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.commons.fileupload.util.Streams;
import org.apache.log4j.Logger;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

/**
 * Created by lantb on 2014-04-22.
 */
public class RtmtHandler extends HttpServlet {
    private static final Logger logger = Logger.getLogger(RtmtHandler.class
            .getName());
    private static String EMAIL_PARA = "email";
    private static String SCREEN_SHOT_PARA = "screenshot";
    private static String DESCRIPTION_PARA = "description";
    private static String VERSION_PARA = "version";
    private static String OS_INFORMATION_PARA = "os_information";
    private static String STEP_TO_ERROR_PARA = "stepERROR";
    private static String TEST_DATA = "testData";
    private static String FOLDER_SCREEN_SHOT = "screenshot";
    private static String FOLDER_TEST_DATA = "testdata";
    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        try {
            //response.setContentType("text/html;charset=UTF-8");
            FileServices fileServices = new FileServices();
            FeedbackModelDAO dao = new FeedbackModelDAO();
            FeedbackModel model = new FeedbackModel();
            logger.info("coming serlet");
            if(ServletFileUpload.isMultipartContent(request)) {
                logger.info("multipart");
                ArrayList<String> screenshotFiles = new ArrayList<String>();
                ArrayList<String> testDataFiles = new ArrayList<String>();
                ServletFileUpload upload = new ServletFileUpload();
                FileItemIterator iter = upload.getItemIterator(request);
                while (iter.hasNext()) {
                    FileItemStream item = iter.next();
                    InputStream stream = item.openStream();
                    if(!item.isFormField()){
                        String field = item.getFieldName();
                        logger.info("field name " + field);
                        if(field.equalsIgnoreCase(TEST_DATA)){
                            String pathDataFolder = fileServices.getPath(FOLDER_TEST_DATA);
                            logger.info("path Data Folder : " + pathDataFolder);
                            if(item.getName()!=null && item.getName()!=""){
                                String testData = fileServices.saveFile(stream,pathDataFolder,item.getName());
                                logger.info("test data : " + testData);
                                testDataFiles.add(testData);
                                //model.setTestData(testData);
                            }
                        }else if(field.equalsIgnoreCase(SCREEN_SHOT_PARA)){
                            String pathScreenshotFolder = fileServices.getPath(FOLDER_SCREEN_SHOT);
                            logger.info("path Screen shot : " + pathScreenshotFolder);
                            if(item.getName()!=null && item.getName()!=""){
                                String pictureError = fileServices.saveFile(stream,pathScreenshotFolder,item.getName());
                                logger.info("picture Error : " + pictureError);
                                screenshotFiles.add(pictureError);
                                //model.setPictureError(pictureError);
                            }
                        }
                    }else{
                        String field = item.getFieldName();
                        String value = Streams.asString(stream);
                        if(field.equalsIgnoreCase(EMAIL_PARA)){
                            model.setEmail(value);
                            logger.info("email " + value);
                        }else if(field.equalsIgnoreCase(DESCRIPTION_PARA)){
                            model.setDescription(value);
                            logger.info("description " + value);
                        }else if(field.equalsIgnoreCase(VERSION_PARA)){
                            model.setVersion(value);
                            logger.info("version " + value);
                        }else if(field.equalsIgnoreCase(OS_INFORMATION_PARA)){
                            model.setOsInformation(value);
                            logger.info("os information " + value);
                        }else if(field.equalsIgnoreCase(STEP_TO_ERROR_PARA)){
                            model.setStepError(value);
                            logger.info("step get error " + value);
                        }
                    }
                }
                if(model.getEmail()!=null && model.getDescription()!=null){
                    if(StringUtil.List2String(screenshotFiles)!=null){
                        model.setPictureError(StringUtil.List2String(screenshotFiles));
                    }
                    if(StringUtil.List2String(testDataFiles)!=null){
                        model.setTestData(StringUtil.List2String(testDataFiles));
                    }
                    //add feedback to database
                    dao.create(model);
                    MailServices mailServices = new MailServices(model);
                    mailServices.start();
                    request.setAttribute("result","success");
                    request.getRequestDispatcher("index.jsp").forward(request,response);
                }
            }else{
                response.sendRedirect("index.jsp");
            }
        }catch (Exception e){
            request.setAttribute("result","fail");
            request.getRequestDispatcher("index.jsp").forward(request, response);
            MailServices mail = new MailServices(e.getMessage());
            mail.start();
            logger.error(e.getMessage());
            e.printStackTrace();
        }
    }

    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        doPost(request,response);
    }
}
