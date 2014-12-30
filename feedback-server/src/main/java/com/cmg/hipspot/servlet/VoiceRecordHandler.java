package com.cmg.hipspot.servlet;

import com.cmg.hipspot.data.dao.impl.FeedbackModelDAO;
import com.cmg.hipspot.data.jdo.FeedbackModel;
import com.cmg.hipspot.properties.Configuration;
import com.cmg.hipspot.service.FileServices;
import com.cmg.hipspot.service.MailServices;
import com.cmg.hipspot.util.FileHelper;
import com.cmg.hipspot.util.StringUtil;
import org.apache.commons.fileupload.FileItemIterator;
import org.apache.commons.fileupload.FileItemStream;
import org.apache.commons.fileupload.FileUploadException;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.commons.fileupload.util.Streams;
import org.apache.commons.io.FileUtils;
import org.apache.log4j.Logger;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by lantb on 2014-04-22.
 */
public class VoiceRecordHandler extends HttpServlet {
    private static final Logger logger = Logger.getLogger(VoiceRecordHandler.class
            .getName());
    private static String PARA_FILE_NAME = "FILE_NAME";
    private static String PARA_FILE_PATH = "FILE_PATH";
    private static String PARA_FILE_TYPE = "FILE_TYPE";
    private static String PARA_UUID = "uuid";
    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        try {
            //create a new Map<String,String> to store all parameter
            Map<String, String> storePara = new HashMap<String, String>();
            // Create a new file upload handler
            ServletFileUpload upload = new ServletFileUpload();
            // Parse the request
            FileItemIterator iter = null;
            iter = upload.getItemIterator(request);
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
            String fileTempName  = sdf.format(new Date(System.currentTimeMillis())) + ".wav";
            String targetDir = Configuration.getValue(Configuration.VOICE_RECORD_DIR);
            String tmpFile = UUID.randomUUID().toString();
            String tmpDir = System.getProperty("java.io.tmpdir");
            while (iter.hasNext()) {
                FileItemStream item = iter.next();
                String name = item.getFieldName();
                InputStream stream = item.openStream();
                if (item.isFormField()) {
                    logger.info(name);
                    String value = Streams.asString(stream);
                    storePara.put(name, value);
                }else{
                    String getName = item.getName();
                    logger.info("getname = :" +getName);
                    // Process the input stream
                    if(getName.endsWith(".wav")){
                        FileHelper.saveFile(tmpDir, tmpFile, stream);
                    }
                }
            }
            String uuid = storePara.get(PARA_UUID).toString();
            if (uuid != null && uuid.length() > 0) {
                File target = new File(targetDir, uuid);
                if (!target.exists() && !target.isDirectory()) {
                    target.mkdirs();
                }
                File tmpFileIn = new File(tmpDir, tmpFile);

                FileUtils.moveFile(tmpFileIn, new File(target, fileTempName));
                tmpFileIn.delete();
            }
        } catch (FileUploadException e) {
            logger.error("Error when upload file. FileUploadException, message: " + e.getMessage(),e);
        } catch (Exception e) {
            logger.error("Error when upload file. Common exception, message: " + e.getMessage(),e);
        }
    }

    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        doPost(request,response);
    }
}
