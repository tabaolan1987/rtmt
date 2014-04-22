package com.cmg.hipspot.service;

import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileItemStream;

import javax.servlet.http.Part;
import java.io.*;
import java.util.Properties;
import java.util.UUID;

/**
 * Created by lantb on 2014-04-21.
 */
public class FileServices {
    public static final String SYSTEM_PROPERTIES = "system.properties";

    public static final String XML_URL_PENSIONER = "xml.url.pensioner";
    public static final String XML_URL_EMPLOYEE = "xml.url.employee";
    public static final String NEWSLETTER_IMAGE_URL = "newsletter.image.url";

    public static final String PROJECT_LIST = "project.list";
    public static final String PROJECT_DIR = "project.dir";
    public static final String TEMP_FOLDER="project.temp.folder.dir";
    public static final String ROOT_FOLDER = "project.root.folder.dir";
    private static Properties prop;


    public String saveFile(InputStream item, String path, String FileName) throws IOException{
        OutputStream out = null;
        InputStream filecontent = null;
        FileName = setFileName(FileName);
        File file = null;
        try {
            file = new File(path + File.separator
                    + FileName);
            out = new FileOutputStream(file);
            filecontent = item;
            int read = 0;
            final byte[] bytes = new byte[1024];

            while ((read = filecontent.read(bytes)) != -1) {
                out.write(bytes, 0, read);
            }
        } catch (FileNotFoundException fne) {
            fne.printStackTrace();
        } finally {
            if (out != null) {
                out.close();
            }
        }
        return file.getAbsolutePath();
    }

    public String saveFile(Part filePart, String fileName, String path) throws IOException{
        OutputStream out = null;
        InputStream filecontent = null;
        File file = null;
        try {
            file = new File(path + File.separator
                    + fileName);
            out = new FileOutputStream(file);
            filecontent = filePart.getInputStream();
            int read = 0;
            final byte[] bytes = new byte[1024];

            while ((read = filecontent.read(bytes)) != -1) {
                out.write(bytes, 0, read);
            }
        } catch (FileNotFoundException fne) {
            fne.printStackTrace();
        } finally {
            if (out != null) {
                out.close();
            }
            if (filecontent != null) {
                filecontent.close();
            }
        }
        return file.getAbsolutePath();
    }

    public String getPath(String folder){
        String tempFolder = getValue(TEMP_FOLDER);
        if(tempFolder!=""){
            File foldertemp = new File(tempFolder);
            if (!foldertemp.exists()){
                foldertemp.mkdirs();
            }
            File file = new File(tempFolder + File.separator + folder);
            if(!file.exists()){
                file.mkdirs();
            }
            return file.getAbsolutePath();
        }
        return null;
    }

    public String setFileName(String clientName){
      String uniqueFile = UUID.randomUUID().toString()+ "-" + clientName;
      return uniqueFile;
    }
    public static String getValue(String key) {
        if (prop == null)
            getProperties(SYSTEM_PROPERTIES);
        return prop != null ? prop.getProperty(key) : "";
    }

    public static void getProperties(String propName) {
        prop = new Properties();
        try {
            // load a properties file from class path, inside static method
            prop.load(FileServices.class.getClassLoader().getResourceAsStream(
                    propName));
        } catch (Exception ex) {
        }
    }


}
