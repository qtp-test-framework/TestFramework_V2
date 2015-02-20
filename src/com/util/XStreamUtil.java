
package com.util;

import com.pojo.MailTemplate;
import com.thoughtworks.xstream.XStream;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;


public class XStreamUtil {
    
    public MailTemplate load_data_from_XML(String filePath){
        StringBuffer xml_str;
        MailTemplate mailTemplate = null;
        try {
            mailTemplate = new MailTemplate("", "", "", "");
            
            //Load the XML file
            File xml_file = new File(filePath);

            //create a new blank file and return if XML file not found
            if (!xml_file.exists()) {
                xml_file.createNewFile();
                return mailTemplate;
            }

            FileReader f_read = new FileReader(xml_file.getAbsoluteFile());
            BufferedReader br = new BufferedReader(f_read);
            xml_str = new StringBuffer();
            String sCurrentLine = "";

            XStream xstream = new XStream();
            xstream.alias("template", MailTemplate.class);

            while ((sCurrentLine = br.readLine()) != null) {
                xml_str.append(sCurrentLine);
            }
            System.out.println("xml_str = " + xml_str);
            //XML to Object Conversion
            mailTemplate = (MailTemplate) xstream.fromXML(xml_str.toString());

            br.close();

        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return mailTemplate;
    }
}
