
package com.util;

import java.io.*;
import java.util.Properties;
 

public class ConfigUtility {
    private File configFile = new File(Constants.PROPERTIES_FILE);
    private Properties configProps;
     
    public Properties loadProperties() throws IOException {
        Properties defaultProps = new Properties();
        // sets default properties
        defaultProps.setProperty("mail.smtp.host", "smtp.gmail.com");
        defaultProps.setProperty("mail.smtp.port", "587");
        defaultProps.setProperty("mail.user", "");
        defaultProps.setProperty("mail.password", "");
        defaultProps.setProperty("mail.smtp.starttls.enable", "true");
        defaultProps.setProperty("mail.smtp.auth", "true");
        defaultProps.setProperty("mail.smtp.ssl.trust", "smtp.gmail.com");
         
        configProps = new Properties(defaultProps);
         
        // loads properties from file
        if (configFile.exists()) {
            InputStream inputStream = new FileInputStream(configFile);
            configProps.load(inputStream);
            inputStream.close();
        }
        return configProps;
    }
     
    public void saveProperties(String host, String port, String user, String pass) throws IOException {
        configProps.setProperty("mail.smtp.host", host.trim());
        configProps.setProperty("mail.smtp.port", port.trim());
        configProps.setProperty("mail.user", user.trim());
        configProps.setProperty("mail.password", pass.trim());
        configProps.setProperty("mail.smtp.starttls.enable", "true");
        configProps.setProperty("mail.smtp.auth", "true");
        configProps.setProperty("mail.smtp.ssl.trust", "smtp.gmail.com");
         
        System.out.println("host = "+host);
        
        OutputStream outputStream = new FileOutputStream(configFile);
        configProps.store(outputStream, "host setttings");
        outputStream.close();
    }  
}