
package com.util;

import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.util.Properties;
import javax.mail.*;
import javax.mail.internet.*;


public class MailClient {

    public static void sendEmail(Properties smtpProperties, String toAddress,
            String subject, String message, File attachedFile)
            throws AddressException, MessagingException, IOException {
            final String userName = smtpProperties.getProperty("mail.user");
            final String password = smtpProperties.getProperty("mail.password");

            System.out.println("userName = " + userName);
            System.out.println("password = " + password);
            System.out.println("toAddress = " + toAddress);

            // creates a new session with an authenticator
            Authenticator auth = new Authenticator() {
                public PasswordAuthentication getPasswordAuthentication() {
                    return new PasswordAuthentication(userName, password);
                }
            };
            Session session = Session.getInstance(smtpProperties, auth);

            // creates a new e-mail message
            Message msg = new MimeMessage(session);

            msg.setFrom(new InternetAddress(userName));
            InternetAddress[] toAddresses = {new InternetAddress(toAddress)};
            msg.setRecipients(Message.RecipientType.TO, toAddresses);
            msg.setSubject(subject);
            msg.setSentDate(new Date());

            // creates message part
            MimeBodyPart messageBodyPart = new MimeBodyPart();
            messageBodyPart.setContent(message, "text/html");

            // creates multi-part
            Multipart multipart = new MimeMultipart();
            multipart.addBodyPart(messageBodyPart);

            // adds attachments
            MimeBodyPart attachPart = new MimeBodyPart();
            attachPart.attachFile(attachedFile);
            multipart.addBodyPart(attachPart);

            // sets the multi-part as e-mail's content
            msg.setContent(multipart);
            
            // sends the e-mail
            Transport.send(msg);

        
    }
}