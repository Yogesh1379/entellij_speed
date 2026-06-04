package Send_Mail_Package;

import jakarta.mail.*;
import jakarta.mail.internet.*;

import java.io.File;
import java.io.IOException;
import java.util.Properties;

public class SendMailOf_ImportDataCreation {
    public static void sedReports(String outputPath) throws MessagingException, IOException {
        String SenderEmailID = "yogesh.dangade@winnersoft.co.in";
        String SenderPassword = "mnkj ddcp nblf ycqd";
        String ReceiverEmailID = "yogeshdangade123@gmail.com,nashik@winnersoft.co.in";


        Properties prop = new Properties();
        prop.setProperty("mail.smtp.host", "smtp.gmail.com");
        prop.setProperty("mail.smtp.auth", "true");
        prop.setProperty("mail.smtp.starttls.enable", "true");
        prop.setProperty("mail.smtp.port", "587");
        Session session = Session.getInstance(prop, new Authenticator() {
            @Override
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(SenderEmailID, SenderPassword);
            }
        });
        MimeMessage message = new MimeMessage(session);

        message.setFrom(new InternetAddress(SenderEmailID));

        message.setRecipients(MimeMessage.RecipientType.BCC, InternetAddress.parse(ReceiverEmailID));
        message.setSubject("Student Import Data with Form serial number July 2026");
        MimeBodyPart messagebodyPart = new MimeBodyPart();
        messagebodyPart.setText( "Dear Team,\r\n"
                + "\r\n"
                + "Please find attached file the Student import data with form serial number generated through automation.\r\n"
                + "\r\n"
                + "The report contains form serial number and in import format.\r\n"
                + "\r\n"
                + "Regards,\r\n"
                + "QA Automation Team");

        MimeBodyPart attachment1 = new MimeBodyPart();
        MimeBodyPart attachment2 = new MimeBodyPart();
        attachment1.attachFile(new File(outputPath));
        MimeMultipart multipart = new MimeMultipart();
        multipart.addBodyPart(messagebodyPart);
        multipart.addBodyPart(attachment1);
//        multipart.addBodyPart(attachment2);
        message.setContent(multipart);
        Transport.send(message);

        System.out.println("Mail Sent Successfully");


    }
}
