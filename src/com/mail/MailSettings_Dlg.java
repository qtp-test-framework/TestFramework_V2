package com.mail;

import com.main.MainWindow;
import com.pojo.MailTemplate;
import com.thoughtworks.xstream.XStream;
import com.util.ConfigUtility;
import com.util.Constants;
import com.util.XStreamUtil;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.Properties;
import javax.swing.*;

public class MailSettings_Dlg extends JDialog {

    private MainWindow parent_Window;
    private static JTabbedPane tp;
    private static JPanel tab_SMTP;
    private static JPanel tab_MailTemp;
    private static JPanel panel_BtnPanel;
    //SMPT Tab components
    private JLabel lbl_Host;
    private JLabel lbl_Port;
    private JLabel lbl_User;
    private JLabel lbl_Pass;
    private JTextField jt_Host;
    private JTextField jt_Port;
    private JTextField jt_User;
    private JTextField jt_Pass;
    //Mail Template Tab components
    private JLabel lbl_To;
    private JLabel lbl_Subject;
    private JLabel lbl_CC;
    private JTextField jt_To;
    private JTextField jt_Subject;
    private JTextField jt_CC;
    private JTextArea textArea_Message;
    //buttons panel components
    private static JButton btn_Apply;
    private static JButton btn_Ok;
    private static JButton btn_Cancel;
    ConfigUtility configUtil = new ConfigUtility();

    public MailSettings_Dlg(JFrame parentFrame, String title) {
        super(parentFrame, title, true);
        this.parent_Window = (MainWindow) parentFrame;

        //setting padding for the border layout
        int hGap = 0, vGap = 5;
        this.setLayout(new BorderLayout(hGap, vGap));

        createUI_SMTP();
        load_SMTP_Data();

        createUI_MailTemp();
        load_Template_Data();

        createUI_Buttons();
        createUI(parent_Window, this);
    }

    public void createUI_SMTP() {
        tab_SMTP = new JPanel(new GridBagLayout());

        GridBagConstraints constraints = new GridBagConstraints();
        constraints.anchor = GridBagConstraints.NORTHWEST;

        lbl_Host = new JLabel("Host name: ");
        constraints.gridx = 0;
        constraints.gridy = 0;
        constraints.weightx = 1;
        constraints.weighty = 0;
        constraints.insets = new Insets(30, 20, 5, 10);
        tab_SMTP.add(lbl_Host, constraints);

        jt_Host = new JTextField(20);
        constraints.gridx = 1;
        constraints.gridy = 0;
        constraints.weightx = 10;
        constraints.weighty = 0;
        constraints.insets = new Insets(30, 10, 5, 10);
        tab_SMTP.add(jt_Host, constraints);

        lbl_Port = new JLabel("Port number: ");
        constraints.gridx = 0;
        constraints.gridy = 1;
        constraints.weightx = 1;
        constraints.weighty = 0;
        constraints.insets = new Insets(10, 20, 5, 10);
        tab_SMTP.add(lbl_Port, constraints);

        jt_Port = new JTextField(20);
        constraints.gridx = 1;
        constraints.gridy = 1;
        constraints.weightx = 1;
        constraints.weighty = 0;
        constraints.insets = new Insets(10, 10, 5, 10);
        tab_SMTP.add(jt_Port, constraints);

        lbl_User = new JLabel("Username: ");
        constraints.gridx = 0;
        constraints.gridy = 2;
        constraints.weightx = 1;
        constraints.weighty = 0;
        constraints.insets = new Insets(10, 20, 5, 10);
        tab_SMTP.add(lbl_User, constraints);

        jt_User = new JTextField(20);
        constraints.gridx = 1;
        constraints.gridy = 2;
        constraints.weightx = 1;
        constraints.weighty = 0;
        constraints.insets = new Insets(10, 10, 5, 10);
        tab_SMTP.add(jt_User, constraints);

        lbl_Pass = new JLabel("Password: ");
        constraints.gridx = 0;
        constraints.gridy = 3;
        constraints.weightx = 1;
        constraints.weighty = 0;
        constraints.insets = new Insets(10, 20, 5, 10);
        tab_SMTP.add(lbl_Pass, constraints);

        jt_Pass = new JTextField(20);
        constraints.gridx = 1;
        constraints.gridy = 3;
        constraints.weightx = 1;
        constraints.weighty = 1;
        constraints.insets = new Insets(10, 10, 5, 10);
        tab_SMTP.add(jt_Pass, constraints);

    }

    public void createUI_MailTemp() {
        tab_MailTemp = new JPanel(new GridBagLayout());
        GridBagConstraints constraints = new GridBagConstraints();
        constraints.insets = new Insets(10, 10, 5, 10);
        constraints.anchor = GridBagConstraints.WEST;

        lbl_To = new JLabel("To: ");
        constraints.gridx = 0;
        constraints.gridy = 0;
        tab_MailTemp.add(lbl_To, constraints);

        jt_To = new JTextField(30);
        constraints.gridx = 1;
        constraints.gridy = 0;
        constraints.fill = GridBagConstraints.HORIZONTAL;
        tab_MailTemp.add(jt_To, constraints);

        lbl_CC = new JLabel("CC: ");
        constraints.gridx = 0;
        constraints.gridy = 1;
        tab_MailTemp.add(lbl_CC, constraints);

        jt_CC = new JTextField(30);
        constraints.gridx = 1;
        constraints.gridy = 1;
        constraints.fill = GridBagConstraints.HORIZONTAL;
        tab_MailTemp.add(jt_CC, constraints);

        lbl_Subject = new JLabel("Subject: ");
        constraints.gridx = 0;
        constraints.gridy = 2;
        tab_MailTemp.add(lbl_Subject, constraints);

        jt_Subject = new JTextField(30);
        constraints.gridx = 1;
        constraints.gridy = 2;
        constraints.fill = GridBagConstraints.HORIZONTAL;
        tab_MailTemp.add(jt_Subject, constraints);

        textArea_Message = new JTextArea(7, 30);
        textArea_Message.setLineWrap(true);    // If text doesn't fit on a line, jump to the next
        textArea_Message.setWrapStyleWord(true);   // Makes sure that words stay intact if a line wrap occurs
        JScrollPane scrollbar_Body = new JScrollPane(textArea_Message, JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED, JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        constraints.gridx = 1;
        constraints.gridy = 4;
        constraints.weightx = 1.0;
        constraints.weighty = 1.0;
        constraints.fill = GridBagConstraints.HORIZONTAL;
        tab_MailTemp.add(scrollbar_Body, constraints);

        //tool tip msg to user
        JLabel lbl_TO_Msg = new JLabel(Constants.MAIL_TOOL_TIP);
        java.awt.Font tooltip_font = new java.awt.Font("TimesRoman", Font.PLAIN, 10);
        lbl_TO_Msg.setForeground(Color.GRAY);
        lbl_TO_Msg.setFont(tooltip_font);
        constraints.gridx = 1;
        constraints.gridy = 0;
        constraints.weightx = 0;
        constraints.weighty = 0;
        constraints.insets = new Insets(39, 10, 0, 10);
        tab_MailTemp.add(lbl_TO_Msg, constraints);
    }

    public void createUI_Buttons() {
        panel_BtnPanel = new JPanel(new FlowLayout());

        btn_Apply = new JButton("Apply");
        panel_BtnPanel.add(btn_Apply);

        btn_Ok = new JButton("Ok");
        panel_BtnPanel.add(btn_Ok);

        btn_Cancel = new JButton("Cancel");
        panel_BtnPanel.add(btn_Cancel);

        //add listeners
        btn_Apply.addActionListener(new CustomActionListener());
        btn_Ok.addActionListener(new CustomActionListener());
        btn_Cancel.addActionListener(new CustomActionListener());
    }

    public void createUI(MainWindow parentFrame, MailSettings_Dlg this_frame) {

        JLabel lbl_mail = new JLabel(" Mail Settings");
        java.awt.Font font_Heading = new java.awt.Font("TimesRoman", Font.BOLD, 20);
        lbl_mail.setFont(font_Heading);
        this_frame.add(lbl_mail, BorderLayout.NORTH);

        //initialising the tabbed pane view and adding Tabs
        tp = new JTabbedPane();
        tp.addTab("SMTP Settings", tab_SMTP);
        tp.addTab("Mail Template", tab_MailTemp);
        this_frame.add(tp, BorderLayout.CENTER);

        //this panel will contain all the action bttns :Apply/Ok/Cancel
        this_frame.add(panel_BtnPanel, BorderLayout.SOUTH);

        //set window location relative to its parent window
        if (parentFrame != null) {
            Dimension parentSize = parentFrame.getSize();
            Point p = parentFrame.getLocation();
            this_frame.setLocation(p.x + parentSize.width / 6, p.y + parentSize.height / 7);
        }

        this_frame.setPreferredSize(new Dimension(500, 400));
        this_frame.setMaximumSize(new Dimension(500, 400));
        this_frame.setMinimumSize(new Dimension(500, 400));
        this_frame.setDefaultCloseOperation(DISPOSE_ON_CLOSE);
        this_frame.setResizable(false);
        this_frame.pack();
        this_frame.setVisible(true);
    }

    private void load_SMTP_Data() {
        Properties configProps = null;
        try {
            configProps = configUtil.loadProperties();
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this,
                    "Error reading settings: " + ex.getMessage(),
                    "Error", JOptionPane.ERROR_MESSAGE);
        }

        jt_Host.setText(configProps.getProperty("mail.smtp.host"));
        jt_Port.setText(configProps.getProperty("mail.smtp.port"));
        jt_User.setText(configProps.getProperty("mail.user"));
        jt_Pass.setText(configProps.getProperty("mail.password"));
    }

    private boolean saveMailSettings(ActionEvent ae) {
        try {
            //1. Validations
            if (!validateMailSettings()) {
                return false;
            }
            //2. Save SMTP Settings
            saveSMTP_Settings();

            //3. Save Mail Template Settings
            saveTemplate_Settings();

        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return true;
    }

    private boolean validateMailSettings() {
        boolean result = true;
        try {
            if (jt_Host.getText().equalsIgnoreCase("")) {
                JOptionPane.showMessageDialog(this,
                        "Host cannot be blank!",
                        "Incomplete Input!!", JOptionPane.WARNING_MESSAGE);
                jt_Host.requestFocus();
                return false;
            } else if (jt_Port.getText().equalsIgnoreCase("")) {
                JOptionPane.showMessageDialog(this,
                        "Port cannot be blank!",
                        "Incomplete Input!!", JOptionPane.WARNING_MESSAGE);
                jt_Port.requestFocus();
                return false;
            } else if (jt_User.getText().equalsIgnoreCase("")) {
                JOptionPane.showMessageDialog(this,
                        "User Name cannot be blank!",
                        "Incomplete Input!!", JOptionPane.WARNING_MESSAGE);
                jt_User.requestFocus();
                return false;
            } else if (jt_Pass.getText().equalsIgnoreCase("")) {
                JOptionPane.showMessageDialog(this,
                        "Password cannot be blank!",
                        "Incomplete Input!!", JOptionPane.WARNING_MESSAGE);
                jt_Pass.requestFocus();
                return false;
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return result;
    }

    private void saveSMTP_Settings() {
        try {
            configUtil.saveProperties(jt_Host.getText(),
                    jt_Port.getText(),
                    jt_User.getText(),
                    jt_Pass.getText());
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    private void saveTemplate_Settings() {
        try {
            XStream xstream = new XStream();
            xstream.alias("template", MailTemplate.class);

            MailTemplate mailTemplate = new MailTemplate(jt_To.getText(),
                    jt_CC.getText(),
                    jt_Subject.getText(),
                    textArea_Message.getText());

            //Object to XML Conversion
            String xml_str = xstream.toXML(mailTemplate);

            //Load the XML file
            File to_xml_file = new File(Constants.MAIL_TEMPLATE_FILE);
            //if file doesnt exists then create a new file
            if (!to_xml_file.exists()) {
                to_xml_file.createNewFile();
            }

            //Write to the XML file
            FileWriter fw = new FileWriter(to_xml_file.getAbsoluteFile());
            BufferedWriter br = new BufferedWriter(fw);
            br.write(xml_str);
            br.close();

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    private void load_Template_Data() {
        try {
            XStreamUtil xUtil = new XStreamUtil();
            MailTemplate mailTemplate = (MailTemplate) xUtil.load_data_from_XML(Constants.MAIL_TEMPLATE_FILE);
            
            jt_To.setText(mailTemplate.getTo());
            jt_CC.setText(mailTemplate.getCc());
            jt_Subject.setText(mailTemplate.getSubject());
            textArea_Message.setText(mailTemplate.getBody());

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    public void disposeForm() {
        try {
            dispose();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    private void toggle_Apply_Btn(boolean vAction) {
        btn_Apply.setEnabled(vAction);
    }

    private class CustomActionListener implements ActionListener {

        @Override
        public void actionPerformed(ActionEvent ae) {
            if (ae.getSource() == btn_Apply) {
                if (saveMailSettings(ae)) {
                    toggle_Apply_Btn(false);       //disable Apply btn only if validations pass
                }
            } else if (ae.getSource() == btn_Ok) {
                saveMailSettings(ae);
                disposeForm();
            } else if (ae.getSource() == btn_Cancel) {
                disposeForm();
            }
        }
    }
}
