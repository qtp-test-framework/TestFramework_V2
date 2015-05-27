package com.main;

import com.util.Constants;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.util.prefs.Preferences;

public class LoadExcelDts_Dlg extends JDialog {

    static JTextField txt_ExcelPath;
    static JTextField txt_TesctCaseFldr;
    static JButton btn_BrowseExcel;
    static JButton btn_BrowseTestCase;
    static JButton btn_Import;
    static JButton btn_CloseWindow;
    static JFileChooser jExcelPathChooser;
    static JFileChooser jTestCasePathChooser;
    private MainWindow parent_Window;
    static String previous_ExcelPath;
    static String previous_TestCaseFolderPath;
    Preferences user_prefs;

    public LoadExcelDts_Dlg(JFrame parentFrame, String title, Preferences main_prefs) {
        super(parentFrame, title, true);
        this.parent_Window = (MainWindow) parentFrame;
        this.user_prefs = main_prefs;

        //If user had selected a excel file/Test Case Folder earlier then load it by default 
        this.previous_ExcelPath = user_prefs.get(Constants.EXCEL_PREF, System.getProperty("user.home"));
        this.previous_TestCaseFolderPath = user_prefs.get(Constants.TEST_FOLDER_PREF, System.getProperty("user.home"));

        //set window location relative to its parent window
        if (parentFrame != null) {
            Dimension parentSize = parentFrame.getSize();
            Point p = parentFrame.getLocation();
            setLocation(p.x + parentSize.width / 7, p.y + parentSize.height / 4);
        }

        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints gbConst = new GridBagConstraints();

        JLabel lbl_UsrName = new JLabel("Excel File Path: ");
        gbConst.gridy = 0;
        gbConst.gridx = 0;
        gbConst.insets = new Insets(10, 10, 10, 10);
        panel.add(lbl_UsrName, gbConst);

        txt_ExcelPath = new JTextField(this.previous_ExcelPath, 35);
        txt_ExcelPath.setEditable(false);
        gbConst.gridy = 0;
        gbConst.gridx = 1;
        gbConst.insets = new Insets(10, 10, 10, 10);
        panel.add(txt_ExcelPath, gbConst);

        btn_BrowseExcel = new JButton("Browse");
        gbConst.gridy = 0;
        gbConst.gridx = 2;
        gbConst.insets = new Insets(10, 10, 10, 10);
        panel.add(btn_BrowseExcel, gbConst);
        btn_BrowseExcel.addActionListener(new CustomButtonListener());

        JLabel lbl_Psswd = new JLabel("Test Case Folder: ");
        gbConst.gridy = 1;
        gbConst.gridx = 0;
        gbConst.insets = new Insets(10, 10, 10, 10);
        panel.add(lbl_Psswd, gbConst);

        txt_TesctCaseFldr = new JTextField(this.previous_TestCaseFolderPath, 35);
        txt_TesctCaseFldr.setEditable(false);
        gbConst.gridy = 1;
        gbConst.gridx = 1;
        gbConst.insets = new Insets(10, 10, 10, 10);
        panel.add(txt_TesctCaseFldr, gbConst);

        btn_BrowseTestCase = new JButton("Browse");
        gbConst.gridy = 1;
        gbConst.gridx = 2;
        gbConst.insets = new Insets(10, 10, 10, 10);
        panel.add(btn_BrowseTestCase, gbConst);
        btn_BrowseTestCase.addActionListener(new CustomButtonListener());

        btn_Import = new JButton("Import");
        gbConst.gridx = 1;
        gbConst.gridy = 2;
        gbConst.ipadx = 40;
        gbConst.insets = new Insets(10, 0, 10, 10);
        gbConst.anchor = GridBagConstraints.WEST;
        panel.add(btn_Import, gbConst);
        btn_Import.addActionListener(new CustomButtonListener());

        btn_CloseWindow = new JButton("Close");
        gbConst.gridx = 1;
        gbConst.gridy = 2;
        gbConst.ipadx = 40;
        gbConst.insets = new Insets(10, 0, 10, 10);
        gbConst.anchor = GridBagConstraints.EAST;
        panel.add(btn_CloseWindow, gbConst);
        btn_CloseWindow.addActionListener(new CustomButtonListener());

        getContentPane().add(panel, BorderLayout.NORTH);

        this.setSize(500, 200);
        this.setDefaultCloseOperation(DISPOSE_ON_CLOSE);
        this.pack();
        this.setResizable(false);
        this.setVisible(true);
    }

    public void excel_ButtonClick(ActionEvent ae) {
        try {
            //If user had not selected a excel file earlier then load user's home directory
//            if (previous_ExcelPath == null || previous_ExcelPath.equalsIgnoreCase("")) {
//                previous_ExcelPath = System.getProperty("user.home");
//            }
            jExcelPathChooser = new JFileChooser(previous_ExcelPath);
            jExcelPathChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
            jExcelPathChooser.setDialogTitle("Select Test Case Excel file");
            jExcelPathChooser.setAcceptAllFileFilterUsed(false);

            //Only allow .xls & .xlsx files
            jExcelPathChooser.addChoosableFileFilter(new FileNameExtensionFilter("Microsoft Excel Documents", "xls", "xlsx"));

            int result = jExcelPathChooser.showOpenDialog(this);
            if (result == JFileChooser.APPROVE_OPTION) {
                File file = jExcelPathChooser.getSelectedFile();
                if (file.getName().endsWith("xls") || file.getName().endsWith("xlsx")) {
                    txt_ExcelPath.setText(file.getPath());
                    user_prefs.put(Constants.EXCEL_PREF, file.getPath());

                } else {
                    JOptionPane.showMessageDialog(null, "Please select only Excel file.", "Error", JOptionPane.ERROR_MESSAGE);
                }
            } else {
                System.out.println("Cancel button clicked");
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    public void testCaseFldr_ButtonClick(ActionEvent ae) {
        try {
            jTestCasePathChooser = new JFileChooser(previous_TestCaseFolderPath);
            jTestCasePathChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            jTestCasePathChooser.setAcceptAllFileFilterUsed(false);
            jTestCasePathChooser.setDialogTitle("Select QTP Test Case Directory");

            int result = jTestCasePathChooser.showOpenDialog(null);
            if (result == JFileChooser.APPROVE_OPTION) {
                if (jTestCasePathChooser.getCurrentDirectory() != null) {
                    String testCaseFolder = jTestCasePathChooser.getSelectedFile().getPath() + "\\";
//                    previous_TestCaseFolderPath = testCaseFolder;
                    txt_TesctCaseFldr.setText(testCaseFolder);
                    user_prefs.put(Constants.TEST_FOLDER_PREF, testCaseFolder);
                }
            } else {
                System.out.println("Cancel button clicked");
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    public void import_ButtonClick(ActionEvent ae) {
        try {
            String excelPath = txt_ExcelPath.getText();
            String testCaseFolder = txt_TesctCaseFldr.getText();

            if (!validateFields(excelPath, testCaseFolder)) {
                return;
            }

            //Setting the Global XLSX Constant variable
            if (GetFileExtension(excelPath).equalsIgnoreCase("xlsx")) {
                Constants.IS_XLSX = true;
            }else{
                Constants.IS_XLSX = false;
            }

            //replacing all spaces in the path with '$$', as spaces cause problems while passing to vbscript file
//            excelPath = excelPath.replace(" ", "$$");
//            testCaseFolder = testCaseFolder.replace(" ", "$$");

            if (parent_Window != null) {
                parent_Window.setImportChanges(excelPath, testCaseFolder);
            }

            //Close current Frame after saving..
            dispose();

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    public void closeWindow_ButtonClick(ActionEvent ae) {
        try {
            dispose();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    private boolean validateFields(String excelPath, String testCaseFolder) {
        try {
            //chk if Excel has not been imported by the user
            if (excelPath == null || excelPath.equals("")) {
                JOptionPane.showMessageDialog(this,
                        "Please select the Test Cases excel file!",
                        "Incomplete Input!!", JOptionPane.WARNING_MESSAGE);
                btn_BrowseExcel.requestFocus();
                return false;
            }

            //chk if Test Case Folder has not been set by the user
            if (testCaseFolder == null || testCaseFolder.equals("")) {
                JOptionPane.showMessageDialog(this,
                        "Please select the Test Cases folder !",
                        "Incomplete Input!!", JOptionPane.WARNING_MESSAGE);
                btn_BrowseTestCase.requestFocus();
                return false;
            }


        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return true;
    }

    private static String GetFileExtension(String fname2) throws Exception{
        String fileName = fname2;
        String ext = "";
        int mid = fileName.lastIndexOf(".");
        ext = fileName.substring(mid + 1, fileName.length());
        return ext;
    }

    public class CustomButtonListener implements ActionListener {

        @Override
        public void actionPerformed(ActionEvent ae) {
            if (ae.getSource() == btn_BrowseExcel) {
                excel_ButtonClick(ae);
            } else if (ae.getSource() == btn_BrowseTestCase) {
                testCaseFldr_ButtonClick(ae);
            } else if (ae.getSource() == btn_Import) {
                import_ButtonClick(ae);
            } else if (ae.getSource() == btn_CloseWindow) {
                closeWindow_ButtonClick(ae);
            }
        }
    }
}
