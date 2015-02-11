package com.main;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

public class LoadExcelDts_Dialog extends JDialog {

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

    public LoadExcelDts_Dialog(JFrame parentFrame, String title, String excelPath, String testCaseFolderPath) {
        super(parentFrame, title, true);
        this.parent_Window = (MainWindow) parentFrame;
        
        //If user had selected a excel file/Test Case Folder earlier then load it by default
        this.previous_ExcelPath = (excelPath==null) ? "" : excelPath;
        this.previous_TestCaseFolderPath = (testCaseFolderPath==null) ? "" : testCaseFolderPath;

        if (parentFrame != null) {
            Dimension parentSize = parentFrame.getSize();
            Point p = parentFrame.getLocation();
            setLocation(p.x + parentSize.width / 4, p.y + parentSize.height / 4);
        }

        JPanel panel = new JPanel(new GridBagLayout());
        getContentPane().add(panel, BorderLayout.NORTH);
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

        this.setSize(500, 200);
        this.setDefaultCloseOperation(DISPOSE_ON_CLOSE);
        this.pack();
        this.setResizable(false);
        this.setVisible(true);
    }

    public void excel_ButtonClick(ActionEvent ae) {
        try {
            //If user had not selected a excel file earlier then load user's home directory
            if(previous_ExcelPath == null || previous_ExcelPath.equalsIgnoreCase("")){
                previous_ExcelPath = System.getProperty("user.home");
            }
            jExcelPathChooser = new JFileChooser(previous_ExcelPath);
            jExcelPathChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
            jExcelPathChooser.setDialogTitle("Select Test Case Excel file");
            jExcelPathChooser.setAcceptAllFileFilterUsed(false);

            //Only allow .xls & .xlsx files
            jExcelPathChooser.addChoosableFileFilter(new FileNameExtensionFilter("Microsoft Excel Documents", "xls", "xlsx"));

            int result = jExcelPathChooser.showOpenDialog(this);
            if (result == JFileChooser.APPROVE_OPTION) {
                File file = jExcelPathChooser.getSelectedFile();
                if (!file.getName().endsWith("xls")) {
                    JOptionPane.showMessageDialog(null, "Please select only Excel file.", "Error", JOptionPane.ERROR_MESSAGE);
                } else {
                    previous_ExcelPath = file.getPath();
                    txt_ExcelPath.setText(file.getPath());
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
                    String testCaseFolder = jTestCasePathChooser.getSelectedFile().getPath();
                    previous_TestCaseFolderPath = testCaseFolder;
                    txt_TesctCaseFldr.setText(testCaseFolder);
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
            String testCaseFolder = txt_TesctCaseFldr.getText() + "\\";
            
            //replacing all spaces in the path with '$$', as spaces cause problems while passing to vbscript file
            excelPath = excelPath.replace(" ", "$$");
            testCaseFolder = testCaseFolder.replace(" ", "$$");
            
            if(parent_Window != null){
                parent_Window.setImportChanges(excelPath, testCaseFolder);
            }
            
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

    public class CustomButtonListener implements ActionListener {

        @Override
        public void actionPerformed(ActionEvent ae) {
            if (ae.getSource() == btn_BrowseExcel) {
                excel_ButtonClick(ae);
            } else if (ae.getSource() == btn_BrowseTestCase) {
                testCaseFldr_ButtonClick(ae);
            } else if (ae.getSource() == btn_Import) {
                import_ButtonClick(ae);
                dispose();
            } else if (ae.getSource() == btn_CloseWindow) {
                closeWindow_ButtonClick(ae);
            }
        }
    }
}
