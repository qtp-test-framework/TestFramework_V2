/*
 * >x      ^Y
 * Insets(top, left, bottom, right);
 * 
 * To add border : lbl_ExcelPath.setBorder(BorderFactory.createMatteBorder(borderWidth, 0, borderWidth, borderWidth, Color.BLACK));
 */
package com.main;

import com.help.About_Dlg;
import com.mail.MailSettings_Dlg;
import com.pojo.MailTemplate;
import com.pojo.TestCase;
import com.util.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.ArrayList;
import java.util.Properties;
import java.util.Vector;
import java.util.prefs.Preferences;
import javax.mail.MessagingException;
import javax.mail.internet.AddressException;
import javax.swing.*;
import javax.swing.border.BevelBorder;
import javax.swing.border.EmptyBorder;
import javax.swing.border.EtchedBorder;
import javax.swing.table.TableColumn;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class MainWindow extends JFrame {

    // Variables declaration - do not modify                     
    //private static JButton btn_Load;
    //private static JButton btn_Execute;
    private static JLabel lbl_ExcelPath;
    private static JLabel lbl_TestCaseFolder;
    private static JTextField jt_ExcelPath_Val;
    private static JTextField jt_TestCaseFolder_Val;
    private static JLabel lbl_Logs;
    private static JPanel jPanel;
    private static JScrollPane jScrollPane1;
    private static JTable table;
    private static JTextArea jt_Execution_Logs;
    private static JMenuItem jSubMenu_Import;
    private static JMenuItem jSubMenu_Exit;
    private static JMenuItem jSubMenu_MailSttngs;
    private static JMenuItem jSubMenu_Execute;
    private static JMenuItem jSubMenu_Run_FStep;
    private static JMenuItem jSubMenu_About;
    //Toolbar buttons
    private static JButton btn_import_excel_TB;
    private static JButton btn_execute_TB;
    private static JButton btn_mail_sttngs_TB;
    private static JButton btn_clearData_TB;
    // End of variables declaration
    static CustomTableModel model = null;
    static Vector headers = new Vector();
    static Vector data = new Vector();
    static String gExcelPath = "";
    static String gTestCaseFolder = "";
    public JFrame main_frame;
    Preferences user_prefs;
    java.util.List<TestCase> testCaseSumm_list = null;  //this will hold the list of testCases with their pass/fail data
    static JProgressBar progressBar;

    public MainWindow(String title) {
        super(title);

        //Setting application icon image
        ImageIcon img_icon = new ImageIcon(Constants.ICON_IMAGE);
        setIconImage(img_icon.getImage());

        this.main_frame = (MainWindow) this;
        this.setLayout(new BorderLayout());

        loadUserPrefData();
        addToolBar(this);
        addFileMenu(this);
        addComponentsToPane(this.getContentPane());
        addComponent_Listeners();

        this.setLocationRelativeTo(null);
        this.setPreferredSize(new Dimension(Constants.MainWindow_WIDTH, Constants.MainWindow_HEIGHT));
        this.setMinimumSize(new Dimension(Constants.MainWindow_WIDTH, Constants.MainWindow_HEIGHT));
        this.setMaximumSize(new Dimension(Constants.MainWindow_WIDTH, Constants.MainWindow_HEIGHT));
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setResizable(false);

        //code to Load the frame in the center of the screen===============
        Toolkit toolkit = Toolkit.getDefaultToolkit();
        Dimension dimn = toolkit.getScreenSize();
        int xPos = (dimn.width / 2) - (this.getWidth() / 2);
        int ypos = (dimn.height / 2) - (this.getHeight() / 2);
        this.setLocation(xPos, ypos);
        //============================================================

        this.pack();
        this.setVisible(true);
    }

    //Code Moved to Splash Screen Class
//    public static void main(String args[]) {
//        try {
//            //Get System (Microsoft Windows) Look and Feel
//            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
//
//        } catch (ClassNotFoundException ex) {
//            java.util.logging.Logger.getLogger(MainWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
//        } catch (InstantiationException ex) {
//            java.util.logging.Logger.getLogger(MainWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
//        } catch (IllegalAccessException ex) {
//            java.util.logging.Logger.getLogger(MainWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
//        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
//            java.util.logging.Logger.getLogger(MainWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
//        }
//
//        EventQueue.invokeLater(new Runnable() {
//
//            public void run() {
//                new MainWindow(Constants.APPLICATION_NAME);
//            }
//        });
//    }
    public void loadUserPrefData() {
        //API for persistent storage of user preferences
        user_prefs = Preferences.userNodeForPackage(MainWindow.class);

        gExcelPath = user_prefs.get(Constants.EXCEL_PREF, "");
        gTestCaseFolder = user_prefs.get(Constants.TEST_FOLDER_PREF, "");
    }

    private void addToolBar(JFrame this_frame) {
        try {
            Insets margins = new Insets(0, 5, 0, 5);    //give spacing between buttons

            ImageIcon img_import_excel = new ImageIcon("src/Resources/Images/import_excel_24.png");
            btn_import_excel_TB = new JButton(img_import_excel);
            btn_import_excel_TB.setToolTipText("Import Test Cases from Excel");
            btn_import_excel_TB.addActionListener(new CustomButtonListener());
            btn_import_excel_TB.setMargin(margins);

            ImageIcon img_execute = new ImageIcon("src/Resources/Images/execute_24.png");
            btn_execute_TB = new JButton(img_execute);
            btn_execute_TB.setToolTipText("Execute the test cases");
            btn_execute_TB.addActionListener(new CustomButtonListener());
            btn_execute_TB.setMargin(margins);

            ImageIcon img_mail_sttngs = new ImageIcon("src/Resources/Images/mail_settings_dark_24.png"); //#1F7A3E 
            btn_mail_sttngs_TB = new JButton(img_mail_sttngs);
            btn_mail_sttngs_TB.setToolTipText("Mail Settings");
            btn_mail_sttngs_TB.addActionListener(new CustomButtonListener());
            btn_mail_sttngs_TB.setMargin(margins);

            ImageIcon img_clear = new ImageIcon("src/Resources/Images/btn_clearData_TB.png");   //#5BBB66
            btn_clearData_TB = new JButton(img_clear);
            btn_clearData_TB.setToolTipText("Clear All Data");
            btn_clearData_TB.addActionListener(new CustomButtonListener());
            btn_clearData_TB.setMargin(margins);

            JToolBar tool_bar = new JToolBar();
            tool_bar.setFloatable(false);       //to make a tool bar immovable.
            tool_bar.setRollover(true);      //to visually indicate tool bar buttons when the user passes over them with the cursor.
            tool_bar.setBorder(new EtchedBorder());
            //tool_bar.setBorder(new BevelBorder(BevelBorder.RAISED));

            tool_bar.add(btn_import_excel_TB);
            tool_bar.addSeparator();
            tool_bar.add(btn_execute_TB);
            tool_bar.addSeparator();
            tool_bar.add(btn_mail_sttngs_TB);
            tool_bar.addSeparator();
            tool_bar.add(btn_clearData_TB);
            tool_bar.addSeparator();

            this_frame.add(tool_bar, BorderLayout.NORTH);
        } catch (Exception ex) {
            ex.printStackTrace();;
        }
    }

    private static void addFileMenu(JFrame frame) {
        try {
            JMenuBar menubar = new JMenuBar();
            frame.setJMenuBar(menubar);

            // File Menu------------------------------------------------------------------
            JMenu file = new JMenu("File");
            file.setMnemonic('F');
            menubar.add(file);

            // File Sub Menus
            jSubMenu_Import = new JMenuItem("Import excel");
            ImageIcon img_import_excel = new ImageIcon("src/Resources/Images/import_excel_16.png");
            jSubMenu_Import.setIcon(img_import_excel);
            file.add(jSubMenu_Import);

            jSubMenu_Exit = new JMenuItem("Exit");
            file.add(jSubMenu_Exit);

            //Settings Menu------------------------------------------------------------------
            JMenu setting = new JMenu("Settings");
            setting.setMnemonic('S');
            menubar.add(setting);

            //Settings Sub Menu
            jSubMenu_MailSttngs = new JMenuItem("Mail Settings");
            ImageIcon img_mail_sttngs = new ImageIcon("src/Resources/Images/mail_settings_16.png");
            jSubMenu_MailSttngs.setIcon(img_mail_sttngs);
            jSubMenu_MailSttngs.setMnemonic('M');
            setting.add(jSubMenu_MailSttngs);

            //Run Menu------------------------------------------------------------------
            JMenu run = new JMenu("Run");
            run.setMnemonic('R');
            menubar.add(run);

            //Run Sub Menu
            jSubMenu_Execute = new JMenuItem("Execute");
            ImageIcon img_execute_16 = new ImageIcon("src/Resources/Images/execute_16.png");
            jSubMenu_Execute.setIcon(img_execute_16);
            jSubMenu_Execute.setMnemonic('E');
            run.add(jSubMenu_Execute);

            //Run from failed step Menu
//            jSubMenu_Run_FStep = new JMenuItem("Run from failed step");
//            jSubMenu_Run_FStep.setMnemonic('F');
//            run.add(jSubMenu_Run_FStep);
            // Help Menu------------------------------------------------------------------
            JMenu help = new JMenu("Help");
            help.setMnemonic('H');
            menubar.add(help);

            //About
            jSubMenu_About = new JMenuItem("About");
            jSubMenu_About.setMnemonic('A');
            help.add(jSubMenu_About);

        } catch (Exception ex) {
            ex.printStackTrace();;
        }
    }

    public static void addComponentsToPane(Container pane) {
        try {
            int width = 0;
            int height = 0;
            jPanel = new JPanel(new GridBagLayout());
            GridBagConstraints gbConst = new GridBagConstraints();

            //Adding the Table
            initTable();

            //Adding scrollpane around the Table
            jScrollPane1 = new JScrollPane(table, JScrollPane.VERTICAL_SCROLLBAR_ALWAYS, JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
            jScrollPane1.setPreferredSize(new Dimension(Constants.TABLE_WIDTH + 18, Constants.TABLE_HEIGHT));
            jScrollPane1.setMaximumSize(new Dimension(Constants.TABLE_WIDTH + 18, Constants.TABLE_HEIGHT));
            jScrollPane1.setMinimumSize(new Dimension(Constants.TABLE_WIDTH + 18, Constants.TABLE_HEIGHT));

            gbConst.gridy = 0;
            gbConst.gridx = 0;
            gbConst.gridheight = 5;
            gbConst.gridwidth = 2;
            gbConst.insets = new Insets(0, 5, 0, 0);
            gbConst.weightx = 50;
            gbConst.weighty = 5;
            gbConst.anchor = GridBagConstraints.NORTHWEST;
            jPanel.add(jScrollPane1, gbConst);

            //ExcelPath Lbl
            lbl_ExcelPath = new JLabel("Excel Path : ");
            gbConst.gridx = 2;
            gbConst.gridy = 0;
            gbConst.gridheight = 1;
            gbConst.gridwidth = 1;
            gbConst.weightx = 0;
            gbConst.weighty = 0;
            gbConst.insets = new Insets(10, 20, 0, 0);
            gbConst.weightx = 0;
            gbConst.weighty = 0;
            gbConst.ipadx = 0;
            gbConst.anchor = GridBagConstraints.WEST;
            jPanel.add(lbl_ExcelPath, gbConst);

            //ExcelPath Value
            jt_ExcelPath_Val = new JTextField();     //36
            jt_ExcelPath_Val.setPreferredSize(new Dimension(380, 15));
            jt_ExcelPath_Val.setMaximumSize(new Dimension(380, 15));
            jt_ExcelPath_Val.setMinimumSize(new Dimension(380, 15));
            jt_ExcelPath_Val.setBorder(BorderFactory.createMatteBorder(0, 0, 0, 0, Color.BLACK));
            jt_ExcelPath_Val.setEditable(false);
            gbConst.gridx = 3;
            gbConst.gridy = 0;
            gbConst.gridheight = 1;
            gbConst.gridwidth = 1;
            gbConst.weightx = 0;
            gbConst.weighty = 0;
            gbConst.insets = new Insets(10, 0, 0, 0);
            gbConst.weightx = 150;
            gbConst.weighty = 0;
            gbConst.anchor = GridBagConstraints.WEST;
            jPanel.add(jt_ExcelPath_Val, gbConst);
            jt_ExcelPath_Val.setText(gExcelPath);    //"D:\\Mohit\\QTP\\QTP_Excel\\QTP_Excel\\QTP_Excel\\Batch.xls

            //Test Case Folder Label
            lbl_TestCaseFolder = new JLabel("Test Case Folder : ");
            gbConst.gridx = 2;
            gbConst.gridy = 1;
            gbConst.gridheight = 1;
            gbConst.gridwidth = 1;
            gbConst.weightx = 1;
            gbConst.weighty = 1;
            gbConst.insets = new Insets(15, 20, 0, 0);
            gbConst.weightx = 0;
            gbConst.weighty = 0;
            gbConst.anchor = GridBagConstraints.WEST;
            jPanel.add(lbl_TestCaseFolder, gbConst);

            //Test Case Folder Value
            jt_TestCaseFolder_Val = new JTextField();            //C:\\Program Files\\HP\\Unified Functional Testing\\QTP\\QTP_OUTPUT_DIR\\
            jt_TestCaseFolder_Val.setPreferredSize(new Dimension(380, 15));
            jt_TestCaseFolder_Val.setMaximumSize(new Dimension(380, 15));
            jt_TestCaseFolder_Val.setMinimumSize(new Dimension(380, 15));
            jt_TestCaseFolder_Val.setBorder(BorderFactory.createMatteBorder(0, 0, 0, 0, Color.BLACK));
            jt_TestCaseFolder_Val.setEditable(false);
            jt_TestCaseFolder_Val.setText(gTestCaseFolder);
            gbConst.gridx = 3;
            gbConst.gridy = 1;
            gbConst.gridheight = 1;
            gbConst.gridwidth = 1;
            gbConst.weightx = 1;
            gbConst.weighty = 1;
            gbConst.insets = new Insets(15, 0, 0, 0);
            gbConst.weightx = 0;
            gbConst.weighty = 0;
            gbConst.anchor = GridBagConstraints.WEST;
            jPanel.add(jt_TestCaseFolder_Val, gbConst);

            //Display Progress Bar
            progressBar = new JProgressBar();
            progressBar.setIndeterminate(true);
            progressBar.setVisible(false);
            width = 500;
            height = 20;
            progressBar.setPreferredSize(new Dimension(width, height));
            progressBar.setMaximumSize(new Dimension(width, height));
            progressBar.setMinimumSize(new Dimension(width, height));
            gbConst.gridx = 2;
            gbConst.gridy = 2;
            gbConst.gridheight = 1;
            gbConst.gridwidth = 0;
            gbConst.weightx = 50;
            gbConst.weighty = 0;
            gbConst.insets = new Insets(20, 20, 0, 0);
            gbConst.weightx = 0;
            gbConst.weighty = 0;
            gbConst.anchor = GridBagConstraints.WEST;
            jPanel.add(progressBar, gbConst);

            //Logs Label
            lbl_Logs = new JLabel("Logs : ");
            gbConst.gridx = 2;
            gbConst.gridy = 3;
            gbConst.gridheight = 1;
            gbConst.gridwidth = 1;
            gbConst.weightx = 50;
            gbConst.weighty = 0;
            gbConst.insets = new Insets(20, 20, 0, 0);
            gbConst.weightx = 0;
            gbConst.weighty = 0;
            gbConst.anchor = GridBagConstraints.WEST;
            jPanel.add(lbl_Logs, gbConst);

            //Logs Text Field
            jt_Execution_Logs = new JTextArea();//10, 15
            width = 500;
            height = 299;
            jt_Execution_Logs.setPreferredSize(new Dimension(width, height));
            jt_Execution_Logs.setMaximumSize(new Dimension(width, height));
            jt_Execution_Logs.setMinimumSize(new Dimension(width, height));
            jt_Execution_Logs.setText("");
            jt_Execution_Logs.setLineWrap(true);    // If text doesn't fit on a line, jump to the next
            jt_Execution_Logs.setWrapStyleWord(true);   // Makes sure that words stay intact if a line wrap occurs
            jt_Execution_Logs.setEditable(false);

            gbConst.gridx = 2;
            gbConst.gridy = 4;
            gbConst.gridheight = 5;
            gbConst.gridwidth = 0;
            gbConst.insets = new Insets(3, 20, 0, 0);
            gbConst.anchor = GridBagConstraints.NORTHWEST;
            //gbConst.fill = GridBagConstraints.HORIZONTAL;
            JScrollPane scrollbar_Log = new JScrollPane(jt_Execution_Logs, JScrollPane.VERTICAL_SCROLLBAR_ALWAYS, JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
            scrollbar_Log.setPreferredSize(new Dimension(width, height + 5));
            scrollbar_Log.setMaximumSize(new Dimension(width, height + 5));
            scrollbar_Log.setMinimumSize(new Dimension(width, height + 5));
            jPanel.add(scrollbar_Log, gbConst);

            //used to padding for components from the panel's edges
            jPanel.setBorder(new EmptyBorder(10, 10, 10, 10));
            pane.add(jPanel, BorderLayout.CENTER);
        } catch (Exception ex) {
            ex.printStackTrace();;
        }
    }

    public void addComponent_Listeners() {
        try {
            //Menu Components
            jSubMenu_Exit.addActionListener(new CustomButtonListener());
            jSubMenu_Import.addActionListener(new CustomButtonListener());
            jSubMenu_MailSttngs.addActionListener(new CustomButtonListener());
            jSubMenu_Execute.addActionListener(new CustomButtonListener());
            jSubMenu_About.addActionListener(new CustomButtonListener());

            //Frame Components
//            btn_Load.addActionListener(new CustomButtonListener());
//            btn_Execute.addActionListener(new CustomButtonListener());
        } catch (Exception ex) {
            ex.printStackTrace();;
        }
    }

    public static void initTable() {
        try {
            //initializing the Model with Table Column Headers
            headers.clear();
            headers.add(new Boolean(false));
            headers.add("Test Cases");

            model = new CustomTableModel(data, headers);

            //initializing Table
            table = new JTable();
            //table.setAutoCreateRowSorter(true);
            table.setModel(model);
            table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
            table.setEnabled(true);
            table.setRowHeight(25);
            table.setRowMargin(4);
            //the table uses the entire height of the container, even if the table doesn't have enough rows to use the whole vertical space
            table.setFillsViewportHeight(true);

//            int tableWidth = model.getColumnCount() * 150; //not in use
//            int tableHeight = model.getRowCount() * 25; //not in use
            //table.setPreferredSize(new Dimension(Constants.TABLE_WIDTH, Constants.TABLE_HEIGHT));
            table.setPreferredScrollableViewportSize(new Dimension(Constants.TABLE_WIDTH, Constants.TABLE_HEIGHT));
            //setting Table Column Header properties
            setTblColumn_Properties(Constants.TABLE_WIDTH);

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    //Finction to set properties of the JTable column header
    public static void setTblColumn_Properties(int tableWidth) {
        try {
            //Setting width of individual column
            double chkBoxWidth = 0.15 * tableWidth;        //15%
            double testCaseWidth = tableWidth - chkBoxWidth;

            for (int i = 0; i < 2; i++) {
                TableColumn tc1 = table.getColumnModel().getColumn(i);
                if (i == 1) {
                    tc1.setPreferredWidth((int) testCaseWidth); //first column is smaller
                } else {
                    tc1.setPreferredWidth((int) chkBoxWidth);
                }
            }

            //Including the Checkbox object in the column header
            TableColumn tc = table.getColumnModel().getColumn(0);
            tc.setCellEditor(table.getDefaultEditor(Boolean.class));
            tc.setCellRenderer(table.getDefaultRenderer(Boolean.class));
            tc.setHeaderRenderer(new CheckBoxHeader(new CustomItemListener(table), ""));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    //Function to save all the changes made in Excel dialog onto the main form
    public void setImportChanges(String vExcelPath, String vTestCaseFolder) {
        try {
            //setting global variables
            //also replacing all spaces in the path with '$$', as spaces cause problems while passing to vbscript file
            gExcelPath = vExcelPath.replace(" ", "$$");
            gTestCaseFolder = vTestCaseFolder.replace(" ", "$$");

            //setting values in labels
            jt_ExcelPath_Val.setText(vExcelPath);
            jt_TestCaseFolder_Val.setText(vTestCaseFolder);

            //Function to populate the test cases from excel to Jtable
            loadTestCases_inTable(vExcelPath);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    //Function to populate the test cases from excel to Jtable
    public void loadTestCases_inTable(String vExcelPath) {
        Workbook workbook = null;
        try {
            //Sheet sheet = workbook.getSheetAt(0);
            System.out.println("Constants.IS_XLSX = " + Constants.IS_XLSX);
            Sheet sheet = ExcelUtility.getExcelSheet_ByPosition(workbook, vExcelPath, Constants.IS_XLSX, 0);

            if (table != null) {
                data.clear();
                model = (CustomTableModel) table.getModel();

                //remove all existing rows from Table
                int rows = model.getRowCount();
                if (rows > 0) {
                    for (int i = rows - 1; i >= 0; i--) {
                        model.removeRow(i);
                    }
                }
                Row row = null;
                Cell cell = null;

                //Adding new rows to Table
                for (int j = 1; j < sheet.getPhysicalNumberOfRows(); j++) {
                    Vector d = new Vector();

                    row = ExcelUtility.getExcelRow_BySheet(sheet, Constants.IS_XLSX, j);
                    cell = row.getCell((short) 0);
                    d.add(new Boolean(false));
                    d.add(cell.getStringCellValue());
                    model.addRow(d);
                }
            }
            repaint();

            if (workbook != null) {
                workbook.close();
            }
        } catch (Exception e) {
            System.out.println("Exception in function loadTestCases_inTable : " + e.getMessage());
            e.printStackTrace();
        }
    }

    public void load_ButtonClick(ActionEvent ae) {
        try {
            //Component component = (Component) ae.getSource();
            //JFrame mainFrame = (JFrame) SwingUtilities.getRoot(component);

            new LoadExcelDts_Dlg(main_frame, "Excel Import Details", user_prefs);

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    public void execute_ButtonClick(ActionEvent ae) {
        try {
            jt_Execution_Logs.setText("");
            LogUtility.writeToLog(jt_Execution_Logs, "Validating Input data...", true);

            //dont proceed with execution if validation fails
            if (!validateBeforeExecute()) {
                LogUtility.writeToLog(jt_Execution_Logs, "Validation failed", true);
                return;
            }

            LogUtility.writeToLog(jt_Execution_Logs, "Validation complete", true);

            //Disable all the Actions buttons during execution
            toggleActionButtons(true);

            //Resetting the Execution Log File =========================================================
            File log_file = new File(Constants.Execution_Log_Path);
            if (log_file.exists()) {
                //deleting the existing log file
                if (!log_file.delete()) {
                    LogUtility.writeToLog(jt_Execution_Logs, "Unable to delete the previous log file", true);
                    System.out.println("Delete Operation Failed...");
                }
            }
            //creating a new blank file
            log_file = new File(Constants.Execution_Log_Path);

            if (!log_file.createNewFile()) {
                LogUtility.writeToLog(jt_Execution_Logs, "Unable to create the log file", true);
                System.out.println("File not created....");
            }
            log_file = null;
            //===================================================================================================

            LogUtility.writeToLog(jt_Execution_Logs, "Initialising Execution...", true);

            //Executing the code for running test cases in a seperate Java Thread (SwingWorker in Swing)=========
            SwingWorker<Boolean, Void> runTestCase = new ExecuteTestCases(ae);
            runTestCase.execute();
            //===================================================================================================

            //Executing the code for displaying Execution Logs in a seperate Java Thread (SwingWorker in Swing)=========
            SwingWorker<Boolean, Void> displayLogs = new DisplayLogs(ae);
            displayLogs.execute();
            //===================================================================================================

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    public void toggleActionButtons(boolean disableButtons) {//throws Exception {
        if (disableButtons) {
            //disabling buttons
            btn_import_excel_TB.setEnabled(false);
            btn_execute_TB.setEnabled(false);
            btn_mail_sttngs_TB.setEnabled(false);
            btn_clearData_TB.setEnabled(false);

            jSubMenu_Import.setEnabled(false);
            jSubMenu_MailSttngs.setEnabled(false);
            jSubMenu_Execute.setEnabled(false);
            jSubMenu_About.setEnabled(false);

            //displaying the progress bar
            progressBar.setVisible(true);
        } else {
            btn_import_excel_TB.setEnabled(true);
            btn_execute_TB.setEnabled(true);
            btn_mail_sttngs_TB.setEnabled(true);
            btn_clearData_TB.setEnabled(true);

            jSubMenu_Import.setEnabled(true);
            jSubMenu_MailSttngs.setEnabled(true);
            jSubMenu_Execute.setEnabled(true);
            jSubMenu_About.setEnabled(true);

            //hiding the progress bar
            progressBar.setVisible(false);
        }

    }

    public void displayAboutDialog() {
        try {
            new About_Dlg(main_frame, "About");

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    public void clearTestCaseData() {
        try {
            //Clearing JTable Data
            model = (CustomTableModel) table.getModel();
            model.setRowCount(0);

            //Clearing Excel Path
            gExcelPath = "";
            gTestCaseFolder = "";
            jt_ExcelPath_Val.setText("");
            jt_TestCaseFolder_Val.setText("");

            //Clearing Log Data
            jt_Execution_Logs.setText("");
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    //Email Settings Dialog
    public void mail_MenuClick(ActionEvent ae) {
        try {
            new MailSettings_Dlg(main_frame, "Mail Settings", true);
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    private boolean validateBeforeExecute() {
        try {
            //chk if Excel has not been imported by the user
            if (gExcelPath == null || gExcelPath.equals("")) {
                JOptionPane.showMessageDialog(this,
                        "Please import the Test Cases excel file first by clicking the \"Import Test Cases\" button!",
                        "Incomplete Input!!", JOptionPane.WARNING_MESSAGE);
                btn_import_excel_TB.requestFocus();
                return false;
            }

            //chk if Test Case Folder has not been set by the user
            if (gTestCaseFolder == null || gTestCaseFolder.equals("")) {
                JOptionPane.showMessageDialog(this,
                        "Please select the Test Cases folder!",
                        "Incomplete Input!!", JOptionPane.WARNING_MESSAGE);
                btn_import_excel_TB.requestFocus();
                return false;
            }

            //chk if atleast one Test case has been checked by the user
            if (!checkTestCases_Checked()) {
                JOptionPane.showMessageDialog(this,
                        "Please select atleast one test case from the list!",
                        "Incomplete Input!!", JOptionPane.WARNING_MESSAGE);
                btn_import_excel_TB.requestFocus();
                return false;
            }

        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return true;
    }

    private boolean checkTestCases_Checked() {
        boolean isChecked = false;
        try {
            for (int i = 0; i < table.getRowCount(); i++) {
                isChecked = (Boolean) table.getValueAt(i, 0);

                if (isChecked) {
                    break;
                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return isChecked;
    }

    public void exit_MenuClick(ActionEvent ae) {
        System.exit(0);
    }

    //Here 1st parameter is the returnType
    //and 2nd parameter is data to be written in GUI
    private final class ExecuteTestCases extends SwingWorker<Boolean, Void> {

        private ActionEvent mActionEvent;
        private boolean isError = false;
        private StringBuffer selChkBox_Str;
        private java.util.List<String> selTestCaseName_List;    //this will contain the list all the selected test case names
        private java.util.List<String> selTestCaseHTML_List;    //this will contain the list all the HTML content

        public ExecuteTestCases(ActionEvent vActionEvent) {
            this.mActionEvent = vActionEvent;
        }

        @Override
        protected Boolean doInBackground() throws Exception {
            boolean result = false;
            selChkBox_Str = new StringBuffer();
            int colNumIn_Excel = 0;
            selTestCaseName_List = new ArrayList();
            selTestCaseHTML_List = new ArrayList();
            try {
                //get the parent Jframe Object
                Component component = (Component) mActionEvent.getSource();
                JFrame frame = (JFrame) SwingUtilities.getRoot(component);

                //minimize the Framework window when the execute button is clicked
                frame.setState(Frame.ICONIFIED);

                //get the list of selected checkboxes
                for (int i = 0; i < table.getRowCount(); i++) {
                    boolean isChecked = (Boolean) table.getValueAt(i, 0);

                    if (isChecked) {
                        //Column at pos 0 in Jtable would be in position 2 in main excel
                        colNumIn_Excel = i + 2;

                        if (selChkBox_Str.toString().equalsIgnoreCase("")) {
                            selChkBox_Str.append(colNumIn_Excel);
                        } else {
                            selChkBox_Str.append("*" + colNumIn_Excel);
                        }

                        //Store the selected test case names for mailing later on
                        String testCaseName = (String) table.getValueAt(i, 1);
                        selTestCaseName_List.add(testCaseName);
                    }
                }
                //path where the driver script is stored within the netbeans project
                String driverScript_path = Constants.DRIVER_SCRIPT_PATH;

                //execute the MasterScript vbs file
                Process p = Runtime.getRuntime().exec("wscript " + driverScript_path + " " + gExcelPath + "  " + gTestCaseFolder + " " + selChkBox_Str.toString());
                p.waitFor();

                System.out.println("Execution Complete and Exit Value = " + p.exitValue());
                if (p.exitValue() != 0) {
                    isError = true;
                }

                //Send Email : Logic====================================================
                if (!isError) {
                    loadResults_SendMail();
                }
            } catch (Exception ex) {
                ex.printStackTrace();
            }
            return result;
        }

        //This function will be executed after all the Test cases has been executed
        @Override
        protected void done() {
            System.out.println("inside done...");
            //Enable all the Actions buttons during execution
            toggleActionButtons(false);
        }

        public void generateSummaryResult() throws Exception {
            //1. get the list of checked Test cases
            //2. start a loop and load current test case excel sheet
            //3. In the sheet, get last column and chk if its name is Test_Results
            //4. if yes then count passes & fails in that column
            //5. if no then just put zero or - in pass/fail
            //6. Finally write the pass/fail count in the main sheet
            Workbook workbook = null;
            Row row_batch = null;
            Row curr_row = null;

            Sheet batch_sheet = ExcelUtility.getExcelSheet_ByPosition(workbook, gExcelPath, Constants.IS_XLSX, 0);

            //write the headers in batch sheet for no. of pass/fail test cases
            row_batch = ExcelUtility.getExcelRow_BySheet(batch_sheet, Constants.IS_XLSX, 0);
            Cell cell_batch = row_batch.getCell((short) 0);

            testCaseSumm_list = new ArrayList<TestCase>();
            System.out.println("selTestCaseName_List.size() = " + selTestCaseName_List.size());
            for (int i = 0; i < selTestCaseName_List.size(); i++) {
                int fail_count = 0;
                int pass_count = 0;
                String currTestCase = selTestCaseName_List.get(i);

                //load the sheet by name as: TestCaseName = Sheet name in excel
                Sheet sheet = ExcelUtility.getExcelSheet_ByName(workbook, gExcelPath, Constants.IS_XLSX, currTestCase);
                int tot_rows = ExcelUtility.getTotalRows(sheet);
                int tot_cols = ExcelUtility.getTotalColumns(sheet);
                //if there are 4 cols in a sheet then position of last column is 3
                int testResultsCol_Pos = tot_cols - 1;

                //Check if the last column is the results column
                curr_row = ExcelUtility.getExcelRow_BySheet(sheet, Constants.IS_XLSX, 0);
                System.out.println("curr_row.getLastCellNum() = " + curr_row.getLastCellNum());

                Cell cell = null;
                Cell cell_hdr = curr_row.getCell((short) testResultsCol_Pos);

                if (cell_hdr != null && cell_hdr.getStringCellValue().trim().equalsIgnoreCase("Test_Results")) {

                    //loop through all the rows
                    for (int row = 1; row < tot_rows; row++) {
                        curr_row = ExcelUtility.getExcelRow_BySheet(sheet, Constants.IS_XLSX, row);
                        cell = curr_row.getCell((short) testResultsCol_Pos);
                        //Cell cell = sheet.getCell(tot_cols, row);

                        if (cell.getStringCellValue().equalsIgnoreCase("Fail")) {
                            fail_count++;
                        } else if (cell.getStringCellValue().equalsIgnoreCase("Pass")) {
                            pass_count++;
                        }
                    }
                }
                //now write the results in the main batch sheet
                TestCase ts = new TestCase(currTestCase, pass_count, fail_count);
                testCaseSumm_list.add(ts);
            }

            if (workbook != null) {
                workbook.close();
            }
        }

        public void loadResults_SendMail() throws Exception {
            boolean canSend_TCMail = true;           //get these settings from configurations
            boolean canSend_SuiteMail = true;           //get these settings from configurations
            //isError = false;            //remove this later on

            if (!isError) {
                //Loading the Mail Template Saved by user
                XStreamUtil xUtil = new XStreamUtil();
                MailTemplate mailTemplate = (MailTemplate) xUtil.load_data_from_XML(Constants.MAIL_TEMPLATE_FILE);

                if (canSend_TCMail) {
                    try {
                        sendTCMail(mailTemplate);
                    } catch (Exception ex) {
                        LogUtility.writeToLog(jt_Execution_Logs, "Error while sending mail : " + ex.getMessage(), true);
                        LogUtility.writeToLog(jt_Execution_Logs, "Email not sent", true);
                        return;
                    }
                }

                if (canSend_SuiteMail) {
                    try {
                        generateSummaryResult();
                        //Sending mail for Test Suite (summary)---------------------------------------------------------------------------------
                        sendSummaryMail(mailTemplate);
                    } catch (Exception ex) {
                        LogUtility.writeToLog(jt_Execution_Logs, "Error while sending summary mail : " + ex.getMessage(), true);
                        LogUtility.writeToLog(jt_Execution_Logs, "Summary Email not sent", true);
                        return;
                    }
                }
                //---------------------------------------------------------------------------------

                LogUtility.writeToLog(jt_Execution_Logs, "All test cases have been executed", true);
            } else {
                LogUtility.writeToLog(jt_Execution_Logs, "Execution failed due to some error.", true);
            }

            System.out.println("ExecuteTestCases completed");
        }

        private void sendTCMail(MailTemplate mailTemplate) throws Exception {
            //Logic to load data for Sending individual Test case mail
            //1. get total checked test cases
            //2. loop through the test cases and select current test case
            //3. load current test case result from the appropiate sheet in excel
            //4. generate HTML based on the test case result

            LogUtility.writeToLog(jt_Execution_Logs, "Mailing Test Case Results", true);

            if (selTestCaseName_List == null || selTestCaseName_List.size() <= 0) {
                LogUtility.writeToLog(jt_Execution_Logs, "Error: Test Case list is empty", true);
                return;
            }

            Workbook workbook = null;
            try {
                workbook = ExcelUtility.loadExcel_Workbook(gExcelPath, Constants.IS_XLSX);
            } catch (IOException ex) {
                ex.printStackTrace();
            }

            //looping through all the selected test cases
            for (int k = 0; k < selTestCaseName_List.size(); k++) {
                String currTestCase = selTestCaseName_List.get(k);

                //load the sheet by name as: TestCaseName = Sheet name in excel
                Sheet sheet = workbook.getSheet(currTestCase);

                //generate the HTML Table equivalent of the Excel data
                String htmlStr = generateHTML_TestCases(sheet, selTestCaseName_List.get(k));
                System.out.println("htmlStr >> " + htmlStr);

                //maintain the html string of all test cases in a list
                selTestCaseHTML_List.add(htmlStr);
                //System.out.println("htmlStr = " + htmlStr);
            }
            workbook.close();

            for (int list = 0; list < selTestCaseHTML_List.size(); list++) {
                //Important::Uncomment This Later On Mohit
                sendMail(mailTemplate, selTestCaseHTML_List.get(list), false);
            }
        }

        private void sendSummaryMail(MailTemplate mailTemplate) throws Exception {
            String summaryHTML = "";

            if (testCaseSumm_list == null || testCaseSumm_list.size() <= 0) {
                LogUtility.writeToLog(jt_Execution_Logs, "Error: Test Case list is empty", true);
                return;
            }

            summaryHTML = generateHTML_Summary(testCaseSumm_list);
            System.out.println("summaryHTML = " + summaryHTML);

            //Important::Uncomment This Later On Mohit
            sendMail(mailTemplate, summaryHTML, true);
        }

        private String generateHTML_TestCases(Sheet vSheet, String vTestCaseName) {
            StringBuffer html = new StringBuffer();
            StringBuffer rowData = new StringBuffer();
            Cell cell = null;
            String colHeader_css = "font-weight:bold;text-align:center;";
            String row_css = "color:#ffffff; margin:0; font-family:arial; padding:5px 0; font-size:16px; line-height:16px;";
            String css = "";
            int tot_rows = 0;
            int tot_cols = 0;
            Row curr_row = null;
            String bgColor = "";
            try {
                tot_rows = ExcelUtility.getTotalRows(vSheet);
                tot_cols = ExcelUtility.getTotalColumns(vSheet);

                html.append("<span style='font-weight:bold'>Dear Sir, </span><br/>");
                html.append("<span style='margin-left:25px;'>Following are the results of Test case : </span>");
                html.append("<span style='font-weight:bold;'>").append(vTestCaseName).append("</span>");
                html.append("<br/><br/><br/>");

                html.append("<table style=\"width:70%;border:'1px solid #ffffff';border-collapse: collapse;\" border=\"1\"  >");

                //first header row
                html.append("<tr bgcolor=\"#34495e\">");
                html.append("<td style='" + row_css + colHeader_css + "' colspan=" + (tot_cols - 1) + ">Test Data Used</td>");
                html.append("<td style='" + row_css + colHeader_css + "'>Results</td>");
                html.append("</tr>");

                for (int row = 0; row < tot_rows; row++) {
                    if (row == 0) {
                        css = colHeader_css;
                    } else {
                        css = "";
                    }

                    //Clear the row buffer 
                    rowData.delete(0, rowData.length());

                    for (int col = 0; col < tot_cols; col++) {
                        bgColor = "bgcolor = \"#34495e\"";

                        curr_row = ExcelUtility.getExcelRow_BySheet(vSheet, Constants.IS_XLSX, row);
                        cell = curr_row.getCell(col);
                        //cell = vSheet.getCell(col, row);

                        if (col == (tot_cols - 1)) {
                            //Red background for Pass ; Green Bg for failed test cases
                            if (cell == null || cell.getStringCellValue().equalsIgnoreCase("Fail")) {
                                bgColor = "bgcolor = \"#e84c3d\"";
                            } else if (cell.getStringCellValue().equalsIgnoreCase("Pass")) {
                                bgColor = "bgcolor = \"#2fcc71\"";
                            }
                        }

                        rowData.append("<td style='").append(row_css).append(css).append("'>");
                        rowData.append("<p style='margin-left:3px'>");
                        rowData.append(ExcelUtility.getCellValue_Str(cell));
                        rowData.append("</p>");
                        rowData.append("</td>");
                    }
                    html.append("<tr " + bgColor + ">");
                    html.append(rowData);
                    html.append("</tr>");
                }
                html.append("</table>");
            } catch (Exception ex) {
                ex.printStackTrace();
            }
            return html.toString();
        }

        private String generateHTML_Summary(java.util.List vTestCaseSumm_list) {
            StringBuffer html = new StringBuffer();
            int tot_rows = vTestCaseSumm_list.size();
            System.out.println("tot_rows = " + tot_rows);
            int tot_cols = 3;
            TestCase testCaseObj = null;
            try {
                String colHeader_css = "text-align:center;border:1px solid #ccc;background:#B8B8B8;padding:4px;";
                String row_css = "border:1px solid #ccc;padding:4px;";

                html.append("<span style='font-weight:bold'>Dear Sir, </span><br/>");
                html.append("<span style='margin-left:25px;'>Following are the Summary Results of all Test cases: </span>");
                html.append("<br/><br/><br/>");

                html.append("<table style='width:70%;border:1px solid #ccc;border-collapse: collapse;' cellspacing=0 cellpadding=0>");

                //generating Columns
                html.append("<tr style='font-weight:bold;border:1px solid #ccc'>");
                html.append("<td style='width:8%;" + colHeader_css + "'> Sr No </td>");
                html.append("<td style='" + colHeader_css + "'> Test Case </td>");
                html.append("<td style='width:15%;" + colHeader_css + "'> No. of Passes</td>");
                html.append("<td style='width:15%;" + colHeader_css + "'> No. of Fails </td>");
                html.append("</tr>");

                for (int k = 0; k < tot_rows; k++) {
                    testCaseObj = (TestCase) vTestCaseSumm_list.get(k);
                    html.append("<tr>");
                    html.append("<td style='" + row_css + "'>").append(k + 1).append("</td>");    //Sr No
                    html.append("<td style='" + row_css + "'>").append(testCaseObj.getTestCaseName()).append("</td>");     //test Case
                    html.append("<td style='" + row_css + "'>" + testCaseObj.getNo_of_pass() + "</td>");     //No. of Passes
                    html.append("<td style='" + row_css + "'>" + testCaseObj.getNo_of_fail() + "</td>");     //No. of Fails
                    html.append("</tr>");
                }
                html.append("</table>");
            } catch (Exception ex) {
                ex.printStackTrace();
            }
            return html.toString();
        }

        private void sendMail(MailTemplate mailTemplate, String vMailBody, boolean hasAttachment)
                throws AddressException, MessagingException, IOException {
            File attachFile = null;

            if (hasAttachment) {
                attachFile = new File(gExcelPath);
            }

            ConfigUtility configUtil = new ConfigUtility();
            Properties smtpProperties = configUtil.loadProperties();

            MailClient.sendEmail(smtpProperties,
                    mailTemplate.getTo(),
                    mailTemplate.getCc(),
                    mailTemplate.getSubject(),
                    vMailBody,
                    hasAttachment,
                    attachFile);
        }
    }

    //A seperate thread to display Execution logs in a JTextArea
    private final class DisplayLogs extends SwingWorker<Boolean, Void> {

        private ActionEvent mActionEvent;

        public DisplayLogs(ActionEvent vActionEvent) {
            this.mActionEvent = vActionEvent;
        }

        @Override
        protected Boolean doInBackground() throws Exception {
            boolean result = true;
            FileInputStream fstream = null;
            try {
                //Displaying Logs of Test Case Exeuction
                fstream = new FileInputStream(Constants.Execution_Log_Path);

                BufferedReader br = new BufferedReader(new InputStreamReader(fstream));
                String line;

                while (true) {
                    line = br.readLine();
                    if (line == null) {
                        Thread.sleep(500);
                    } else {
                        if (line.indexOf("stop") != -1) {
                            break;
                        }

                        LogUtility.writeToLog(jt_Execution_Logs, line, true);
                        System.out.println(line);
                    }
                }
            } catch (Exception ex) {
                result = false;
                ex.printStackTrace();
            } finally {
                if (fstream != null) {
                    fstream.close();
                }
            }
            return result;
        }

        @Override
        protected void done() {
            try {
                System.out.println("DisplayLogs completed");
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
    }

    public class CustomButtonListener implements ActionListener {

        @Override
        public void actionPerformed(ActionEvent ae) {
            if (ae.getSource() == btn_import_excel_TB || ae.getSource() == jSubMenu_Import) {
                load_ButtonClick(ae);
            } else if (ae.getSource() == btn_execute_TB || ae.getSource() == jSubMenu_Execute) {
                execute_ButtonClick(ae);
            } else if (ae.getSource() == jSubMenu_Run_FStep) {
                //runFromFStep_ButtonClick(ae);
            } else if (ae.getSource() == jSubMenu_Exit) {
                exit_MenuClick(ae);
            } else if (ae.getSource() == jSubMenu_MailSttngs || ae.getSource() == btn_mail_sttngs_TB) {
                mail_MenuClick(ae);
            } else if (ae.getSource() == jSubMenu_About) {
                displayAboutDialog();
            } else if (ae.getSource() == btn_clearData_TB) {
                clearTestCaseData();
            }
        }
    }
}
