/*
 * >x      ^Y
 * Insets(top, left, bottom, right);
 * 
 * To add border : lbl_ExcelPath.setBorder(BorderFactory.createMatteBorder(borderWidth, 0, borderWidth, borderWidth, Color.BLACK));
 */
package com.main;

import com.util.CheckBoxHeader;
import com.util.Constants;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.Vector;
import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.table.TableColumn;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class MainWindow extends JFrame {

    // Variables declaration - do not modify                     
    private JButton btn_Load;
    private JButton btn_Execute;
    private JButton jButton7;
    private JCheckBox jCheckBox1;
    private JFrame jFrame1;
    private JLabel lbl_ExcelPath;
    private JLabel lbl_TestCaseFolder;
    private JTextField jt_ExcelPath_Val;
    private JTextField jt_TestCaseFolder_Val;
    private JLabel lbl_Logs;
    private JPanel jPanel;
    private JScrollPane jScrollPane1;
    private JScrollPane jScrollPane2;
    private JScrollPane jScrollPane4;
    private JTable table;
    private JTextArea jt_Execution_Logs;
    // End of variables declaration
    static CustomTableModel model = null;
    static Vector headers = new Vector();
    static Vector data = new Vector();
    String gExcelPath = "";
    String gTestCaseFolder = "";

    public MainWindow(String title) {
        super(title);

        this.setLocationRelativeTo(null);
        this.setPreferredSize(new Dimension(Constants.MainWindow_WIDTH, Constants.MainWindow_HEIGHT));
        this.setMinimumSize(new Dimension(Constants.MainWindow_WIDTH, Constants.MainWindow_HEIGHT));
        this.setMaximumSize(new Dimension(Constants.MainWindow_WIDTH, Constants.MainWindow_HEIGHT));
//        this.setSize(700, 500);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setResizable(false);

        //code to Load the frame in the center of the screen===============
        Toolkit toolkit = Toolkit.getDefaultToolkit();
        Dimension dimn = toolkit.getScreenSize();
        int xPos = (dimn.width / 2) - (this.getWidth() / 2);
        int ypos = (dimn.height / 2) - (this.getHeight() / 2);
        this.setLocation(xPos, ypos);
        //============================================================

        jPanel = new JPanel(new GridBagLayout());
        GridBagConstraints gbConst = new GridBagConstraints();

        //Adding the Table
        initTable();

        //Adding scrollpane around the Table
        jScrollPane1 = new JScrollPane(table, JScrollPane.VERTICAL_SCROLLBAR_ALWAYS, JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        jScrollPane1.setPreferredSize(new Dimension(Constants.TABLE_WIDTH + 18, Constants.TABLE_HEIGHT));
        jScrollPane1.setMaximumSize(new Dimension(Constants.TABLE_WIDTH + 18, Constants.TABLE_HEIGHT));
        jScrollPane1.setMinimumSize(new Dimension(Constants.TABLE_WIDTH + 18, Constants.TABLE_HEIGHT));

        //jScrollPane1 = new JScrollPane(table);
        //jScrollPane1.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        //jScrollPane1.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        //jScrollPane1.setViewportView(table);
        gbConst.gridy = 0;
        gbConst.gridx = 0;
        gbConst.gridheight = 3;
        gbConst.gridwidth = 2;
        gbConst.insets = new Insets(10, 10, 0, 0);
        gbConst.weightx = 100;
        gbConst.weighty = 100;
        gbConst.anchor = GridBagConstraints.WEST;

        //jScrollPane1.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);

        //addComp(JPanel thePanel, JComponent comp, int xPos, int yPos, int compWidth, int compHeight, int place, int stretch)
        //addComp(jPanel, jScrollPane1, 0, 0, 3, 5, GridBagConstraints.WEST, GridBagConstraints.NONE);
        jPanel.add(jScrollPane1, gbConst);

        //Load Button
        btn_Load = new JButton("LOAD");
        gbConst.gridx = 0;
        gbConst.gridy = 5;
        gbConst.gridheight = 1;
        gbConst.gridwidth = 1;
        gbConst.insets = new Insets(0, 35, 0, 0);
        gbConst.weightx = 0;
        gbConst.weighty = 0;
        gbConst.ipady = 5;
        gbConst.ipadx = 7;
        gbConst.anchor = GridBagConstraints.WEST;
        jPanel.add(btn_Load, gbConst);
        btn_Load.addActionListener(new CustomButtonListener());
        //addComp(jPanel, btn_Load, 0, 5, 1, 1, GridBagConstraints.WEST, GridBagConstraints.NONE);

        //Execute Button
        btn_Execute = new JButton("EXECUTE");
        gbConst.gridx = 0;
        gbConst.gridy = 5;
        gbConst.gridheight = 1;
        gbConst.gridwidth = 1;
        gbConst.insets = new Insets(0, 120, 0, 0);
        gbConst.weightx = 0;
        gbConst.weighty = 0;
        gbConst.anchor = GridBagConstraints.EAST;
        //addComp(jPanel, btn_Execute, 1, 5, 1, 1, GridBagConstraints.WEST, GridBagConstraints.NONE);
        jPanel.add(btn_Execute, gbConst);
        btn_Execute.addActionListener(new CustomButtonListener());

        //ExcelPath Lbl
        lbl_ExcelPath = new JLabel("Excel Path : ");
        gbConst.gridx = 2;
        gbConst.gridy = 0;
        gbConst.gridheight = 1;
        gbConst.gridwidth = 1;
        gbConst.weightx = 0;
        gbConst.weighty = 0;
        gbConst.insets = new Insets(40, 20, 0, 0);
        gbConst.weightx = 0;
        gbConst.weighty = 0;
        gbConst.anchor = GridBagConstraints.WEST;
        //addComp(jPanel, lbl_ExcelPath, 3, 0, 1, 1, GridBagConstraints.EAST, GridBagConstraints.NONE);
        jPanel.add(lbl_ExcelPath, gbConst);

        //ExcelPath Value
        jt_ExcelPath_Val = new JTextField();     //36
        jt_ExcelPath_Val.setPreferredSize(new Dimension(280, 10));
        jt_ExcelPath_Val.setMaximumSize(new Dimension(280, 10));
        jt_ExcelPath_Val.setMinimumSize(new Dimension(280, 10));
        jt_ExcelPath_Val.setBorder(BorderFactory.createMatteBorder(0, 0, 1, 0, Color.BLACK));
        jt_ExcelPath_Val.setEditable(false);
        gbConst.gridx = 3;
        gbConst.gridy = 0;
        gbConst.gridheight = 1;
        gbConst.gridwidth = 1;
        gbConst.weightx = 0;
        gbConst.weighty = 0;
        gbConst.insets = new Insets(40, 0, 0, 0);
        gbConst.weightx = 150;
        gbConst.weighty = 0;
        gbConst.anchor = GridBagConstraints.WEST;
        //addComp(jPanel, lbl_ExcelPath, 3, 0, 1, 1, GridBagConstraints.EAST, GridBagConstraints.NONE);
        jPanel.add(jt_ExcelPath_Val, gbConst);
        jt_ExcelPath_Val.setText("");    //"D:\\Mohit\\QTP\\QTP_Excel\\QTP_Excel\\QTP_Excel\\Batch.xls

        //Test Case Folder Label
        lbl_TestCaseFolder = new JLabel("Test Case Folder : ");
        gbConst.gridx = 2;
        gbConst.gridy = 1;
        gbConst.gridheight = 1;
        gbConst.gridwidth = 1;
        gbConst.weightx = 1;
        gbConst.weighty = 1;
        gbConst.insets = new Insets(20, 20, 0, 0);
        gbConst.weightx = 0;
        gbConst.weighty = 0;
        gbConst.anchor = GridBagConstraints.WEST;
        jPanel.add(lbl_TestCaseFolder, gbConst);

        //Test Case Folder Value
        jt_TestCaseFolder_Val = new JTextField();            //C:\\Program Files\\HP\\Unified Functional Testing\\QTP\\QTP_OUTPUT_DIR\\
        jt_TestCaseFolder_Val.setPreferredSize(new Dimension(280, 10));
        jt_TestCaseFolder_Val.setMaximumSize(new Dimension(280, 10));
        jt_TestCaseFolder_Val.setMinimumSize(new Dimension(280, 10));
        jt_TestCaseFolder_Val.setBorder(BorderFactory.createMatteBorder(0, 0, 1, 0, Color.BLACK));
        jt_TestCaseFolder_Val.setEditable(false);
        gbConst.gridx = 3;
        gbConst.gridy = 1;
        gbConst.gridheight = 1;
        gbConst.gridwidth = 1;
        gbConst.weightx = 1;
        gbConst.weighty = 1;
        gbConst.insets = new Insets(20, 0, 0, 0);
        gbConst.weightx = 0;
        gbConst.weighty = 0;
        gbConst.anchor = GridBagConstraints.WEST;
        jPanel.add(jt_TestCaseFolder_Val, gbConst);

        //Logs Label
        lbl_Logs = new JLabel("Logs : ");
        gbConst.gridx = 0;
        gbConst.gridy = 6;
        gbConst.gridheight = 1;
        gbConst.gridwidth = 1;
        gbConst.weightx = 0;
        gbConst.weighty = 0;
        gbConst.insets = new Insets(15, 15, 0, 0);
        gbConst.weightx = 0;
        gbConst.weighty = 0;
        gbConst.anchor = GridBagConstraints.WEST;
        jPanel.add(lbl_Logs, gbConst);

        //Logs Text Field
        jt_Execution_Logs = new JTextArea(4,5);
        //jt_Execution_Logs.setPreferredSize(new Dimension(104, 76));
        jt_Execution_Logs.setMaximumSize(new Dimension(104, 76));
        jt_Execution_Logs.setMinimumSize(new Dimension(104, 76));

        jt_Execution_Logs.setText("");
        jt_Execution_Logs.setLineWrap(true);    // If text doesn't fit on a line, jump to the next
        jt_Execution_Logs.setWrapStyleWord(true);   // Makes sure that words stay intact if a line wrap occurs
        jt_Execution_Logs.setEditable(false);
        

        gbConst.gridx = 0;
        gbConst.gridy = 7;
        gbConst.gridheight = 0;
        gbConst.gridwidth = 0;
        gbConst.insets = new Insets(3, 10, 8, 0);
        gbConst.anchor = GridBagConstraints.WEST;
        gbConst.fill = GridBagConstraints.HORIZONTAL;
        JScrollPane scrollbar_Log = new JScrollPane(jt_Execution_Logs, JScrollPane.VERTICAL_SCROLLBAR_ALWAYS, JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        scrollbar_Log.setPreferredSize(jt_Execution_Logs.getPreferredSize());
        scrollbar_Log.setMaximumSize(jt_Execution_Logs.getPreferredSize());
        scrollbar_Log.setMinimumSize(jt_Execution_Logs.getPreferredSize());
        jPanel.add(scrollbar_Log, gbConst);

        jPanel.setBorder(new EmptyBorder(10, 10, 10, 10));
        this.add(jPanel);
        this.pack();
        this.setVisible(true);

    }

    public static void main(String args[]) {
        try {
            //Get System (Microsoft Windows) Look and Feel
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());

        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(MainWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(MainWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(MainWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(MainWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }

        /*
         * Create and display the form
         */
        EventQueue.invokeLater(new Runnable() {

            public void run() {
                new MainWindow("QTP Automation Framework");
            }
        });
    }

    public void initTable() {
        try {
            //initializing the Model with Table Column Headers
            headers.clear();
            headers.add(new Boolean(false));
            headers.add("Test Cases");
            //remove this later on
//            data.clear();
//            Vector d = new Vector();
//            d.add(new Boolean(false));
//            d.add("LoadFromExcel 1");
//            data.add(d);

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

            int tableWidth = model.getColumnCount() * 150; //not in use
            int tableHeight = model.getRowCount() * 25; //not in use

            //table.setPreferredSize(new Dimension(Constants.TABLE_WIDTH, Constants.TABLE_HEIGHT));

            table.setPreferredScrollableViewportSize(new Dimension(Constants.TABLE_WIDTH, Constants.TABLE_HEIGHT));

            //setting Table Column Header properties
            setTblColumn_Properties(Constants.TABLE_WIDTH);

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    private void addComp(JPanel thePanel, JComponent comp, int xPos, int yPos, int compWidth, int compHeight, int place, int stretch) {

        GridBagConstraints gridConstraints = new GridBagConstraints();

        gridConstraints.gridx = xPos;
        gridConstraints.gridy = yPos;
        gridConstraints.gridwidth = compWidth;
        gridConstraints.gridheight = compHeight;
        gridConstraints.weightx = 100;
        gridConstraints.weighty = 100;
        gridConstraints.insets = new Insets(5, 5, 5, 5);
        gridConstraints.anchor = place;
        gridConstraints.fill = stretch;

        thePanel.add(comp, gridConstraints);

    }

    public void setTblColumn_Properties(int tableWidth) {
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
            gExcelPath = vExcelPath;
            gTestCaseFolder = vTestCaseFolder;

            //setting values in labels
            jt_ExcelPath_Val.setText(vExcelPath);
            jt_TestCaseFolder_Val.setText(gTestCaseFolder);

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
            try {
                File file = new File(vExcelPath);
                workbook = Workbook.getWorkbook(file);
            } catch (IOException ex) {
                ex.printStackTrace();
            }

            Sheet sheet = workbook.getSheet(0);

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

                //Adding new rows to Table
                for (int j = 1; j < sheet.getRows(); j++) {
                    Vector d = new Vector();
                    Cell cell = sheet.getCell(0, j);
                    d.add(new Boolean(false));
                    d.add(cell.getContents());
                    model.addRow(d);
                }
            }
            repaint();

        } catch (BiffException e) {
            System.out.println("BiffException in function loadTestCases_inTable : " + e.getMessage());
            e.printStackTrace();
        } catch (Exception e) {
            System.out.println("Exception in function loadTestCases_inTable : " + e.getMessage());
            e.printStackTrace();
        }
    }

    public void load_ButtonClick(ActionEvent ae) {
        try {
            Component component = (Component) ae.getSource();
            JFrame mainFrame = (JFrame) SwingUtilities.getRoot(component);

            LoadExcelDts_Dialog excel_Dlg = new LoadExcelDts_Dialog(mainFrame, "Excel Import Details");
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    public void execute_ButtonClick(ActionEvent ae) {
        try {
            //Resetting the Execution Log File =========================================================
            File log_file = new File(Constants.Execution_Log_Path);
            if (log_file.exists()) {
                //deleting the existing log file
                if(!log_file.delete()){
                    System.out.println("Delete Operation Failed...");
                }
             }
            //creating a new blank file
            log_file = new File(Constants.Execution_Log_Path);
            
            if(! log_file.createNewFile()){
                System.out.println("File not created....");
            }
            log_file = null;
            //===================================================================================================
            
            
            //Executing the code for running test cases in a seperate Java Thread (SwingWorker in Swing)=========
            SwingWorker<Boolean, Void> runTestCase = new ExecuteTestCases(ae);
            runTestCase.execute();
            //===================================================================================================

            

            //Displaying Logs of Test Case Exeuction
            FileInputStream fstream = new FileInputStream(
                    "Execution_Logs/Execution_Logs.txt");

            BufferedReader br = new BufferedReader(new InputStreamReader(fstream));
            String line;
            jt_Execution_Logs.setText("");

            while (true) {

                line = br.readLine();
                if (line == null) {
                    Thread.sleep(500);
                } else {
                    jt_Execution_Logs.append(line + "\n");
                    System.out.println(line);
                    
                    //JtextArea does not display the appended data while other processing is going on without this line
                    jt_Execution_Logs.update(jt_Execution_Logs.getGraphics());
                    if (line.indexOf("Ending") != -1) {
                        break;
                    }
                }

            }
            fstream.close();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    //Here 1st parameter is the returnType
    //and 2nd parameter is data to be written in GUI
    private final class ExecuteTestCases extends SwingWorker<Boolean, Void> {

        private ActionEvent mActionEvent;

        public ExecuteTestCases(ActionEvent vActionEvent) {
            this.mActionEvent = vActionEvent;
        }

        @Override
        protected Boolean doInBackground() throws Exception {
            boolean result = false;
            StringBuffer selChkBox_Str = new StringBuffer();
            int colNumIn_Excel = 0;
            try {
                //get the parent Jframe Object
                Component component = (Component) mActionEvent.getSource();
                JFrame frame = (JFrame) SwingUtilities.getRoot(component);

                //minimize the Framework window when the execute button is clicked
                //frame.setState(Frame.ICONIFIED);

                //get the list of selected checkboxes
                for (int i = 0; i < table.getRowCount(); i++) {
                    boolean isChecked = (Boolean) table.getValueAt(i, 0);

                    if (isChecked) {
                        //Column at pos 0 in Jtable would be in position 2 in main excel
                        colNumIn_Excel = i + 2;

                        System.out.println("colNumIn_Excel = " + colNumIn_Excel);
                        if (selChkBox_Str.toString().equalsIgnoreCase("")) {
                            selChkBox_Str.append(colNumIn_Excel);
                        } else {
                            selChkBox_Str.append("*" + colNumIn_Excel);
                        }
                    }
                }
                System.out.println("selChkBox_Str = " + selChkBox_Str);
                //path where the driver script is stored within the netbeans project
                String driverScript_path = Constants.DRIVER_SCRIPT_PATH;

                //System.out.println("script = "+"wscript " + driverScript_path + " " + gExcelPath + "  " + gTestCaseFolder + " " + selChkBox_Str.toString());

                //execute the MasterScript vbs file
                Process p = Runtime.getRuntime().exec("wscript " + driverScript_path + " " + gExcelPath + "  " + gTestCaseFolder + " " + selChkBox_Str.toString());
                //Process p = Runtime.getRuntime().exec("wscript  ../TestFramework_1/src/Resources/TestDriverScript.vbs "+gExcelPath+", "+selChkBox_Str);

                p.waitFor();

                System.out.println("Execution Complete and Exit Value = " + p.exitValue());
            } catch (Exception ex) {
                ex.printStackTrace();
            }
            return result;
        }

        @Override
        protected void done() {
            try {
                System.out.println("Inside Done.....");
                jt_Execution_Logs.append("All test cases have been executed..." + "\n");
                jt_Execution_Logs.update(jt_Execution_Logs.getGraphics());
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
    }

    public class CustomButtonListener implements ActionListener {

        @Override
        public void actionPerformed(ActionEvent ae) {
            if (ae.getSource() == btn_Load) {
                load_ButtonClick(ae);
            } else if (ae.getSource() == btn_Execute) {
                execute_ButtonClick(ae);
            }
        }
    }
}