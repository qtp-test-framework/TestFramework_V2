package com.util;

import javax.swing.JTextArea;

public class LogUtility {
    //function used to write the execution logs into the Log Text Area
    public static void writeToLog(JTextArea jt_Execution_Logs, String message, boolean breakToNextLine) {
        if (jt_Execution_Logs != null) {
            if (breakToNextLine) {
                jt_Execution_Logs.append(message + "\n");
            } else {
                jt_Execution_Logs.append(message);

            }
            //JtextArea doesnt display the appended data while other processing is going on without this line
            jt_Execution_Logs.update(jt_Execution_Logs.getGraphics());
            
            //Scroll to the bottom of text area
            jt_Execution_Logs.setCaretPosition(jt_Execution_Logs.getDocument().getLength());
        }
    }
}
