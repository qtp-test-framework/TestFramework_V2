
package com.main;

import java.util.Vector;
import javax.swing.table.DefaultTableModel;


public class CustomTableModel extends DefaultTableModel {

    public CustomTableModel(Vector rowData, Vector columnNames) {
        super(rowData, columnNames);
    }
    
    @Override
    public Class<?> getColumnClass(int columnIndex) {
        Class clazz = String.class;
        switch (columnIndex) {
            case 0:
                clazz = Boolean.class;
                break;
            case 2:
                clazz = String.class;
                break;
        }
        return clazz;
    }
}
