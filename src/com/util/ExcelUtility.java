package com.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtility {

    public static Sheet getExcelSheet_ByPosition(Workbook workbook, String vExcelPath, boolean is_xlsx, int vSheetNumber)
            throws Exception {
        try {
            File file = new File(vExcelPath);
            FileInputStream file_stream = new FileInputStream(file);

            if (workbook == null) {
                if (is_xlsx) {
                    workbook = new XSSFWorkbook(file_stream);
                } else {
                    workbook = new HSSFWorkbook(file_stream);
                }
            }

        } catch (IOException ex) {
            ex.printStackTrace();
        }
        return workbook.getSheetAt(vSheetNumber);
    }

    public static Sheet getExcelSheet_ByName(Workbook workbook, String vExcelPath, boolean is_xlsx, String vSheetName)
            throws Exception {
        try {
            File file = new File(vExcelPath);
            FileInputStream file_stream = new FileInputStream(file);

            if (workbook == null) {
                if (is_xlsx) {
                    workbook = new XSSFWorkbook(file_stream);
                } else {
                    workbook = new HSSFWorkbook(file_stream);
                }
            }

        } catch (IOException ex) {
            ex.printStackTrace();
        }
        return workbook.getSheet(vSheetName);
    }

    public static Workbook loadExcel_Workbook(String vExcelPath, boolean is_xlsx)
            throws IOException {
        Workbook workbook = null;
        File file = new File(vExcelPath);
        FileInputStream file_stream = new FileInputStream(file);

        if (is_xlsx) {
            workbook = new XSSFWorkbook(file_stream);
        } else {
            workbook = new HSSFWorkbook(file_stream);
        }

        return workbook;
    }

    public static Row getExcelRow_BySheet(Sheet vSheet, boolean is_xlsx, int vRowNum)
            throws Exception {
        Row row = null;
        if (Constants.IS_XLSX) {
            row = (XSSFRow) vSheet.getRow(vRowNum);
        } else {
            row = (HSSFRow) vSheet.getRow(vRowNum);
        }
        return row;
    }

    public static int getTotalRows(Sheet vSheet) throws Exception {
        return vSheet.getPhysicalNumberOfRows();
    }

    public static int getTotalColumns(Sheet vSheet) throws Exception {
        return vSheet.getRow(0).getPhysicalNumberOfCells();
    }

    public static String getCellValue_Str(Cell cell) throws Exception {
        String _return = "";

        //if a cell is *blank* then it is null
        if (cell == null) {
            _return = "";
        } else {

            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    _return = String.valueOf(cell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    _return = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_BLANK:
                    _return = "";
                    break;
            }
        }
        return _return;
    }
}
