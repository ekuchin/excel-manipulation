package ru.ekuchin;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class Excel {
    private Workbook workbook;
    private Sheet sheet;

    public void setSheet(int i) {
        sheet = workbook.getSheetAt(i);
    }

    public Excel(String filename) throws IOException {
        workbook = new XSSFWorkbook(new FileInputStream(filename));
        sheet = workbook.getSheetAt(0);
    }

    public Cell getCell(int row, int cell){
        return sheet.getRow(row).getCell(cell);
    }

    public int getRowCount(){
        return sheet.getPhysicalNumberOfRows();
    }

    public String getCellValueString(Cell cell){
        switch (cell.getCellType()) {
            case NUMERIC:
                //return String.format("%0$,.0f",cell.getNumericCellValue());
                return ""+(int)cell.getNumericCellValue();
            case BOOLEAN:
                return Boolean.toString(cell.getBooleanCellValue());
            default:
                return cell.getStringCellValue();
        }
    }
}