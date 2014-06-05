import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

/**
 * Created with IntelliJ IDEA13.0
 * User:cdliu
 * Date:14-6-5
 * Time:下午3:43
 * Version:0.1
 */
public class ApachePOIofficialDemo {
    public static void createNewXLSFile(){
        Workbook wb1 = new HSSFWorkbook();
        FileOutputStream fileOut1 = null;
        Sheet sheet1 = wb1.createSheet("new sheet");
        Sheet sheet2 = wb1.createSheet("second sheet");
        CreationHelper createHelper = wb1.getCreationHelper();
        Row row = sheet1.createRow((short)0);
        Cell cell = row.createCell(0);
        cell.setCellValue(1);
        row.createCell(1).setCellValue(1.2);
        row.createCell(2).setCellValue(createHelper.createRichTextString("This is a string"));
        row.createCell(3).setCellValue(true);
        try {
            fileOut1 = new FileOutputStream("E:\\AttendanceConverter\\inputFile\\workbook.xls");
            wb1.write(fileOut1);
            fileOut1.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    public static void createNewXLSXFile(){
        Workbook wb2 = new XSSFWorkbook();
        FileOutputStream fileOut2 = null;
        CreationHelper creationHelper = wb2.getCreationHelper();
        Sheet sheet = wb2.createSheet("Sheet1");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue(new Date());

        CellStyle cellStyle = wb2.createCellStyle();
        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy/m/d"));
        cell = row.createCell(1);
        cell.setCellValue(new Date());
        cell.setCellStyle(cellStyle);

        cell = row.createCell(2);
        cell.setCellValue(Calendar.getInstance());
        cell.setCellStyle(cellStyle);

        try {
            fileOut2 = new FileOutputStream("E:\\AttendanceConverter\\inputFile\\workbook.xlsx");
            wb2.write(fileOut2);
            fileOut2.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void testAlignment(){
        Workbook wb = new XSSFWorkbook(); //or new HSSFWorkbook();

        Sheet sheet = wb.createSheet();
        Row row = sheet.createRow((short)2);
        row.setHeightInPoints(30);

        createCell(wb, row, (short) 0, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_BOTTOM);
        createCell(wb, row, (short) 1, CellStyle.ALIGN_CENTER_SELECTION, CellStyle.VERTICAL_BOTTOM);
        createCell(wb, row, (short) 2, CellStyle.ALIGN_FILL, CellStyle.VERTICAL_CENTER);
        createCell(wb, row, (short) 3, CellStyle.ALIGN_GENERAL, CellStyle.VERTICAL_CENTER);
        createCell(wb, row, (short) 4, CellStyle.ALIGN_JUSTIFY, CellStyle.VERTICAL_JUSTIFY);
        createCell(wb, row, (short) 5, CellStyle.ALIGN_LEFT, CellStyle.VERTICAL_TOP);
        createCell(wb, row, (short) 6, CellStyle.ALIGN_RIGHT, CellStyle.VERTICAL_TOP);

        // Write the output to a file
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream("E:\\AttendanceConverter\\inputFile\\workbook.xlsx");
            wb.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
    /**
     * Creates a cell and aligns it a certain way.
     *
     * @param wb     the workbook
     * @param row    the row to create the cell in
     * @param column the column number to create the cell in
     * @param halign the horizontal alignment for the cell.
     */
    private static void createCell(Workbook wb, Row row, short column, short halign, short valign) {
        Cell cell = row.createCell(column);
        cell.setCellValue("位置");
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(halign);
        cellStyle.setVerticalAlignment(valign);
        cell.setCellStyle(cellStyle);
    }

    public static void setBorder(){
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        // Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet.createRow(2);

        // Create a cell and put a value in it.
        Cell cell = row.createCell(1);
        cell.setCellValue(4);

        // Style the cell with borders all around.
        CellStyle style = wb.createCellStyle();
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.GREEN.getIndex());
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.BLUE.getIndex());
        style.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cell.setCellStyle(style);

        // Write the output to a file
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream("E:\\AttendanceConverter\\inputFile\\workbook.xlsx");
            wb.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void mergeCell(){
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        Row row = sheet.createRow((short) 1);
        Cell cell = row.createCell((short) 1);
        cell.setCellValue("This is a test of merging");

        sheet.addMergedRegion(new CellRangeAddress(
                1, //first row (0-based)
                1, //last row  (0-based)
                1, //first column (0-based)
                2  //last column  (0-based)
        ));

        // Write the output to a file
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream("E:\\AttendanceFile\\workbook.xls");
            wb.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e){
            e.printStackTrace();
        }

    }

    public static void main(String[] args){
        //createNewXLSFile();
        //createNewXLSXFile();
        //testAlignment();
        //setBorder();
        mergeCell();
    }
}
