import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import sun.org.mozilla.javascript.internal.regexp.RegExpImpl;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created with IntelliJ IDEA13.0
 * User:cdliu
 * Date:14-6-3
 * Time:上午10:36
 * Version:0.1
 */
public class AttendanceConverter {

    public void createNewXLSFile(){
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
    public void createNewXLSXFile(){
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

    public void testAlignment(){
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

    public void setBorder(){
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
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                fileOut.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static Workbook readExcelFile(File file) throws IOException {
        String fileName = file.getName();
        String fileType = fileName.lastIndexOf(".") == -1 ? "" : fileName.substring(fileName.lastIndexOf(".") + 1);
        if (fileType.equals("xls") || fileType.equals("XLS")) {
            return new HSSFWorkbook(new FileInputStream(file));
        } else if (fileType.equals("xlsx") || fileType.equals("XLSX")){
            return new XSSFWorkbook(new FileInputStream(file));
        } else {
            throw new IOException("不支持的文件类型!");
        }
    }
    /**
     * 用于测试的main方法
     * @param args
     */
    public static void main(String[] args){
        AttendanceConverter attendanceConverter = new AttendanceConverter();
        //attendanceConverter.createNewXLSFile();
        //attendanceConverter.createNewXLSXFile();
        //attendanceConverter.testAlignment();
        try {
            Workbook wb = readExcelFile(new File("E:\\AttendanceConverter\\inputFile\\01原始记录表-14年05月.xls"));
            Sheet sheet = wb.getSheetAt(0);
            Row row0 = sheet.getRow(0);
            Row row1 = sheet.getRow(1);
            Row row2 = sheet.getRow(2);
            Row row3 = sheet.getRow(3);

            Cell cell0 = row0.getCell(0);
            Cell cell1 = row1.getCell(0);

            String title = row0.getCell(0).getStringCellValue();
            String header = row1.getCell(0).getStringCellValue();

            String month = getMonth(title,1);
            System.out.println("第一行标题为:"+title);
            System.out.println("从第一行中提取的月份:"+month);

            String[] info = getInfo(header);
            System.out.println(info[3]);

            Cell cell;
            Map<Integer,String> map = new HashMap<Integer,String>();
            for(int i = 0; i < row2.getLastCellNum(); i++){
                cell = row3.getCell(i);
                map.put(i+1,cell.getStringCellValue());
            }


            String date = (String)map.get(1);

            System.out.println(attendanceTime(date)[0]);


        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static String getMonth(String str, int needIndex){
        String[] tempStr = str.split("\\D+");   //把数字剥离出来
        if (needIndex < tempStr.length){
            return tempStr[needIndex];  //第二组是月份
        }else {
            throw new ArrayIndexOutOfBoundsException("给的index超过String数组长度!");
        }
    }

    public static String[] getInfo(String str){
        return str.split("(:| )+");
    }

    public static String[] attendanceTime(String time){
        return time.split(" ");
    }

    private static void createCell(Row row, int column, CellStyle cellStyle, String name, String id, HashMap<Integer,String> hashMap) {
        Cell cell = row.createCell(column);
        setCellStyle(cell, cellStyle);
        cell.setCellValue("位置");
    }

    private static void createRow(Sheet sheet,int rowNo){
        Row row = sheet.createRow(rowNo);

    }

    private static void createSheet(){

    }

    /**
     * 设置单元格样式，主要是边框与居中
     * @param cell
     * @param cellStyle
     */
    public static void setCellStyle(Cell cell, CellStyle cellStyle){
        cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
        cellStyle.setBorderRight(CellStyle.BORDER_THIN);
        cellStyle.setBorderTop(CellStyle.BORDER_THIN);
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cell.setCellStyle(cellStyle);
    }
}
