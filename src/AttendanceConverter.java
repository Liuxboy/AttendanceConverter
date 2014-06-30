import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
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

    /**
     * 读取源Excel
     * @param file 源文件
     * @return Workbook
     * @throws IOException
     */
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
     * 获取考勤数据
     * @param filePath 文件路径
     * @return InitAttendanceDataBean
     */
    public static InitAttendanceDataBean getAttendanceData(String filePath){
        String year = null;
        String month = null;
        String[] info = null;
        Map<Integer,String> map = new HashMap<Integer,String>();

        Workbook wb;
        Sheet sheet;
        Row row0;
        Row row1;
        Row row2;
        Row row3;
        Cell cell0;
        Cell cell1;

        try {
            wb = readExcelFile(new File(filePath));
            if (wb != null) {
                sheet = wb.getSheetAt(0);
                row0 = sheet.getRow(0);         //取第一行Row0（程序员你懂的，从0开始），这一行是title
                row1 = sheet.getRow(1);         //取第二行Row1，个人信息
                row2 = sheet.getRow(2);         //取第三行Row2，日期
                row3 = sheet.getRow(3);         //取第四行Row3，时间

                cell0 = row0.getCell(0);        //取title，提取月份
                cell1 = row1.getCell(0);        //取个人信息，提取姓名与工号

                year = getYearAndMonth(cell0.getStringCellValue(),0);
                month = getYearAndMonth(cell0.getStringCellValue(),1);
                info = getInfo(cell1.getStringCellValue());

                Cell cell;
                for(int i = 0; i < row2.getLastCellNum(); i++){ //遍历第二行与第三行Cell，将日期与考勤时间塞进一个HashMap中
                    cell = row3.getCell(i);
                    map.put(i+1,cell.getStringCellValue());
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        InitAttendanceDataBean initAttendanceDataBean = new InitAttendanceDataBean();
        initAttendanceDataBean.setYear(year);
        initAttendanceDataBean.setMonth(month);
        initAttendanceDataBean.setInfo(info);
        initAttendanceDataBean.setAttendanceMap(map);
        return initAttendanceDataBean;
    }

    /**
     * 创建转换好的Excel
     * @param rowCursor 行标
     * @param initAttendanceDataBean 数据Bean
     */
    public static void createNewExcel(int rowCursor, InitAttendanceDataBean initAttendanceDataBean, String filePath){
        Map<Integer,String> map = initAttendanceDataBean.getAttendanceMap();

        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet();
        sheet.setColumnWidth(0,3766);
        for(int k = 0; k < map.size()+3; k++){
            sheet.setColumnWidth(k,3800);
        }
        sheet.setColumnWidth(2,4700);

        CreationHelper creationHelper = wb.getCreationHelper();
        CellStyle cellStyleOfTime = wb.createCellStyle();
        cellStyleOfTime.setDataFormat(creationHelper.createDataFormat().getFormat("h:mm:ss"));

        CellStyle styleOfBGC = wb.createCellStyle();
        styleOfBGC.setFillForegroundColor(IndexedColors.RED.getIndex());
        styleOfBGC.setFillPattern(CellStyle.SOLID_FOREGROUND);

        Row row0 = sheet.createRow(rowCursor);       //第一行写header
        row0.createCell(0).setCellValue("姓名");
        row0.createCell(1).setCellValue("工号");
        row0.createCell(2).setCellValue("标准日期");
        row0.createCell(3).setCellValue("标准打卡时间");
        row0.createCell(4).setCellValue("备注");
        for (Cell cell0 : row0){
            cell0.setCellStyle(styleOfBGC);
        }
        Row row;
        Cell cell;
        String[] attendanceTime;
        String name = initAttendanceDataBean.getInfo()[3];
        String id = initAttendanceDataBean.getInfo()[1];
        String year = initAttendanceDataBean.getYear();
        String month = initAttendanceDataBean.getMonth();
        String time;

        //设置打卡记录
        for (int i = 0; i < map.size(); i++){
            time = map.get(i+1);
            if (time==null || time.equals("") || time.equals(" ")){
                continue;       //碰到空的，则跳过
            }else {
                attendanceTime = attendanceTime(map.get(i+1));
                for(int j = 0; j < 2; j++){
                    row = sheet.createRow(++rowCursor);                      //新建一行，自增
                    row.createCell(0).setCellValue(name);                   //姓名
                    row.createCell(1).setCellValue(id);                     //工号

                    cell = row.createCell(2);
                    cell.setCellValue(year +"/"+ month +"/"+ (i+1));        //标准日期
                    cell.setCellStyle(cellStyleOfTime);

                    row.createCell(3).setCellValue(attendanceTime[j]);      //标准打卡时间
                    row.createCell(4).setCellValue("");                     //备注
                }
            }
        }

        //空出一行
        rowCursor++;

        //工作日历设置
        Row row1 = sheet.createRow(++rowCursor);
        row1.createCell(2).setCellValue("工作日历设置");
        row1.createCell(3).setCellValue(" ");
        //合并“工作日历设置”
        sheet.addMergedRegion(new CellRangeAddress(
                rowCursor, //开始行(以0开始)
                rowCursor, //结束行(以0开始)
                2, //开始列(以0开始)
                3  //结束列(以0开始)
        ));
        //TODO:将该合并处居中
        //创建第二个表

        Row row2 = sheet.createRow(++rowCursor);
        row2.createCell(0).setCellValue("员工编码");
        row2.createCell(1).setCellValue("姓名");
        row2.createCell(2).setCellValue("部门");
        for (Cell cell2 : row2){
            cell2.setCellStyle(styleOfBGC);
        }
        Row row3 = sheet.createRow(++rowCursor);
        row3.createCell(0).setCellValue(id);
        row3.createCell(1).setCellValue(name);
        row3.createCell(2).setCellValue("北京分中心销售室");
        int starTime,endTime;
        Cell tempCell;
        for(int i = 0; i < map.size(); i++){
            row2.createCell(i + 3).setCellValue(i + 1);
            time = map.get(i+1);
            if (time==null || time.equals("") || time.equals(" ")){
                row3.createCell(i+3).setCellValue("公休");
            }else {
                attendanceTime = attendanceTime(map.get(i+1));
                starTime = attendanceTime2Int(attendanceTime[0]);
                endTime = attendanceTime2Int(attendanceTime[1]);
                if(starTime <= 900 && endTime >= 2100){   //全天班
                    row3.createCell(i+3).setCellValue("家乐福项目3");
                }else if(starTime <= 900 && endTime <= 2100){
                    row3.createCell(i+3).setCellValue("家乐福项目1");
                }else if(starTime >=900 && endTime >= 2100){
                    row3.createCell(i+3).setCellValue("家乐福项目2");
                }
            }
        }
        //设置样式
        setStyle(wb,row1);

        //写入文件
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(filePath + name + month + "月份考勤记录.xlsx");
            wb.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 设置样式
     * @param workbook
     */
    public static void setStyle(Workbook workbook, Row rowNoBorder){
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderBottom(CellStyle.BORDER_THIN);   //黑边框
        cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
        cellStyle.setBorderRight(CellStyle.BORDER_THIN);
        cellStyle.setBorderTop(CellStyle.BORDER_THIN);
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);     //居中

        Sheet sheet = workbook.getSheetAt(0);
        for (Row row : sheet) {
            for (Cell cell : row) {
               if (!row.equals(rowNoBorder)){
                   cell.setCellStyle(cellStyle);
               }
            }
        }

       //TODO:怎么将表格尺寸放大，或自动放大尺寸
    }

    /**
     * 获取年月
     * @param str
     * @param needIndex
     * @return
     */
    public static String getYearAndMonth(String str, int needIndex){
        String[] tempStr = str.split("\\D+");   //把数字剥离出来
        if (needIndex < tempStr.length){
            return tempStr[needIndex];          //第二组是月份
        }else {
            throw new ArrayIndexOutOfBoundsException("给的index超过String数组长度!");
        }
    }

    /**
     * 获取个人信息
     * @param str
     * @return
     */
    public static String[] getInfo(String str){
        if(str!=null){
            return str.split("(:|\\s)+");
        }else {
            return null;
        }
    }

    /**
     * 分割打卡时间
     * @param time
     * @return
     */
    public static String[] attendanceTime(String time){
        if (time!=null){
            return time.split(" ");
        }else {
            return null;
        }
    }

    public static int attendanceTime2Int(String time){
        String temp[] = time.split(":");
        return Integer.parseInt(temp[0]+temp[1]);    //将时间，比如8：40转换成840，好比较
    }

    /**
    * 用于测试的main方法
    * @param args
    */
    public static void main(String[] args){
        String detPath = "E:\\AttendanceFile\\";
        String srcPath = "E:\\AttendanceFile\\01原始记录表-14年06月.xls";
        InitAttendanceDataBean initAttendanceDataBean = getAttendanceData(srcPath);
        createNewExcel(0,initAttendanceDataBean,detPath);
    }
}

