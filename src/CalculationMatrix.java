import com.sun.swing.internal.plaf.synth.resources.synth_sv;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.omg.CORBA.MARSHAL;

import java.io.*;
import java.math.BigDecimal;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by:
 * author 刘春东
 * time 14-8-2下午2:20
 * version 0.1
 */
public class CalculationMatrix {
    /**
     * 读取源Excel
     * @param file 源文件
     * @return Workbook
     * @throws java.io.IOException
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

    public static void calculate(Workbook workbook, Sheet sheet, String detPath){
        Row row1,rowCursor;
        int i, k;
        row1 = sheet.getRow(1);         //读取第二行Row1，人均GDP
        int len = row1.getPhysicalNumberOfCells();
        double[] gdp = new double[len];
        double[][] aa = new double[len][len];
        double denominatorOfIK = 0.0;
        for (int j = 1; j < len; j++){
            gdp[j] = row1.getCell(j).getNumericCellValue();
        }
        for (i = 1; i < len; i++){
            denominatorOfIK = calculateDenominatorOfIK(gdp,i,len);
            for (k = 1; k < len; k++){
                if(k != i)
                    aa[i][k] = formatDecimal( 1/Math.abs(gdp[i] - gdp[k]) / denominatorOfIK);
            }
        }
        System.out.println("已经计算出aa数组，接下来写入Excel对应的表格中");

        //修改值
        Cell cell;
        for (int n = 2; n < len + 1; n++){
            rowCursor = sheet.getRow(n);
            for (int m = 1; m < len; m++){
                rowCursor.getCell(m).setCellValue(aa[n-1][m]);
            }
        }

        //写入文件
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(detPath+"计算矩阵W.xlsx");
            workbook.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("计算成功!");
    }

    public static double calculateDenominatorOfIK(double[] gdp, int i, int len){
        double temp = 0.0;
        for (int j = 1; j < len; j++){
            if (j != i){  //如果两弟的GDP一样了，就认为其空间经济系数为0
                if(gdp[j] == gdp[i])
                    temp += 0.0;
                else
                    temp += (1/Math.abs(gdp[i] - gdp[j]));
            }
        }
        return temp;
    }

    /**
     * 格式化最后的数据
     * @param d
     * @return
     */
    public static double formatDecimal(double d){
        BigDecimal b = new BigDecimal(d);
        return b.setScale(4,BigDecimal.ROUND_HALF_UP).doubleValue();
    }
    /**
     * 用于测试的main方法
     * @param args
     */
    public static void main(String[] args){
        Workbook wb;
        Sheet sheet;
        String srcPath = "D:\\DOWNLOAD\\春竹姐论文\\含空间矩阵的偏离-份额分析法.xlsx";
        String detPath = "D:\\DOWNLOAD\\春竹姐论文\\";
        try {
            wb = readExcelFile(new File(srcPath));
            if (wb != null) {
                //sheet = wb.getSheet("2008A");
                //alculate(wb,sheet,detPath);
                sheet = wb.getSheet("2012A");
                calculate(wb,sheet,detPath);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
