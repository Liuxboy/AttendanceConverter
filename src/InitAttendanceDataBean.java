import java.util.HashMap;
import java.util.Map;

/**
 * Created with IntelliJ IDEA13.0
 * User:cdliu
 * Date:14-6-5
 * Time:上午11:09
 * Version:0.1
 */
public class InitAttendanceDataBean {
    private String year;        //年份
    private String month;       //月份
    private String[] info;      //个人信息
    private Map<Integer,String> attendanceMap;  //日期与考勤

    public String getYear() {
        return year;
    }

    public void setYear(String year) {
        this.year = year;
    }

    public String getMonth() {
        return month;
    }

    public void setMonth(String month) {
        this.month = month;
    }

    public String[] getInfo() {
        return info;
    }

    public void setInfo(String[] info) {
        this.info = info;
    }

    public Map<Integer, String> getAttendanceMap() {
        return attendanceMap;
    }

    public void setAttendanceMap(Map<Integer, String> attendanceMap) {
        this.attendanceMap = attendanceMap;
    }
}
