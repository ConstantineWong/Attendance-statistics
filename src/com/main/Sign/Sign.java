package Sign;

import java.io.*;
import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import com.sun.xml.internal.bind.v2.model.core.ID;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 * ClassName:Sign
 * Package:PACKAGE_NAME
 * Description:
 *
 * @date:2020/1/6 14:50
 * @author:Wang Jun
 */
public class Sign {
    public static void main(String[] args) throws IOException {
        File file = new File(".\\src\\com\\main\\Sign\\上海海洋大学智慧海洋_考勤报表_20200106-20200110.xlsx");
        File holidaies = new File(".\\src\\com\\main\\Sign\\leave.xlsx");

        //请假出差人员
        StringBuffer[] holiday = new StringBuffer[5];
        InputStream is = new FileInputStream(holidaies.getAbsolutePath());
        XSSFWorkbook wb = new XSSFWorkbook(is);
        Sheet sheet = wb.getSheetAt(0);
        for (int i = 0;i<5;++i){
            holiday[i] = new StringBuffer();
            for (int j =1;j<3;++j){
                if (sheet.getRow(i+1).getCell(j)!=null && sheet.getRow(i+1).getCell(j).getCellType()!= Cell.CELL_TYPE_BLANK)
                holiday[i].append(sheet.getRow(i+1).getCell(j).toString());
            }
        }
        for (StringBuffer stringBuffer:holiday)
            System.out.println(stringBuffer);

        System.out.println(file.getName());
        List<Students> list = readExcel(file,holiday);
        for (Students students:list)
            System.out.println(students);
        makeExcel(list);
    }

    //读取excel表格
    public static List<Students> readExcel(File file,StringBuffer[] stringBuffers) {
        List<Students> list = new ArrayList<Students>();
        try {
            // 创建输入流，读取Excel
            InputStream is = new FileInputStream(file.getAbsolutePath());
            // jxl提供的Workbook类
            XSSFWorkbook wb = new XSSFWorkbook(is);
            // Excel的页签数量
            Sheet sheet = wb.getSheetAt(0);
            for (int i = 4; i < sheet.getPhysicalNumberOfRows(); i++) {
                // sheet.getColumns()返回该页的总列数
                if(sheet.getRow(i).getCell(0).toString().equals("郑小罗")||sheet.getRow(i).getCell(0).toString().equals("赵丹枫")||sheet.getRow(i).getCell(0).toString().equals("张明华")||sheet.getRow(i).getCell(0).toString().equals("杜艳玲"))
                    continue;
                Students students = new Students();
                students.setName(sheet.getRow(i).getCell(0).toString());
                int a =0;
                int b =0;
                int d =0;
                    for (int j = 19; j < sheet.getRow(0).getPhysicalNumberOfCells(); j++) {
                        if (stringBuffers.toString().contains(sheet.getRow(i).getCell(0).toString())){
                            a+=Count(sheet.getRow(i).getCell(j).toString(),"出差");
                            continue;
                        }
                        a+=Count(sheet.getRow(i).getCell(j).toString(),"出差");
                        b+=Count(sheet.getRow(i).getCell(j).toString(),"迟到");
                        d+=Count(sheet.getRow(i).getCell(j).toString(),"缺卡");
                        if (j==21 || j==23){
                        b-=Count(sheet.getRow(i).getCell(j).toString(),"3迟到");
                        d-=Count(sheet.getRow(i).getCell(j).toString(),"3缺卡");
                        }
                    }
                students.setEvection(a);
                students.setLate(b);
                students.setLose(d);
                students.setNormal(26-d);
                list.add(students);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return list;
    }
    //统计次数
    public static int Count(String string1,String string2){
        int count = 0;
        for (int i = 0;i<string1.length();++i){
            if (string1.indexOf(string2,i) == i){
                ++count;
            }
        }
        return count;
    }
    //制表
    public static void makeExcel(List<Students> list) throws FileNotFoundException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("出勤统计表");
        HSSFRow row = sheet.createRow(0);
        row.createCell(0).setCellValue("姓名");
        row.createCell(1).setCellValue("迟到");
        row.createCell(2).setCellValue("缺卡");
        row.createCell(3).setCellValue("出差/请假");
        row.createCell(4).setCellValue("正常");
        row.createCell(5).setCellValue("出勤率");
        for (int i = 0;i<list.size();++i){
            Students students = list.get(i);
            HSSFRow row1 = sheet.createRow(i+1);
            row1.createCell(0).setCellValue(students.getName());
            row1.createCell(1).setCellValue(students.getLate());
            row1.createCell(2).setCellValue(students.getLose());
            row1.createCell(3).setCellValue(students.getEvection());
            row1.createCell(4).setCellValue(students.getNormal());
            row1.createCell(5).setCellValue(students.attendrat());
        }
        try {
            FileOutputStream fos = new FileOutputStream(".\\src\\com\\main\\Sign\\上海海洋大学智慧海洋_考勤报表_20200106-20200110.xls");
            workbook.write(fos);
            System.out.println("写入成功！");
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
