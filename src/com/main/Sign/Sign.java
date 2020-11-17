package com.main.Sign;

import java.io.*;
import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import com.sun.xml.internal.bind.v2.model.core.ID;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * ClassName:Sign
 * Package:PACKAGE_NAME
 * Description:数据是从19列开始的
 *
 * @date:2020/1/6 14:50
 * @author:Wang Jun
 */
public class Sign {
    public static void main(String[] args) throws IOException {
        String dingding = "上海海洋大学智慧海洋_考勤报表_20201109-20201113";
        String leave = "leave20201109-20201113";
        String white_list = "白名单";
        String gradeOne_str="研一同学";
        int day = 5;    //统计的天数


        File file = new File(".\\src\\com\\main\\钉钉报表\\" + dingding + ".xlsx");
        File holidaies = new File(".\\src\\com\\main\\请假出差\\" + leave + ".xlsx");
        File free = new File(".\\src\\com\\main\\白名单\\" + white_list + ".xlsx");
        File gradeOne = new File("src\\com\\main\\白名单\\"+gradeOne_str+".xlsx");
        //请假出差和请假人员
        List<List<String>> list_holiday = new ArrayList<List<String>>();
        InputStream is = new FileInputStream(holidaies.getAbsolutePath());
        XSSFWorkbook wb = new XSSFWorkbook(is);
        Sheet sheet = wb.getSheetAt(0);
        for (int i = 0; i < day; ++i) {
            List<String> list_holiday_one= new ArrayList<String>();
            for (int j = 1; j < 3; ++j) {
                    list_holiday_one.add(sheet.getRow(i + 1).getCell(j).toString());
            }
            list_holiday.add(list_holiday_one);
        }
        for (List<String> list :list_holiday)
            System.out.println(list);

        //读取研一同学
        InputStream is_one = new FileInputStream(gradeOne.getAbsolutePath());
        XSSFWorkbook wb_one = new XSSFWorkbook(is_one);
        Sheet sheet_gradeOne = wb_one.getSheetAt(0);
        List<String> gradeOne_name =new ArrayList<String>();
        for (int i = 0;i<sheet_gradeOne.getPhysicalNumberOfRows();++i){
            gradeOne_name.add(sheet_gradeOne.getRow(i).getCell(0).toString());
        }

//        读取白名单
        InputStream is_free = new FileInputStream(free.getAbsolutePath());
        XSSFWorkbook wb_free = new XSSFWorkbook(is_free);
        Sheet sheet_free = wb_free.getSheetAt(0);
        List<String> list_free = new ArrayList<String>();
        for (int j = 0; j < sheet_free.getPhysicalNumberOfRows(); ++j) {
            if (sheet_free.getRow(j).getCell(0).getCellType() != Cell.CELL_TYPE_BLANK) {
                list_free.add(sheet_free.getRow(j).getCell(0).toString());
            }
        }


        System.out.println(file.getName());
        List<Students> list = readExcel(file, list_holiday,list_free);
        for (Students students : list)
            System.out.println(students);
        makeExcel(list, dingding,gradeOne_name);
    }

    //读取excel表格
    public static List<Students> readExcel(File file, List<List<String>> list_holiday,List<String> list_free) {
        List<Students> list = new ArrayList<Students>();
        try {
            // 创建输入流，读取Excel
            InputStream is = new FileInputStream(file.getAbsolutePath());
            // jxl提供的Workbook类
            XSSFWorkbook wb = new XSSFWorkbook(is);
            // Excel的页签数量
            Sheet sheet = wb.getSheetAt(0);
            for (int i = 4; i <sheet.getPhysicalNumberOfRows(); i++) {
                // sheet.getColumns()返回该页的总列数
                //去除白名单老师
                if (list_free.contains(sheet.getRow(i).getCell(0).toString()))
                    continue;
                Students students = new Students();
                students.setName(sheet.getRow(i).getCell(0).toString());
                int a = 0;  //出差
                int b = 0;  //迟到
                int d = 0;  //缺卡
                int e = 0;  //旷工
                int unnormal = 0;    //由于旷工产生的异常次数
                for (int j = 19; j < sheet.getRow(0).getPhysicalNumberOfCells(); j++) {
                    //统计请假出差
                    if (list_holiday.get(j-19).get(0).contains(sheet.getRow(i).getCell(0).toString())||list_holiday.get(j-19).get(1).contains(sheet.getRow(i).getCell(0).toString())) {
                        a += 1;
                        continue;
                    }

                    a += Count(sheet.getRow(i).getCell(j).toString(), "出差");
                    b += Count(sheet.getRow(i).getCell(j).toString(), "迟到");
                    d += Count(sheet.getRow(i).getCell(j).toString(), "缺卡");
                    if("旷工".equals(sheet.getRow(i).getCell(j).toString())){
                        ++e;
                        if (j==21 || j==23)
                            unnormal+=4;
                        else unnormal+=6;
                    }
                    if (j == 21 || j == 23) {
                        b -= Count(sheet.getRow(i).getCell(j).toString(), "3迟到");
                        b -= Count(sheet.getRow(i).getCell(j).toString(), "3严重迟到");
                        d -= Count(sheet.getRow(i).getCell(j).toString(), "3缺卡");
                    }
                }
                students.setEvection(a);
                students.setLate(b);
                students.setLose(d);
                students.setAbsenteeism(e);
                students.setNormal(students.getNormal() -b-d-unnormal);
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
    public static int Count(String string1, String string2) {
        int count = 0;
        for (int i = 0; i < string1.length(); ++i) {
            if (string1.indexOf(string2, i) == i) {
                ++count;
            }
        }
        return count;
    }

    //制表
    public static void makeExcel(List<Students> list, String tableName,List<String> gradeOne) throws FileNotFoundException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("出勤统计表");
        HSSFRow row = sheet.createRow(0);
        row.createCell(0).setCellValue("姓名");
        row.createCell(1).setCellValue("迟到");
        row.createCell(2).setCellValue("缺卡");
        row.createCell(3).setCellValue("出差/请假");
        row.createCell(4).setCellValue("旷工(天)");
        row.createCell(5).setCellValue("正常");
        row.createCell(6).setCellValue("出勤率");
        CellStyle cellStyle = workbook.createCellStyle();
//        cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
//        cellStyle.setFillForegroundColor(HSSFColor.RED.index);
        //声明一个字体对象
        Font font = null;
        //创建一个字体对象
        font = workbook.createFont();
        //给字体对象设置颜色属性
        font.setColor(HSSFColor.RED.index);
        //将字体对象放入XSSFCellStyle对象中
        cellStyle.setFont(font);
        for (int i = 0; i < list.size(); ++i) {
            Students students = list.get(i);
            HSSFRow row1 = sheet.createRow(i + 1);
            if (gradeOne.contains(students.getName())){
                row1.createCell(0).setCellValue(students.getName());
                row1.createCell(1).setCellValue(0);
                row1.createCell(2).setCellValue(0);
                row1.createCell(3).setCellValue(students.getEvection());
                row1.createCell(4).setCellValue(0);
                row1.createCell(5).setCellValue(26);
                row1.createCell(6).setCellValue(1);
                continue;
            }
            row1.createCell(0).setCellValue(students.getName());
            row1.createCell(1).setCellValue(students.getLate());
            if (students.getLate()!=0)
                row1.getCell(1).setCellStyle(cellStyle); //标记迟到次数
            row1.createCell(2).setCellValue(students.getLose());
            if (students.getLose()!=0)
                row1.getCell(2).setCellStyle(cellStyle); //标记缺卡次数
            row1.createCell(3).setCellValue(students.getEvection());
            row1.createCell(4).setCellValue(students.getAbsenteeism());
            if (students.getAbsenteeism()!=0)
                row1.getCell(4).setCellStyle(cellStyle);  //标记旷工次数
            row1.createCell(5).setCellValue(students.getNormal());
            row1.createCell(6).setCellValue(students.attendrat());
        }
        try {
            FileOutputStream fos = new FileOutputStream(".\\src\\com\\main\\统计结果\\" + tableName + "统计结果.xls");
            workbook.write(fos);
            System.out.println("写入成功！");
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
