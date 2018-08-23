package com.Leo.excel;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;

/**
 * @Auther: Leo
 * @Date: 2018/8/23 18:53
 * @Description:
 */
public class PoiExpExcel2 {
    /**
     * POI生成excel文件
     * @param args
     */
    public static void main(String[] args){
        String[] title = {"id","name","sex"};
        //创建Excel工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        //创建一个工作表
        Sheet sheet = workbook.createSheet();
        //创建第一行
        Row row = sheet.createRow(0);
        Cell cell = null;
        //插入第一行数据id，name，sex
        for(int i=0;i<title.length;i++){
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
        }
        //追加数据
        for(int i=1;i<=10;i++){
            Row nextrow = sheet.createRow(i);
            Cell cell2 = nextrow.createCell(0);
            cell2.setCellValue("a"+i);
            cell2 = nextrow.createCell(1);
            cell2.setCellValue("user"+i);
            cell2 = nextrow.createCell(2);
            cell2.setCellValue("男");
        }

        //创建一个文件
        File file = new File("c:/github/poi_test.xlsx");
        try {
            file.createNewFile();
            //将excel内容存盘
            FileOutputStream stream = FileUtils.openOutputStream(file);
            workbook.write(stream);
            stream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
