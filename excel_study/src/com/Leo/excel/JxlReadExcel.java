package com.Leo.excel;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.IOException;

/**
 * @Auther: Leo
 * @Date: 2018/8/23 16:56
 * @Description:
 */
public class JxlReadExcel {
    /**
     * JXL解析Excel
     */
    public static  void main(String[] args)  {
        try {
            //创建workbook
            Workbook workbook = Workbook.getWorkbook(new File("c:/github/jxl_test.xls"));
            //获取工作表sheet
            Sheet sheet = workbook.getSheet(0);
            //获取数据
            for(int i = 0;i<sheet.getRows();i++){
                for (int j = 0;j<sheet.getColumns();j++){
                    Cell cell = sheet.getCell(j,i);
                    System.out.print(cell.getContents()+" ");
                }
                System.out.println();
            }
            workbook.close();

        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
