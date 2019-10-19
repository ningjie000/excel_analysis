package com.utils;


import com.constant.ExcelConstant;
import com.handle.ExcelXlsxReaderWithDefaultHandler;

import java.util.List;

/**
 * @author qjwyss
 * @date 2018/12/19
 * @description 读取EXCEL工具类
 */
public class ExcelReaderUtil {


    public static void readExcel(String filePath) throws Exception {
        int totalRows = 0;
        if (filePath.endsWith(ExcelConstant.EXCEL07_EXTENSION)) {
            ExcelXlsxReaderWithDefaultHandler excelXlsxReader = new ExcelXlsxReaderWithDefaultHandler();
            totalRows = excelXlsxReader.process(filePath);
            List<List<String>> alldata = excelXlsxReader.getAlldata();
            System.out.println(alldata);
        } else {
            throw new Exception("文件格式错误，fileName的扩展名只能是xlsx!");
        }
        System.out.println("读取的数据总行数：" + totalRows);
    }


    public static void main(String[] args) throws Exception {
        System.out.println("开始解析");
        long start = System.currentTimeMillis();
        String path = "E:\\temp.xlsx";
        ExcelReaderUtil.readExcel(path);
        long end = System.currentTimeMillis();
        System.out.println();
        System.out.println("耗时：" + (end -start) /1000);
    }
}
