package com.utils;


import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.constant.ExcelConstant;
import com.handle.ExcelXlsxReaderWithDefaultHandler;

import java.io.*;
import java.util.List;

/**
 * @author NJ
 */
public class ExcelReaderUtil {


    public static List<List<String>> readExcel(String filePath) throws Exception {
        List<List<String>> alldata = null;
        if (filePath.endsWith(ExcelConstant.EXCEL07_EXTENSION)) {
            ExcelXlsxReaderWithDefaultHandler excelXlsxReader = new ExcelXlsxReaderWithDefaultHandler();
            excelXlsxReader.process(filePath);
            alldata = excelXlsxReader.getAlldata();
        } else {
            throw new Exception("文件格式错误，fileName的扩展名只能是xlsx!");
        }

        return alldata;
    }

    public static Boolean writeExcelXLSX(String filePath, List<List<String>> head, List<List<String>> values) throws Exception {
        long startTime = System.currentTimeMillis();
        ByteArrayOutputStream baos = null;
        InputStream swapStream = null;
        OutputStream outStream = null;

        try {
            baos = new ByteArrayOutputStream();
            EasyExcel.write(baos)
                    // 这里放入动态头
                    .head(head).excelType(ExcelTypeEnum.XLSX).sheet(1)
                    // 当然这里数据也可以用 List<List<String>> 去传入
                    .doWrite(values);
            long end = System.currentTimeMillis();
            System.out.println("耗时：" + (end -startTime) /1000);

            swapStream = new ByteArrayInputStream(baos.toByteArray());
            outStream = new FileOutputStream("filePath");
            byte[] buffer = new byte[1024];
            int bytesRead;
            while ((bytesRead = swapStream.read(buffer)) != -1) {
                outStream.write(buffer, 0, bytesRead);
            }
        } catch (Exception e){
            e.printStackTrace();
        } finally {
            if(outStream != null){
                outStream.close();
            }
            if(outStream != null){
                swapStream.close();
            }
            if(outStream != null){
                baos.close();
            }
        }
        return true;
    }


}
