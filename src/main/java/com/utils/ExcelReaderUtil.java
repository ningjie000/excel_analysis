package com.utils;


import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.constant.ExcelConstant;
import com.handle.ExcelXlsxReaderWithDefaultHandler;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * @author NJ
 */
public class ExcelReaderUtil {


    public static List<List<String>> readExcel(String filePath) throws Exception {
        List<List<String>> alldata = null;
        if (filePath.endsWith(ExcelConstant.EXCEL07_EXTENSION)) {
            ExcelXlsxReaderWithDefaultHandler excelXlsxReader = new ExcelXlsxReaderWithDefaultHandler(true);
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
            outStream = new FileOutputStream(filePath);
            byte[] buffer = new byte[1024];
            int bytesRead;
            while ((bytesRead = swapStream.read(buffer)) != -1) {
                outStream.write(buffer, 0, bytesRead);
            }



            /************* 输出流转输入流方法 *************/
//            ByteArrayOutputStream baos = new ByteArrayOutputStream();
//            EasyExcel.write(baos)
//                    // 这里放入动态头
//                    .head(head()).excelType(ExcelTypeEnum.XLSX).sheet(1)
//                    // 当然这里数据也可以用 List<List<String>> 去传入
//                    .doWrite(data());
//            long end = System.currentTimeMillis();
//            System.out.println("耗时：" + (end -startTime) /1000);
//
//            InputStream swapStream = new ByteArrayInputStream(baos.toByteArray());

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


    /**
     * 合并两张表格数据
     * @param excelOne 表格1
     * @param excelTwo 表格2
     * @return
     */
    public static List<List<String>> mergeExcel(List<List<String>> excelOne, List<List<String>> excelTwo){
        List<List<String>> result = new ArrayList<List<String>>();

        int tempMaxSize = 0;
        int tempMinSize = 0;

        if(excelOne.size() > excelTwo.size()){
            tempMaxSize = excelOne.size();
            tempMinSize = excelTwo.size();

            /** 获取行数少的list，需要填补 **/
            int columnSize = excelTwo.get(0).size();

            for(int i = 0; i < tempMaxSize; i++){

                List<String> value1 = excelOne.get(i);
                if(i < tempMinSize){
                    List<String> value2 = excelTwo.get(i);
                    value1.addAll(value2);
                }else{
                    /** 空串占位 **/
                    for(int k = 0; k < columnSize; k++){
                        value1.add("");
                    }
                }
                result.add(value1);
            }



        }else{
            tempMaxSize = excelTwo.size();
            tempMinSize = excelOne.size();

            /** 获取行数少的list，需要填补 **/
            int columnSize = excelOne.get(0).size();

            for(int i = 0; i < tempMaxSize; i++){

                List<String> value2 = excelTwo.get(i);
                List<String> temp = new ArrayList<String>();
                if(i < tempMinSize){
                    List<String> value1 = excelOne.get(i);
                    value1.addAll(value2);
                    result.add(value1);
                }else{
                    /** 空串占位 **/
                    for(int j = 0; j < columnSize; j++){
                        temp.add("");
                    }
                    temp.addAll(value2);
                    result.add(temp);
                }
            }
        }


        return result;


    }


    public static void main(String[] args) throws Exception {
        List<List<String>> lists1 = readExcel("D:\\temp\\113asd.xlsx");
        List<List<String>> lists2 = readExcel("D:\\temp\\113asd.xlsx");

        List<List<String>> lists = mergeExcel(lists1, lists2);
//        System.out.println(lists);

        List<List<String>> heads = new ArrayList<List<String>>();
        List<String> strings = lists.get(0);
        for(int i = 0; i< strings.size() ;i++){
            List<String> head = new ArrayList<String>();
            head.add(strings.get(i));
            heads.add(head);
        }

        lists.remove(0);

        List<List<String>> datas = new ArrayList<List<String>>();
        datas.addAll(lists);
        System.out.println("合并开始");
        long l = System.currentTimeMillis();
        writeExcelXLSX("D:\\temp\\ddd1111bew.xlsx",heads,datas);
        long l2 = System.currentTimeMillis();
        System.out.println("合并结束,耗时 ： " + (l2-l)/1000 + "s");
    }




}
