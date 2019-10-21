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
 * @Description: 底层基于阿里的开源项目easyexcel
 *                 但是由于easyexcel自己提供的解析，如果字段中出现空字段，会出现格式乱掉的情况
 *                 所以在读取的时候我采用sax解析xlsx文件，用为xlsx文件解压后为xml文件
 *                 ExcelXlsxReaderWithDefaultHandler构造方法传参true 或者 false，
 *                     如果传入true 则会通过第一行，即head的列数来限制字段，多了的不读，少了的补空占位
 *                     如果是false，按照正常格式读取，每行列数可以不定
 */
public class ExcelReaderUtil {

    /**
     * 通过文件url读取excel
     * @param filePath  文件地址
     * @param isLimitColumnNum 是否更具head的列数来限制读取的cell的列数
     * @return 所有数据，包括head和cell
     * @throws Exception
     */
    public static List<List<String>> readExcel(String filePath, Boolean isLimitColumnNum) throws Exception {
        List<List<String>> alldata = null;
        if (filePath.endsWith(ExcelConstant.EXCEL07_EXTENSION)) {
            ExcelXlsxReaderWithDefaultHandler excelXlsxReader = null;
            if(isLimitColumnNum != null){
                excelXlsxReader = new ExcelXlsxReaderWithDefaultHandler(isLimitColumnNum);
            }else{
                excelXlsxReader = new ExcelXlsxReaderWithDefaultHandler();
            }
            excelXlsxReader.process(filePath);
            alldata = excelXlsxReader.getAlldata();
        } else {
            throw new Exception("文件格式错误，fileName的扩展名只能是xlsx!");
        }

        return alldata;
    }


    /**
     * 通过文件url读取excel
     * @param inputStream  输入的文件流
     * @param isLimitColumnNum 是否更具head的列数来限制读取的cell的列数
     * @return 所有数据，包括head和cell
     * @throws Exception
     */
    public static List<List<String>> readExcel(InputStream inputStream, Boolean isLimitColumnNum) throws Exception {
        List<List<String>> alldata = null;

        try {
            ExcelXlsxReaderWithDefaultHandler excelXlsxReader = null;
            if(isLimitColumnNum != null){
                excelXlsxReader = new ExcelXlsxReaderWithDefaultHandler(isLimitColumnNum);
            }else{
                excelXlsxReader = new ExcelXlsxReaderWithDefaultHandler();
            }
            excelXlsxReader.process(inputStream);
            alldata = excelXlsxReader.getAlldata();
        } catch (Exception e){
            e.printStackTrace();
        }finally {
            if(inputStream != null){
                inputStream.close();
            }
        }

        return alldata;
    }


    /**
     * 将数据写入到excel中
     * @param filePath filePath输出文件地址
     * @param head
     * @param values
     * @return
     * @throws Exception
     */
    public static Boolean writeExcelXLSX(String filePath, List<List<String>> head, List<List<String>> values) throws Exception {
        long startTime = System.currentTimeMillis();
        ByteArrayOutputStream baos = null;
//        InputStream swapStream = null;
        OutputStream outStream = null;

        try {
//            baos = new ByteArrayOutputStream(); 输出流转输入流可用此方法， write(baos)
            outStream = new FileOutputStream(filePath);
            EasyExcel.write(outStream)
                    // 这里放入动态头
                    .head(head).excelType(ExcelTypeEnum.XLSX).sheet(1)
                    // 当然这里数据也可以用 List<List<String>> 去传入
                    .doWrite(values);
            long end = System.currentTimeMillis();
            System.out.println("耗时：" + (end -startTime) /1000);

            /** 将输出流转输入流，输入流 **/
//            swapStream = new ByteArrayInputStream(baos.toByteArray());
//            outStream = new FileOutputStream(filePath);
//            byte[] buffer = new byte[1024];
//            int bytesRead;
//            while ((bytesRead = swapStream.read(buffer)) != -1) {
//                outStream.write(buffer, 0, bytesRead);
//            }


        } catch (Exception e){
            e.printStackTrace();
        } finally {
            if(outStream != null){
                outStream.close();
            }
//            if(outStream != null){
//                swapStream.close();
//            }
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
        List<List<String>> lists1 = readExcel("D:\\temp\\113asd.xlsx", true);
        List<List<String>> lists2 = readExcel("D:\\temp\\113asd.xlsx", true);

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
