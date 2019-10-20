package com.read;

import com.utils.ExcelReaderUtil;
import org.junit.Test;

import java.util.List;

/**
 * @author NJ
 * @date 2019/10/20
 * @description
 */
public class ReadTest {

    @Test
    public void testRead1() throws Exception{
        System.out.println("开始解析");
        long start = System.currentTimeMillis();
        String path = "E:\\temp.xlsx";
        List<List<String>> lists = ExcelReaderUtil.readExcel(path);
        long end = System.currentTimeMillis();
        System.out.println();
        System.out.println("耗时：" + (end -start) /1000);
    }
}
