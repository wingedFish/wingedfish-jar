package com.wingedfish.office.excel;

import org.junit.Test;

import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * Created by lixiuhai on 2017/2/24.
 */
public class ReadExcelT {

    @Test
    public void testRead() throws IOException {

        List<Map<Integer, String>> contextList = ReadExcel.readBefore7Excel("D:\\Git_workespace\\Git_systemalarm\\wingedfish-jar\\大生活讲堂.xlsx", 0, 0);
        contextList.forEach((contextMap) -> contextMap.forEach((key, value) -> System.out.println(key + " , " + value)));

    }

}
