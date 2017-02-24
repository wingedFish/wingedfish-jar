package com.wingedfish.office.excel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by wingedfish on 2017/2/24.
 */
public class ReadExcel {

    private static final String XLSX = "xlsx";
    private static final String XLX = "xlx";

    public static List<Map<Integer, String>> readBefore7Excel(String excelUrl, int startRow, int sheetIndex) throws IOException {
        Workbook wb = createWorkWebByVersion(excelUrl, "xlsx");
        // 读取工作薄
        Sheet imp = wb.getSheetAt(sheetIndex);


        return getSheetContent(startRow, imp);
    }

    public static List<Map<Integer, String>> readBefore7Excel(String excelUrl, int startRow, String sheetName) throws IOException {

        Workbook wb = createWorkWebByVersion(excelUrl, "xlsx");
        // 读取工作薄
        Sheet imp = wb.getSheetAt(wb.getSheetIndex(sheetName));


        return getSheetContent(startRow, imp);
    }

    private static List<Map<Integer, String>> getSheetContent(int startRow, Sheet imp) {
        List<Map<Integer, String>> list = new ArrayList();
        // 承装每一行的值
        Map<Integer, String> content = null;

        // 工作薄中的所有行数
        int rows = imp.getLastRowNum();
        //工作薄总列数
        int columnNum = imp.getPhysicalNumberOfRows();
        // 开始行
        Cell value;
        String cont;
        for (int i = startRow; i < rows + 1; i++) {
            content = new HashMap();
            Row row = imp.getRow(i);
            if (row == null) {
                continue;
            }

            for (int j = 0; j < columnNum; j++) {
                value = row.getCell(j);
                if (value != null) {
                    cont = getContent(value);
                    content.put(j, cont);
                }

            }
            list.add(content);
        }
        return list;
    }

    private static InputStream getInputStream(String excelUrl) throws IOException {
        return Files.newInputStream(Paths.get(excelUrl));
    }

    private static String getContent(Cell cell) {
        String str = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case HSSFCell.CELL_TYPE_STRING:
                    str = cell.getStringCellValue();
                    break;

                case HSSFCell.CELL_TYPE_NUMERIC:
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue();
                        if (date != null) {
                            str = new SimpleDateFormat("yyyymmdd").format(date);
                        } else {
                            str = new SimpleDateFormat("yyyymmdd").format(new Date());
                        }

                    } else {
                        str = new DecimalFormat("0").format(cell.getNumericCellValue());
                    }
                    break;
                default:
                    str = "unknow";
            }

        }
        return str;
    }

    private static Workbook createWorkWebByVersion(String excelUrl, String suffix) throws IOException {
        InputStream in = getInputStream(excelUrl);

        Workbook workbook = null;
        if (XLSX.equals(suffix)) {
            workbook = new XSSFWorkbook(in);
        } else if (XLX.equals(suffix)) {
            POIFSFileSystem fsys = new POIFSFileSystem(in);
            workbook = new HSSFWorkbook(fsys);
        }
        return workbook;
    }
}
