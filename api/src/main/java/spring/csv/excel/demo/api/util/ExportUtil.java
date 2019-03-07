package spring.csv.excel.demo.api.util;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.BufferedWriter;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * ExportUtil
 *
 * @author Lunhao Hu
 * @date 2019-03-02 10:36
 **/
public class ExportUtil {

    /**
     * 导出csv文件
     *
     * @param headers    内容标题
     * @param exportData 要导出的数据集合
     * @return
     */
    public static byte[] exportCSV(LinkedHashMap<String, Object> headers, List<LinkedHashMap<String, Object>> exportData) {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        BufferedWriter buffCvsWriter = null;
        try {
            buffCvsWriter = new BufferedWriter(new OutputStreamWriter(out, StandardCharsets.UTF_8));
            // 将header数据写入表格
            fillDataToCsv(buffCvsWriter, headers);
            buffCvsWriter.newLine();
            // 将body数据写入表格
            for (Iterator<LinkedHashMap<String, Object>> iterator = exportData.iterator(); iterator.hasNext(); ) {
                fillDataToCsv(buffCvsWriter, iterator.next());
                if (iterator.hasNext()) {
                    buffCvsWriter.newLine();
                }
            }
            // 刷新缓冲
            buffCvsWriter.flush();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // 释放资源
            if (buffCvsWriter != null) {
                try {
                    buffCvsWriter.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return out.toByteArray();
    }

    public static byte[] exportXlsx(LinkedHashMap<String, List<LinkedHashMap<String, Object>>> tableData) {
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            HSSFWorkbook workbook = new HSSFWorkbook();
            // 创建多个sheet

            Iterator<Map.Entry<String, List<LinkedHashMap<String, Object>>>> iterator = tableData.entrySet().iterator();
            while (iterator.hasNext()) {
                String sheetName = iterator.next().getKey();
                List<LinkedHashMap<String, Object>> value = iterator.next().getValue();

//                fillDataToXlsx(workbook.createSheet(sheetName), value);
                System.out.println(sheetName);
                System.out.println(value);
            }

            fillDataToXlsx(workbook.createSheet("日报表"), tableData.get("日报表"));
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return out.toByteArray();
    }

    /**
     * 将linkedHashMap中的数据，写入xlsx表格中
     *
     * @param sheet
     * @param data
     */
    private static void fillDataToXlsx(HSSFSheet sheet, List<LinkedHashMap<String, Object>> data) {
        HSSFRow currRow;
        HSSFCell cell;
        LinkedHashMap row;
        Map.Entry propertyEntry;
        int rowIndex = 0;
        int cellIndex = 0;
        for (Iterator<LinkedHashMap<String, Object>> iterator = data.iterator(); iterator.hasNext(); ) {
            row = iterator.next();
            currRow = sheet.createRow(rowIndex++);
            for (Iterator<Map.Entry> propertyIterator = row.entrySet().iterator(); propertyIterator.hasNext(); ) {
                propertyEntry = propertyIterator.next();
                if (propertyIterator.hasNext()) {
                    String value = String.valueOf(propertyEntry.getValue());
                    cell = currRow.createCell(cellIndex++);
                    cell.setCellValue(value);
                } else {
                    String value = String.valueOf(propertyEntry.getValue());
                    cell = currRow.createCell(cellIndex++);
                    cell.setCellValue(value);
                    break;
                }
            }
            if (iterator.hasNext()) {
                cellIndex = 0;
            }
        }
    }

    private static void fillDataToCsv(BufferedWriter buffCvsWriter, LinkedHashMap row) throws IOException {
        Map.Entry propertyEntry;
        for (Iterator<Map.Entry> propertyIterator = row.entrySet().iterator(); propertyIterator.hasNext(); ) {
            propertyEntry = propertyIterator.next();
            buffCvsWriter.write("\"" + propertyEntry.getValue().toString() + "\"");
            if (propertyIterator.hasNext()) {
                buffCvsWriter.write(",");
            }
        }
    }

}
