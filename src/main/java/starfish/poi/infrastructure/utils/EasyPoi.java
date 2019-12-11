package starfish.poi.infrastructure.utils;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;

import java.util.*;

@Slf4j
@Component
public class EasyPoi {

    private static final Set<String> Sheet = new LinkedHashSet<String>();

    public XSSFWorkbook createWorkBook(List<Map<String,String>> array) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("lz");
        //设置单元格数据居中z
        XSSFCellStyle rowStyle = workbook.createCellStyle();
        rowStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);


        log.info("填充数据......");
        int rowNum = 1;
        for (Map<String,String> objectMap: array) {
            XSSFRow row = sheet.createRow(rowNum);
            int fieldIndex = 0;
            for (Map.Entry<String,String> entry:objectMap.entrySet()){
                XSSFCell cell = row.createCell(fieldIndex);
                cell.setCellValue(entry.getValue());
                Sheet.add(entry.getKey());
                cell.setCellStyle(rowStyle);
                fieldIndex ++;
            }
            rowNum++;
        }

        log.info("创建表头......");
        createTitle(workbook, sheet);
        return workbook;
    }

    // 创建表头
    private void createTitle(XSSFWorkbook workbook, XSSFSheet sheet) {
        for (int i=0;i<Sheet.size();i++){
            // 设置列宽，setColumnWidth的第二个参数要乘以256，这个参数的单位是1/256个字符宽度
            sheet.setColumnWidth(i, 20 * 256);
        }

        // 设置为居中加粗
        XSSFCellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        style.setFont(font);

        XSSFCell cell;
        Iterator<String> iterable = Sheet.iterator();
        XSSFRow row = sheet.createRow(0);
        int filedIndex = 0;
        while (iterable.hasNext()){
            cell = row.createCell(filedIndex);
            cell.setCellValue(iterable.next());
            cell.setCellStyle(style);
            filedIndex ++;
        }
    }
}
