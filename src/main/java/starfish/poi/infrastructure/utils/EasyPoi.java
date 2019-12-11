package starfish.poi.infrastructure.utils;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;
import starfish.poi.domain.WaterMeter;

import java.util.Arrays;
import java.util.List;

@Slf4j
@Component
public class EasyPoi {

    private static final List<String> Sheet = Arrays.asList("id","name","time","data","status");

    public XSSFWorkbook createWorkBook(List<WaterMeter> waterMeters) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("lz");

        log.info("创建表头......");
        createTitle(workbook, sheet);

        //设置单元格数据居中z
        XSSFCellStyle rowStyle = workbook.createCellStyle();
        rowStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);

        //固定列
        int rowNum = 1;
        for (WaterMeter waterMeter: waterMeters) {
            XSSFRow row = sheet.createRow(rowNum);

            XSSFCell cell0 = row.createCell(0);
            cell0.setCellValue(waterMeter.getId());// id
            cell0.setCellStyle(rowStyle);

            XSSFCell cell1 = row.createCell(1);
            cell1.setCellValue(waterMeter.getName());// name
            cell1.setCellStyle(rowStyle);

            XSSFCell cell2 = row.createCell(2);
            cell2.setCellValue(waterMeter.getTime());// time
            cell2.setCellStyle(rowStyle);

            XSSFCell cell3 = row.createCell(3);
            cell3.setCellValue(waterMeter.getData());// data
            cell3.setCellStyle(rowStyle);

            XSSFCell cell4 = row.createCell(4);
            cell4.setCellValue(waterMeter.getStatus());// status
            cell4.setCellStyle(rowStyle);

            rowNum++;
        }
        return workbook;
    }

    // 创建表头
    private void createTitle(XSSFWorkbook workbook, XSSFSheet sheet) {
        XSSFRow row = sheet.createRow(0);

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
        for (int i = 0; i < Sheet.size(); i++) {
            cell = row.createCell(i);
            cell.setCellValue(Sheet.get(i));
            cell.setCellStyle(style);
        }
    }
}
