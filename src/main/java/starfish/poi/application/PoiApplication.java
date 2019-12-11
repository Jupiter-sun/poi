package starfish.poi.application;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import starfish.poi.infrastructure.utils.EasyPoi;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

@Slf4j
@RestController
@RequestMapping(value = "poi")
public class PoiApplication {

    private static final String FILE_NAME = "LzExcel.xlsx";
    @Autowired
    public PoiApplication(EasyPoi easyPoi) {
        this.easyPoi = easyPoi;
    }

    private EasyPoi easyPoi;

    @PostMapping("json2excel")
    public void JsonParseExcel(@RequestBody List<Map<String,String>> array) throws IOException {
        XSSFWorkbook workbook = easyPoi.createWorkBook(array);
        log.info("生成excel......");
        workbook.write(new FileOutputStream(new File(FILE_NAME)));
        workbook.close();
    }

}
