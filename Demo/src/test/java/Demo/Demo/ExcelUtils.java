package Demo.Demo;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class ExcelUtils {

    public static Map<String, String> readExcel(String filePath, String sheetName) throws IOException {
        Map<String, String> dataMap = new HashMap<>();
        FileInputStream fileInputStream = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheet(sheetName);

        for (Row row : sheet) {
            Cell keyCell = row.getCell(0); // Assuming key is in the first column
            Cell valueCell = row.getCell(1); // Assuming value is in the second column
            if (keyCell != null && valueCell != null) {
                dataMap.put(keyCell.toString(), valueCell.toString());
            }
        }

        workbook.close();
        return dataMap;
    }
}
