import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

import java.io.File;
import java.io.FileInputStream;

public class ExcelToJSonConvertor {
    private JSONObject readExcelFileAsJsonObject_RowWise(String filePath) {
        DataFormatter dataFormatter = new DataFormatter();
        JSONObject workbookJson = new JSONObject();
        JSONArray sheetJson = new JSONArray();
        JSONObject rowJson = new JSONObject();
        try {

            FileInputStream excelFile = new FileInputStream(new File(filePath));
            Workbook workbook = new XSSFWorkbook(excelFile);
            FormulaEvaluator formulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);

            for (Sheet sheet : workbook) {
                sheetJson = new JSONArray();
                int lastRowNum = sheet.getLastRowNum();
                int lastColumnNum = sheet.getRow(0).getLastCellNum();
                Row firstRowAsKeys = sheet.getRow(0); // first row as a json keys

                for (int i = 1; i <= lastRowNum; i++) {
                    rowJson = new JSONObject();
                    Row row = sheet.getRow(i);

                    if (row != null) {
                        for (int j = 0; j < lastColumnNum; j++) {
                            formulaEvaluator.evaluate(row.getCell(j));
                            rowJson.put(firstRowAsKeys.getCell(j).getStringCellValue(),
                                    dataFormatter.formatCellValue(row.getCell(j), formulaEvaluator));
                        }
                        sheetJson.add(rowJson);
                    }
                }
                workbookJson.put(sheet.getSheetName(), sheetJson);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return workbookJson;
    }

    public static void main(String arg[]) {
        ExcelToJSonConvertor excelConvertor = new ExcelToJSonConvertor();
        String filePath = "C:\\Users\\username\\Documents\\WorkInputFiles\\InputSheet.xlsx";
        JSONObject data = excelConvertor.readExcelFileAsJsonObject_RowWise(filePath);
        System.out.println(data);
    }
}

