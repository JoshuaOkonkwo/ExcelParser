import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Date;
import java.util.Calendar;

public class ExcelParser {

    private static final String EXCEL_PATH = "INSERT EXCEL PATH";
    private static final String FOLDER_PATH = "INSERT FOLDER PATH";
    private static final String SHEET_NAME = "Sheet1";

    public static void main(String[] args) {
        try {
            FileInputStream fis = new FileInputStream(EXCEL_PATH);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(SHEET_NAME);
            if (sheet == null) {
                System.err.println("Sheet not found: " + SHEET_NAME);
                return;
            }

            File folder = new File(FOLDER_PATH);
            if (!folder.exists() || !folder.isDirectory()) {
                System.err.println("Invalid folder path: " + FOLDER_PATH);
                return;
            }

            int rowNum = sheet.getLastRowNum() + 1;
            for (File file : folder.listFiles()) {
                if (file.isFile()) {
                    Calendar cal = Calendar.getInstance();
                    cal.setTime(new Date(file.lastModified()));
                    String year = String.valueOf(cal.get(Calendar.YEAR));

                    String fileName = file.getName();
                    String deliverable = getDeliverableName(fileName);
                    String docType = getFileExtension(fileName).toUpperCase();

                    Row row = sheet.createRow(rowNum++);
                    row.createCell(4).setCellValue(year);  //E
                    row.createCell(5).setCellValue(deliverable); //F
                    row.createCell(7).setCellValue(docType); // H
                }
            }

            fis.close();
            FileOutputStream fos = new FileOutputStream(EXCEL_PATH);
            workbook.write(fos);
            fos.close();
            workbook.close();

            System.out.println("Excel updated successfully.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getDeliverableName(String fileName) {
        int lastDot = fileName.lastIndexOf('.');
        return (lastDot > 0) ? fileName.substring(0, lastDot) : fileName;
    }

    private static String getFileExtension(String fileName) {
        int lastDot = fileName.lastIndexOf('.');
        return (lastDot > 0) ? fileName.substring(lastDot + 1).toLowerCase() : "";
    }
}
