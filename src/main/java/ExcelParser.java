import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;


public class ExcelParser {

    //Use \\ instead of \ for file escape sequences

    private final String EXCEL_PATH;
    private final String FOLDER_PATH;
    private final String SHEET_NAME;
    private final int IGNORE_YEAR; // Ignores

    public ExcelParser(String excelPath, String folderPath, String sheetName) {
        EXCEL_PATH = excelPath;
        FOLDER_PATH = folderPath;
        SHEET_NAME = sheetName;
        IGNORE_YEAR = 9999;
    }

    public ExcelParser(String excelPath, String folderPath, String sheetName, int ignoreYear) {
        EXCEL_PATH = excelPath;
        FOLDER_PATH = folderPath;
        SHEET_NAME = sheetName;
        IGNORE_YEAR = ignoreYear;
    }

    /**
     * Parses file name, year, and extension into the class Excel sheet for each file.
     *
     * Uses default columns and avoids overwriting cells.
     *
     */
    public void parseSafe() {
        parseSafe(4, 5, 7);
    }

    /**
     * Parses file name, year, and extension into the class Excel sheet for each file.
     *
     * Uses specified columns and avoids overwriting cells.
     *
     * @param col1 column to write file name into
     * @param col2 column to write year into
     * @param col3 column to write file extension into
     */
    public void parseSafe(int col1, int col2, int col3) {
        try {
            FileInputStream fis = new FileInputStream(EXCEL_PATH);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(SHEET_NAME);
            if (sheet == null) {
                System.err.println("Sheet " + SHEET_NAME + " not found");
                return;
            }

            File folder = new File(FOLDER_PATH);
            if (!folder.exists() || !folder.isDirectory()) {
                System.err.println("Invalid path: " + FOLDER_PATH);
                return;
            }

            Stack<File> stack = new Stack<>();
            stack.push(folder);
            int rowNum = sheet.getLastRowNum() + 1;
            parseHelper(col1, col2, col3, fis, workbook, sheet, stack, rowNum);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Parses file name, year, and extension into the class Excel sheet for each file
     *
     * Uses default columns and starts from specified row.
     *
     * WARNING: WILL OVERWRITE EXISTING CELLS IF param startingRow SPECIFIES IT
     *
     * @param startingRow row to start writing from
     */
    public void parse(int startingRow) {
        int col1 = 4, col2 = 5, col3 = 7;
        parse(col1, col2, col3, startingRow);
    }

    /**
     * Parses file name, year, and extension into the class excel sheet for each file
     *
     * Uses default columns and starts from specified row.
     *
     * @param col1 column to write file name into
     * @param col2 column to write year into
     * @param col3 column to write file extension into
     * @param startingRow row to start writing from
     */
    public void parse(int col1, int col2, int col3, int startingRow) {
        try {
            FileInputStream fis = new FileInputStream(EXCEL_PATH);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(SHEET_NAME);
            if (sheet == null) {
                System.err.println("Sheet " + SHEET_NAME + " not found");
                return;
            }

            File folder = new File(FOLDER_PATH);
            if (!folder.exists() || !folder.isDirectory()) {
                System.err.println("Invalid path: " + FOLDER_PATH);
                return;
            }

            Stack<File> stack = new Stack<>();
            stack.push(folder);
            parseHelper(col1, col2, col3, fis, workbook, sheet, stack, startingRow);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }



    private String getDeliverableName(String fileName) {
        int lastDot = fileName.lastIndexOf('.');
        return (lastDot > 0) ? fileName.substring(0, lastDot) : fileName;
    }

    private String getFileExtension(String fileName) {
        String[] excelExtensions = {"XLS", "XSLX", "XLSM", "XLAM"};
        String[] wordExtensions = {"DOC", "DOCX"};
        String[] powerpointExtensions = {"PPTX", "PPTM", "PPT"};
        int lastDot = fileName.lastIndexOf('.');
        String extension = (lastDot > 0) ? fileName.substring(lastDot + 1).toUpperCase() : "";

        if (Arrays.asList(excelExtensions).contains(extension)) {
            return "MS Excel";
        } else if (Arrays.asList(wordExtensions).contains(extension)) {
            return "MS Word";
        } else if (Arrays.asList(powerpointExtensions).contains(extension)) {
            return "MS Powerpoint";
        } else if (extension.equals("MSG")) {
            return "Outlook Item";
        } else {
            return extension;
        }

    }

    private void parseHelper(int col1, int col2, int col3, FileInputStream fis, Workbook workbook, Sheet sheet, Stack<File> stack, int rowNum) throws IOException {
        int filesParsed = 0;
        while (!stack.isEmpty()) {
            File currentFile = stack.pop();
            if (currentFile.isDirectory()) {
                for (File f : currentFile.listFiles()) {
                    stack.push(f);
                }
            } else if (currentFile.getName().toLowerCase().endsWith(".zip")) {
                ZipFile zipFile = new ZipFile(currentFile);
                Enumeration<? extends ZipEntry> entries = zipFile.entries();
                while (entries.hasMoreElements()) {
                    ZipEntry entry = entries.nextElement();

                    long timeMillis = entry.getTime();
                    Calendar cal = Calendar.getInstance();
                    cal.setTimeInMillis(timeMillis);
                    int year = cal.get(Calendar.YEAR);

                    if (year > IGNORE_YEAR) continue;

                    String fileName = entry.getName();
                    String deliverable = getDeliverableName(fileName);
                    String docType = getFileExtension(fileName);

                    Row row = sheet.createRow(rowNum++);
                    row.createCell(col1).setCellValue(String.valueOf(year));  //E
                    row.createCell(col2).setCellValue(deliverable); //F
                    row.createCell(col3).setCellValue(docType); // H
                    System.out.println("Parsed File " + ++filesParsed + ": "  + fileName);
                }
                zipFile.close();
            } else {
                Calendar cal = Calendar.getInstance();
                cal.setTime(new Date(currentFile.lastModified()));
                int year = cal.get(Calendar.YEAR);

                if (year > IGNORE_YEAR) continue;

                String fileName = currentFile.getName();
                String deliverable = getDeliverableName(fileName);
                String docType = getFileExtension(fileName).toUpperCase();

                Row row = sheet.createRow(rowNum++);
                row.createCell(col1).setCellValue(year);  //E
                row.createCell(col2).setCellValue(deliverable); //F
                row.createCell(col3).setCellValue(docType); // H
                System.out.println("Parsed File " + ++filesParsed + ": "  + fileName);
            }

        }


        fis.close();
        FileOutputStream fos = new FileOutputStream(EXCEL_PATH);
        workbook.write(fos);
        fos.close();
        workbook.close();

        System.out.println("Complete. Parsed " + filesParsed + " files.");
    }
}
