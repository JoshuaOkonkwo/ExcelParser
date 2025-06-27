import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;


public class ExcelParser {

    //Use \\ instead of \ for file escape sequences

    private final String EXCEL_PATH;
    private final String FOLDER_PATH;
    private final String SHEET_NAME;
    private final int IGNORE_YEAR; // Ignores
    private final ArrayList<String> newExtensions = new ArrayList<>();
    private final HashSet<String> knownExtensions = new HashSet<>(Arrays.asList(
            "XLS", "XSLX", "XLSM", "XLAM", "DOC", "DOCX", "PPTX", "PPTM", "PPT", "JPG", "PDF", "PNG", "TXT"
    ));

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
        int col1 = 3, col2 = 4, col3 = 6;
        parseSafe(col1, col2, col3);
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
        int col1 = 3, col2 = 4, col3 = 6;
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


    private void parseHelper(int col1, int col2, int col3, FileInputStream fis, Workbook workbook, Sheet sheet, Stack<File> stack, int rowNum) throws IOException {
        AtomicInteger filesParsed = new AtomicInteger(rowNum);
        while (!stack.isEmpty()) {
            File currentFile = stack.pop();
            processFile(currentFile, sheet, col1, col2, col3, filesParsed);

        }

        fis.close();
        FileOutputStream fos = new FileOutputStream(EXCEL_PATH);
        workbook.write(fos);
        fos.close();
        workbook.close();

        System.out.println("Parsing Complete. Parsed " + filesParsed + " files.");
        System.out.println("Unrecognized extensions:" + newExtensions);
    }

    private void processFile(File file, Sheet sheet, int col1, int col2, int col3, AtomicInteger rowCounter) throws IOException {
        if (file.isDirectory()) {
            for (File f : file.listFiles()) {
                processFile(f, sheet, col1, col2, col3, rowCounter);
            }
        } else if (file.getName().toLowerCase().endsWith(".zip")) {
            try (ZipFile zipFile = new ZipFile(file)) {
                processZip(zipFile, sheet, col1, col2, col3, rowCounter);
            }
        } else {
            writeRow(file.getName(), file.lastModified(), sheet, col1, col2, col3, rowCounter);
        }
    }

    private void processZip(ZipFile zipfile, Sheet sheet, int col1, int col2, int col3, AtomicInteger rowCounter) {
        Enumeration<? extends ZipEntry> entries = zipfile.entries();

        while (entries.hasMoreElements()) {
            ZipEntry entry = entries.nextElement();

            if (entry.isDirectory()) continue;

            if (entry.getName().toLowerCase().endsWith(".zip")) {
                System.out.println("Skipped Nested Zip Entry: " + entry.getName());
            }

            writeRow(entry.getName(), entry.getTime(), sheet, col1, col2, col3, rowCounter);
        }
    }

    private void writeRow(String filename, long timeMillis, Sheet sheet, int col1, int col2, int col3, AtomicInteger rowCounter) {
        Calendar calendar = Calendar.getInstance();
        calendar.setTimeInMillis(timeMillis);
        int year = calendar.get(Calendar.YEAR);

        if (year > IGNORE_YEAR) return;

        String deliverableName = getDeliverableName(filename);
        String docType = getDocType(filename);

        Row row = sheet.createRow(rowCounter.getAndIncrement());

        row.createCell(col1).setCellValue(deliverableName);
        row.createCell(col2).setCellValue(docType);
        row.createCell(col3).setCellValue(year);

        System.out.println("Parsed File " + rowCounter.get() + ": " + filename);
    }

    private String getDeliverableName(String fileName) {
        int lastDot = fileName.lastIndexOf('.');
        return (lastDot > 0) ? fileName.substring(0, lastDot) : fileName;
    }

    private String getDocType(String fileName) {
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

            if (!knownExtensions.contains(extension)) {
                newExtensions.add(extension);
                knownExtensions.add(extension);
            }
            return extension;
        }

    }
}
