import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.compress.archivers.zip.ZipFile;
import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;

import java.io.*;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;


public class ExcelParser {

    //Use \\ instead of \ for file escape sequences

    private final String EXCEL_PATH;
    private final String FOLDER_PATH;
    private final String SHEET_NAME;
    private final int IGNORE_YEAR; // Ignores
    private final HashSet<String> zipsParsed = new HashSet<>();
    private final ArrayList<String> newExtensions = new ArrayList<>();
    private final HashSet<String> knownExtensions = new HashSet<>(Arrays.asList(
            "XLS", "XSLX", "XLSM", "XLAM", "DOC", "DOCX", "PPTX", "PPTM", "PPT", "JPG", "PDF", "PNG", "TXT"
    ));

    private int filesParsed;

    public ExcelParser(String excelPath, String folderPath, String sheetName) {
        EXCEL_PATH = excelPath;
        FOLDER_PATH = folderPath;
        SHEET_NAME = sheetName;
        IGNORE_YEAR = 9999;
        filesParsed = 0;
    }

    public ExcelParser(String excelPath, String folderPath, String sheetName, int ignoreYear) {
        EXCEL_PATH = excelPath;
        FOLDER_PATH = folderPath;
        SHEET_NAME = sheetName;
        IGNORE_YEAR = ignoreYear;
        filesParsed = 0;
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
         if (file.getName().toLowerCase().endsWith(".zip")) {
            try (ZipFile zipFile = new ZipFile(file)) {
                zipsParsed.add(file.getName());
                System.out.println("PROCESSING ZIP: " + file.getName() + "\n");
                processZip(zipFile, sheet, col1, col2, col3, rowCounter);
            }
        } else if (file.isDirectory()) {
             System.out.println("PROCESSING FOLDER: " + file.getName() + "\n");
             for (File f : Objects.requireNonNull(file.listFiles())) {
                 processFile(f, sheet, col1, col2, col3, rowCounter);
             }
             System.out.println("FOLDER PROCESSING COMPLETE\n");
         } else {
            writeRow(file.getName(), file.lastModified(), sheet, col1, col2, col3, rowCounter);
        }
    }

    private void processZip(ZipFile zipFile, Sheet sheet, int col1, int col2, int col3, AtomicInteger rowCounter) throws IOException {
        Enumeration<ZipArchiveEntry> entries = zipFile.getEntries();


        while (entries.hasMoreElements()) {
            ZipArchiveEntry entry = entries.nextElement();

            if (entry.isDirectory() || zipsParsed.contains(entry.getName())) continue;

            if (entry.getName().toLowerCase().endsWith(".zip")) {
                System.out.println("PROCESSING NESTED ZIP ENTRY: " + entry.getName() + "\n");

                File tempZip = File.createTempFile("nested-", ".zip");
                tempZip.deleteOnExit();

                try (InputStream is = zipFile.getInputStream(entry)) {

                    try (FileOutputStream fos = new FileOutputStream(tempZip)) {
                        byte[] buffer = new byte[8192];
                        int bytesRead;

                        while ((bytesRead = is.read(buffer)) != -1) {
                            fos.write(buffer, 0, bytesRead);
                        }
                    }

                } catch (IOException e) {
                    throw new IOException("Failure creating nested zip: " + entry.getName(), e);
                }

                try (ZipFile nestedZip = new ZipFile(tempZip)) {
                    zipsParsed.add(entry.getName());
                    processZip(nestedZip, sheet, col1, col2, col3, rowCounter);
                } catch (IOException e) {
                    throw new IOException("Failure opening nested zip: " + entry.getName(), e);
                }


            } else {
                writeRow(entry.getName(), entry.getTime(), sheet, col1, col2, col3, rowCounter);
            }


        }

        System.out.println("ZIP PROCESSING COMPLETE\n");
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

        System.out.println("Parsed File " + ++filesParsed + ": " + filename);
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
