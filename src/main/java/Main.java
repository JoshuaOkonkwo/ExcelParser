public class Main {
    public static void main(String[] args) {
        String excelPath = "C:\\Users\\josho\\OneDrive\\Documents\\Excel Test\\Parsed.xlsx";
        String folderPath = "C:\\Users\\josho\\OneDrive\\Documents\\Excel Test\\Files";
        String sheetName = "Sheet1";
        int ignoreYear = 9999;

        /*
        ExcelParser p = new ExcelParser(excelPath, folderPath, sheetName, ignoreYear);
        p.parse();
        Or for specific columns
        p.parse(col1, col2, col3)

        WILL OVERWRITE EXISTING CELLS
        */

        ExcelParser p = new ExcelParser(excelPath, folderPath, sheetName, ignoreYear);
        p.parseSafe();
    }
}
