public class Main {
    public static void main(String[] args) {
        String folderPath = "INSERT PATH";
        String excelPath = "INSERT PATH";
        String sheetName = "Sheet1";
        int ignoreYear = 9999;

        /*
        ExcelParser p = new ExcelParser(excelPath, folderPath, sheetName, ignoreYear);
        p.parse(startingRow);
        Or for specific columns
        p.parse(col1, col2, col3, startingRow)

        WILL OVERWRITE EXISTING CELLS. To append, use parseSafe(col1, col2, col3) or parseSafe() for default
        */

        ExcelParser p = new ExcelParser(excelPath, folderPath, sheetName, ignoreYear);
        p.parse(1);
    }
}
