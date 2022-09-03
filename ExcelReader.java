import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
    public  String path;
    public  FileInputStream file = null;
    public  FileOutputStream fileOut =null;
    private XSSFWorkbook workbook = null;
    private XSSFSheet sheet = null;
    private XSSFRow row   =null;
    private XSSFCell cell = null;
   
    public ExcelReader(String path) {
       
        this.path=path;
        try {
            file = new FileInputStream(path);
            workbook = new XSSFWorkbook(file);
            sheet = workbook.getSheetAt(0);
            file.close();
        } 
        catch (Exception e) {
            e.printStackTrace();
        }
    }
   
    // returns the row count in a sheet
    public int getRowCount(String sheetName){
        int index = workbook.getSheetIndex(sheetName);
        if( index == -1 ){
            return 0;
        }
        else{
            sheet = workbook.getSheetAt(index);
            int number=sheet.getLastRowNum()+1;
            return number;
        }  
    }
   
    // returns the data from a cell
    public String getCellData(String sheetName, String colName, int rowNum) {
        try {
            if (rowNum <= 0)
                return "";
            int index = workbook.getSheetIndex(sheetName);
            int col_Num = -1;
            // if not found sheet return 
            if (index == -1){
                return "";
            }
            sheet = workbook.getSheetAt(index);
            row = sheet.getRow(0);
            // if found update the num
            for (int i = 0; i < row.getLastCellNum(); i++) {
                if (row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
                    col_Num = i;
            }
            // not found return
            if (col_Num == -1){
                return "";
            }
            
            sheet = workbook.getSheetAt(index);
            row = sheet.getRow(rowNum - 1);
            
            if (row == null)
                return "";
            cell = row.getCell(col_Num);
            if (cell == null)
                return "";
            if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                return cell.getStringCellValue();
            }
            else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC || cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                DataFormatter formatter = new DataFormatter();
                String var_name = formatter.formatCellValue(cell);
                String cellText = String.valueOf(var_name);
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    double d = cell.getNumericCellValue();
                    Calendar cal = Calendar.getInstance();
                    cal.setTime(HSSFDateUtil.getJavaDate(d));
                    cellText = (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
                    cellText = cal.get(Calendar.DAY_OF_MONTH) + "/" + cal.get(Calendar.MONTH) + 1 + "/" + cellText;
                }
                return cellText;
            } else if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
                return "";
            }
            else {
                return String.valueOf(cell.getBooleanCellValue());
            }
        }
        catch (Exception e) {
            e.printStackTrace();
            return "row " + rowNum + " or column " + colName + " does not exist in excel file";
        }
    }
}
