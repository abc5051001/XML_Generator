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

public class Xls_Reader {
    public  String path;
    public  FileInputStream fis = null;
    public  FileOutputStream fileOut =null;
    private XSSFWorkbook workbook = null;
    private XSSFSheet sheet = null;
    private XSSFRow row   =null;
    private XSSFCell cell = null;
   
    public Xls_Reader(String path) {
       
        this.path=path;
        try {
            //ZipSecureFile.setMinInflateRatio(0.0d);
            fis = new FileInputStream(path);
            workbook = new XSSFWorkbook(fis);
            sheet = workbook.getSheetAt(0);
            fis.close();
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
   
   
    // returns the row count in a specific column in the sheet
    public int getRowCountOfaColumn(String sheetName, String colName) {
        int index = workbook.getSheetIndex(sheetName);
        int col_Num = -1;
        sheet = workbook.getSheetAt(index);
        row = sheet.getRow(0);
        for (int i = 0; i < row.getLastCellNum(); i++) {
            if (row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
                col_Num = i;
        }
        int number = 0;
        int rowCount = sheet.getLastRowNum() + 1;
        for (int i = 0; i <= rowCount; i++) {
            if(sheet.getRow(i).getCell(col_Num) != null) {
                number += 1;
            }
            else {
                break;
            }
        }
        return number;
    }
   
    // returns the row count in a sheet
    public int getRowCount(String sheetName){
        int index = workbook.getSheetIndex(sheetName);
        if(index==-1)
            return 0;
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
            if (index == -1)
                return "";
            sheet = workbook.getSheetAt(index);
            row = sheet.getRow(0);
            for (int i = 0; i < row.getLastCellNum(); i++) {
                if (row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
                    col_Num = i;
            }
            if (col_Num == -1)
                return "";
            sheet = workbook.getSheetAt(index);
            row = sheet.getRow(rowNum - 1);
            if (row == null)
                return "";
            cell = row.getCell(col_Num);
            if (cell == null)
                return "";
            //System.out.println(cell.getCellTypeEnum());
            if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                //getCellTypeEnum
                return cell.getStringCellValue();
            }
            else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC || cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
//              System.out.println("Formula is " + cell.getCellFormula());
//              System.out.println(cell.getNumericCellValue());
                DataFormatter formatter = new DataFormatter();
                String var_name = formatter.formatCellValue(cell);
                //System.out.println(var_name);
                String cellText = String.valueOf(var_name);
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    // format in form of M/D/YY
                    double d = cell.getNumericCellValue();
                    Calendar cal = Calendar.getInstance();
                    cal.setTime(HSSFDateUtil.getJavaDate(d));
                    cellText = (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
                    cellText = cal.get(Calendar.DAY_OF_MONTH) + "/" + cal.get(Calendar.MONTH) + 1 + "/" + cellText;
                }
                //return "pass";
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
            return "row " + rowNum + " or column " + colName + " does not exist in xls";
        }
    }
   
   
    public String getCellDataFromFormula(String sheetName, String colName, int rowNum) {
        try {
            if (rowNum <= 0)
                return "";
            int index = workbook.getSheetIndex(sheetName);
            int col_Num = -1;
            if (index == -1)
                return "";
            sheet = workbook.getSheetAt(index);
            row = sheet.getRow(0);
            for (int i = 0; i < row.getLastCellNum(); i++) {
                if (row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
                    col_Num = i;
            }
            if (col_Num == -1)
                return "";
            sheet = workbook.getSheetAt(index);
            row = sheet.getRow(rowNum - 1);
            if (row == null)
                return "";
            cell = row.getCell(col_Num);
            if (cell == null)
                return "";
            //System.out.println(cell.getCellTypeEnum());
            if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                //getCellTypeEnum
                return cell.getStringCellValue();
            }
            else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC || cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                String cellText = String.valueOf(cell.getNumericCellValue());
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    // format in form of M/D/YY
                    double d = cell.getNumericCellValue();
                    Calendar cal = Calendar.getInstance();
                    cal.setTime(HSSFDateUtil.getJavaDate(d));
                    cellText = (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
                    cellText = cal.get(Calendar.DAY_OF_MONTH) + "/" + cal.get(Calendar.MONTH) + 1 + "/" + cellText;
                }
                //return "pass";
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
            return "row " + rowNum + " or column " + colName + " does not exist in xls";
        }
    }
       
    public boolean fisClose(String  sheetname){    
       
        FileOutputStream fileOut;
        try {
             workbook.createSheet(sheetname);  
             fileOut = new FileOutputStream(path);
             workbook.write(fileOut);
             fileOut.close();          
        } catch (Exception e) {        
            e.printStackTrace();
            return false;
        }
        return true;
    }
   
 
    // to run this on stand alone
    public static void main(String arg[]) throws IOException{
   
    }  
}