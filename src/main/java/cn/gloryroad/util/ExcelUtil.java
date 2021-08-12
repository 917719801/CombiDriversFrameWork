package cn.gloryroad.util;

import cn.gloryroad.configutation.Constants;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class ExcelUtil {
    private static XSSFSheet ExcelWSheet;
    private static XSSFWorkbook ExcelWBook;
    private static XSSFCell Cell;
    private static XSSFRow Row;

    public static void setExcelFile(String Path){
        FileInputStream ExcelFile;
        try {
            ExcelFile = new FileInputStream(Path);
            ExcelWBook = new XSSFWorkbook(ExcelFile);
        }catch (Exception e){
            System.out.println("Excel路径设置失败");
            e.printStackTrace();
        }
    }


    public static void setExcelFile(String path,String SheetName){
        FileInputStream ExcelFile;
        try{
            ExcelFile = new FileInputStream(path);
            ExcelWBook= new XSSFWorkbook(ExcelFile);
            ExcelWSheet = ExcelWBook.getSheet(SheetName);
        }catch (Exception e){
            System.out.println("Excel路径设置失败");
            e.printStackTrace();
        }
    }

    public static int excelColStrToNum(String colStr){
        int num = 0;
        int result=0;
        char ch = colStr.charAt(0);
        num = (int)(ch-'A');
        result+=num;
        return result;
    }

    public static String getCellData(String SheetName,int RoNum,int ColNum){
        ExcelWSheet= ExcelWBook.getSheet(SheetName);
        try{
            Cell = ExcelWSheet.getRow(RoNum).getCell(ColNum);
            String CellData = Cell.getCellType()== XSSFCell.CELL_TYPE_STRING?Cell.getStringCellValue()+"":String.valueOf(Math.round(Cell.getNumericCellValue()));
            return CellData;
        }catch (Exception e){
            e.printStackTrace();

            return "";
        }
    }

    public static int getLastRowNum(){
        return ExcelWSheet.getLastRowNum();
    }


    public static int getRowCount(String SheetName){
        ExcelWSheet= ExcelWBook.getSheet(SheetName);
        int number=ExcelWSheet.getLastRowNum();
        return number;
    }

    public static int getFirstRowContainsTestCaseId(String sheetName,String testCaseName,int colNum){
        int i;
        try{
            ExcelWSheet = ExcelWBook.getSheet(sheetName);
            int rowCount=ExcelUtil.getRowCount(sheetName);
            for (i=0;i<rowCount;i++){
                if (ExcelUtil.getCellData(sheetName,i,colNum).equals(testCaseName)){
                    break;
                }
            }
            return i;
        }catch (Exception e){
            return 0;
        }
    }

    public static int getTestCaseLastStepRow(String SheetName,String testCaseID,int testCaseStartRowNumber){
        try{
            ExcelWSheet= ExcelWBook.getSheet(SheetName);
            for (int i= testCaseStartRowNumber;i<=ExcelUtil.getRowCount(SheetName)-1;i++){
                if (!testCaseID.equals(ExcelUtil.getCellData(SheetName,i, Constants.TESTCASE_TESTCASEID))){
                    int number =i;
                    return number;
                }
            }
            int number = ExcelWSheet.getLastRowNum()+1;
            return number;
        }catch (Exception e){
            return 0;
        }
    }
    @SuppressWarnings("static-access")
    public static void setCellData(String SheetName,int RowNum,int ColNum,String Result){
        ExcelWSheet = ExcelWBook.getSheet(SheetName);
        try{
            Row=ExcelWSheet.getRow(RowNum);
            Cell = Row.getCell(ColNum, Row.RETURN_BLANK_AS_NULL);
            if (Cell ==null){
                Cell = Row.createCell(ColNum);
                Cell.setCellValue(Result);
            }else {
                Cell.setCellValue(Result);
            }
            FileOutputStream fileOut = new FileOutputStream(Constants.CONBIDRIVER_DATA_EXCEL);
            ExcelWBook.write(fileOut);
            fileOut.flush();
            fileOut.close();
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
