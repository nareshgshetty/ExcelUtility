package com.qa.utils;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelUtility {

    private XSSFWorkbook book;
    private XSSFSheet dataSheet;
    private XSSFRow row;
    private XSSFCell cell;

    /**
     *
     * @param file - expects file name with absolute path
     * @param sheetName - expects name of the sheet to be read
     * @return instance of XSSFSheet
     */
    public XSSFSheet readExcel(String file, String sheetName){

        try{
            FileInputStream excelFile = new FileInputStream(new File(file));
            book = new XSSFWorkbook(excelFile);

        }catch (FileNotFoundException fnfe){
            System.out.println(fnfe.getMessage());
        }catch (IOException ioe){
            ioe.getMessage();
        }
        return dataSheet = book.getSheet(sheetName);
    }

    /**
     *
     * @param sheet - expects instance of XSSFSheet
     * @return number of rows in the sheet
     */
    public int getRowCount(XSSFSheet sheet){
        return sheet.getLastRowNum();
    }

    /**
     *
     * @param sheet - expects instance of XSSFSheet
     * @param rowNum - expects the row number to be read
     * @return data in the row as an instance of XSSFRow
     */
    public XSSFRow readRow(XSSFSheet sheet, int rowNum){
        return sheet.getRow(rowNum);
    }

    /**
     *
     * @param row - expects row as instance of XSSFRow
     * @return number of columns in the given row
     */
    public int getColumnCount(XSSFRow row){
        return row.getPhysicalNumberOfCells();
    }

    /**
     *
     * @param row - expects row to be read as an instance of XSSFRow
     * @return value of each column is read as an object in an ArrayList. Assumes that excel contains
     *              either numeric or string values
     */
    public ArrayList<Object> readRowData(XSSFRow row){
        int columnCount=0;
        if(row!=null){
            columnCount = getColumnCount(row);
        }

        ArrayList<Object> rowData = new ArrayList();
        for(int i=0;i<columnCount;i++){
            XSSFCell cell = row.getCell(i);
            if(cell.getCellTypeEnum()== CellType.STRING){
                //System.out.println(cell.getStringCellValue());
                rowData.add(cell.getStringCellValue());
            }
            if(cell.getCellTypeEnum()==CellType.NUMERIC){
                //System.out.println(cell.getNumericCellValue());
                rowData.add(cell.getNumericCellValue());
            }
            if(cell.getCellTypeEnum()== CellType._NONE){
                rowData.add(null);
            }
            if(cell.getCellTypeEnum() == CellType.BLANK){
                rowData.add("");
            }
        }
        System.out.println("Number of values in list: "+rowData.size());
        return rowData;

    }

    /**
     *
     * @param sheet - Expects the sheet to be read and an instance of XSSFSheet
     * @return All the rows of sheet as an ArrayList of Maps.
     *          Each row of data is converted into HashMap,
     *          with all the row headers as keys and cell values in rows as values
     */
    public ArrayList<Map<String,Object>> readSheetAsList(XSSFSheet sheet){
        ArrayList<Map<String,Object>> sheetData = new ArrayList();
        int rowCount = sheet.getLastRowNum();
        if(rowCount>0){
            ArrayList<String> header = new ArrayList();
            row = sheet.getRow(0);
            int colCount = row.getPhysicalNumberOfCells();
            for(int i=0;i<colCount;i++){
                cell = row.getCell(i);
                if(cell.getCellTypeEnum()== CellType.STRING){
                    header.add(cell.getStringCellValue());
                }
            }

            //System.out.println(header);

            for(int i=1;i<=rowCount;i++){
                Map<String,Object> map = new HashMap<>();
                row = sheet.getRow(i);
                for(int j=0;j<header.size();j++){
                    cell = row.getCell(j);
                    if(cell == null){
                        map.put(header.get(j),null);
                    }else{
                        if(cell.getCellTypeEnum()== CellType._NONE){
                            map.put(header.get(j),null);
                        }
                        if(cell.getCellTypeEnum()== CellType.STRING){
                            map.put(header.get(j),cell.getStringCellValue());
                        }
                        if(cell.getCellTypeEnum() == CellType.NUMERIC){
                            map.put(header.get(j),cell.getNumericCellValue());
                        }
                        if(cell.getCellTypeEnum() == CellType.BLANK){
                            map.put(header.get(j),"");
                        }
                    }
                }
                sheetData.add(map);
            }


        }
        return sheetData;
    }

    /**
     *
     * @param fileName - expects the name of the file to be created, without extension
     *                 Creates file in the current project directory
     *
     * @return name of the file created
     *
     */
    public String createExcelFile(String fileName) throws FileNotFoundException{
        String filePath = System.getProperty("user.dir");
        File fileFolder = new File(filePath);

        String file = fileName+".xlsx";
        String vfile = filePath+File.separator+file;
        System.out.println(vfile);
        new FileOutputStream(new File(vfile));
        return vfile;
    }


    /**
     *
     * @param fileName - expects the name of the file to be created, without extension
     *                 Appends timestamp and creates file <fileName>_<timeStamp>.xlsx
     *                 Creates file in the current project directory
     *
     * @return name of the file created
     *
     */
    public String createExcelFileWithTimeStamp(String fileName) throws FileNotFoundException{
        String filePath = System.getProperty("user.dir");
        File fileFolder = new File(filePath);
        String date = new SimpleDateFormat("yyyyMMddHHmm").format(new Date());

        String file = fileName+"_"+date+".xlsx";
        String vfile = filePath+File.separator+file;
        System.out.println(vfile);
        new FileOutputStream(new File(vfile));
        return vfile;
    }

    /**
     *
     * @param fileName - expects absolute path of the file, to which data is to be written
     * @param rowList - expects the data to be written in the row
     *                Data will be written in the row after the last existing row
     *
     */
    public void writeRow(String fileName, ArrayList<Object> rowList) throws IOException,InvalidFormatException {
        File oFile = new File(fileName);
        FileInputStream input = new FileInputStream(oFile);
        book = new XSSFWorkbook(input);
        dataSheet = book.getSheet("TestResults");
        //Sheet sheet = workbook.getSheetAt(0);

        int rowNum = dataSheet.getLastRowNum();
        System.out.println("Last Row: "+rowNum);
        XSSFRow row = dataSheet.createRow(++rowNum);
        System.out.println("New Row: "+row);
        int cellNumber = 0;
        for(Object cellVal : rowList){
            System.out.println("Current Cell: "+ cellNumber);
            XSSFCell cell = row.createCell(cellNumber);
            System.out.println(cell+"-"+cellNumber);
            if(cellVal instanceof String){
                cell.setCellValue((String) cellVal);
            } else {
                cell.setCellValue("String");
            }
            cellNumber++;
        }
        input.close();

        FileOutputStream fos = new FileOutputStream(oFile);
        book.write(fos);
        book.close();
        fos.close();
    }

    /**
     *
     * @param fileName - expects absolute path of the file, to which data is to be written
     * @param headerList - expects all the column names in the file, as a List of String
     * @return name of the file with absolute path, in which the header is written.
     */

    public String writeHeader(String fileName, List<String> headerList) throws IOException {
        FileOutputStream fos = new FileOutputStream(new File(fileName));
        XSSFWorkbook book = new XSSFWorkbook();
        XSSFSheet sheet = book.createSheet("TestResults");
        Row row = sheet.createRow(0);
        int cellNumber = 0;
        for(String header : headerList){
            Cell cell = row.createCell(cellNumber++);
            cell.setCellValue(header);
        }
        book.write(fos);
        book.close();
        fos.close();
        return fileName;
    }
}
