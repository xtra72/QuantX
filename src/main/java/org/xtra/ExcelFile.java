package org.xtra;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;

import org.apache.logging.log4j.Level;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.ss.usermodel.*;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.logging.log4j.Logger;

public abstract class ExcelFile {

    public class XWorkbook implements AutoCloseable {
        Workbook workbook;
        Map<String, CellStyle>  cellStyles;
        Map<String, Font>  fonts;
        Map<String, Short> dataFormats;
        List<XSheet> sheets;

        public XWorkbook() {
            this(new XSSFWorkbook());
        }

        public XWorkbook(Workbook workbook) {
            this.workbook = workbook;
            this.cellStyles = new HashMap<>();
            this.fonts = new HashMap<>();
            this.dataFormats = new HashMap<>();
            this.sheets = new ArrayList<>();
        }

        @Override
        public void close() throws IOException {
            this.workbook.close();
        }

        public XSheet createXSheet() {
            XSheet sheet = new XSheet(this, this.workbook.createSheet());
            this.sheets.add(sheet) ;

            return  sheet;
        }

        public CellStyle createCellStyle(String name) {
            CellStyle   cellStyle = this.workbook.createCellStyle();
            this.cellStyles.put(name, cellStyle);

            return     cellStyle;
        }

        public CellStyle getCellStyle(String name) {
            if (this.cellStyles.containsKey(name)) {
                return  this.cellStyles.get(name);
            }

            throw new NullPointerException();
        }

        public Font createFont(String name) {
            Font font = this.workbook.createFont();
            this.fonts.put(name, font);

            return     font;
        }

        public Font getFont(String name) {
            if (this.fonts.containsKey(name)) {
                return  this.fonts.get(name);
            }

            throw new NullPointerException();
        }

        public short createDataFormat(String name, String pattern) {
            DataFormat dataFormat = this.workbook.createDataFormat();
            short index = dataFormat.getFormat(pattern);

            this.dataFormats.put(name, index);
            return  index;
        }

        public short getDataFormat(String name) {
            if (this.dataFormats.containsKey(name)) {
                return  this.dataFormats.get(name);
            }

            throw new NullPointerException();
        }

        public void write(OutputStream outputStream) throws IOException {
            this.workbook.write(outputStream);
        }
    }

    public class XSheet {
        XWorkbook xWorkbook;
        Sheet sheet;
        Map<Integer, XRow>  xRows;

        public XSheet(XWorkbook workbook, Sheet sheet) {
            this.xWorkbook = workbook;
            this.sheet = sheet;
            this.xRows = new TreeMap<>();
        }

        public XWorkbook getXWorkbook() {
            return  this.xWorkbook;
        }

        public void addMergedRegion(CellRangeAddress cellRangeAddress) {
            this.sheet.addMergedRegion(cellRangeAddress);
        }

        public XRow createXRow(int i) {
            XRow xRow = new XRow(this, this.sheet.createRow(i)) ;
            this.xRows.put(i, xRow);

            return  xRow;
        }

        public Row createRow(int i) {
            return  this.sheet.createRow(i);
        }

        public Row getRow(int i) {
            return  this.sheet.getRow(i) ;
        }

        public int getLastRowNum() {
            return  this.sheet.getLastRowNum();
        }

        public void autoSizeColumn(int c, boolean b) {
            this.sheet.autoSizeColumn(c, b);
        }
    }

    public class XRow {
        XSheet xSheet;
        Row row;

        public XRow(XSheet xSheet, Row row) {
            this.xSheet = xSheet;
            this.row = row;
        }

        public XSheet getXSheet() {
            return  this.xSheet;
        }

        public Cell createCell(int i , CellType cellType) {
            return  this.row.createCell(i, cellType);
        }

        public Cell createCell(int i , CellType cellType, CellStyle cellStyle) {
            Cell cell = this.createCell(i, cellType);
            cell.setCellStyle(cellStyle);

            return  cell;
        }

        public int getLastCellNum() {
            return  this.row.getLastCellNum();
        }
    }
    protected final Logger logger;
    Map<Integer, String> header;
    Sheet      currentSheet;

    ExcelFile() {
        this.header = new HashMap<>();
        this.logger = LogManager.getLogger();
    }
    public void load(String path) {
        try(FileInputStream inStream = new FileInputStream(path)) {
            Workbook workbook = new XSSFWorkbook(inStream) ;

            for(int page = 0 ; page < workbook.getNumberOfSheets() ; page++) {
                this.currentSheet = workbook.getSheetAt(page);
                this.logger.log(Level.DEBUG, "Load page : {}", this.currentSheet.getSheetName());

                this.load(this.currentSheet);
            }
        } catch (FileNotFoundException e) {
            this.logger.log(Level.ERROR, "File not found : {}", path);
        } catch (IOException e) {
            this.logger.log(Level.ERROR, e);
        }
    }

    protected void load(Sheet sheet) {
        this.logger.log(Level.DEBUG, "Sheet : {}", sheet.getSheetName());

        this.header = this.getHeader(sheet);

        System.out.println(this.header);
        for (Row row : sheet) {
            if (this.isData(row)) {
                this.load(row);
            }
        }
    }

    protected abstract void load(Row row);
    protected abstract boolean isHeader(Row row);
    protected abstract boolean isData(Row row);
    protected abstract Map<Integer, String> getHeader(Sheet sheet);

    protected double toPercent(Cell cell) {
        if (cell.getCellType() == CellType.STRING) {
            return  Double.parseDouble(cell.getStringCellValue().replace("%", ""));
        } else {
            return  cell.getNumericCellValue();
        }
    }

    protected String toString(Cell cell) {
        if (cell.getCellType() == CellType.NUMERIC) {
            return  String.valueOf((int)cell.getNumericCellValue());
        } else {
            return  cell.getStringCellValue();
        }
    }

    protected Integer toInteger(Cell cell) {
        if (cell.getCellType() == CellType.NUMERIC) {
            return  (int)cell.getNumericCellValue();
        } else if ((cell.getStringCellValue().length() != 0) && (!cell.getStringCellValue().equals("N/A"))) {
            return (int)Double.parseDouble(cell.getStringCellValue().replace(",", ""));
        }

       return  null;
    }

    protected Long toLong(Cell cell) {
        if (cell.getCellType() == CellType.NUMERIC) {
            return  (long)cell.getNumericCellValue();
        } else if ((cell.getStringCellValue().length() != 0) && (!cell.getStringCellValue().equals("N/A"))) {
            return (long) Double.parseDouble(cell.getStringCellValue().replace(",", ""));
        }

        return  null;
    }

    protected Double toDouble(Cell cell) {
        if (cell.getCellType() == CellType.NUMERIC) {
            return  cell.getNumericCellValue();
        } else if ((cell.getStringCellValue().length() != 0) && (!cell.getStringCellValue().equals("N/A"))) {
            return Double.parseDouble(cell.getStringCellValue().replace(",",""));
        }

        return  null;
    }
}
