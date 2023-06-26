package org.xtra;

import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.*;
import java.util.List;
import java.util.concurrent.atomic.AtomicReference;
import java.util.stream.Collectors;

import org.apache.logging.log4j.Level;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.json.JSONObject;

import javax.management.InvalidAttributeValueException;

public class FullCatalog extends ExcelFile {
    static class HeaderField {
        private final String name;
        private final int firstRow;
        private final int lastRow;
        private final int firstCol;
        private final int lastCol;
        private final CellType type;

        public HeaderField(String name, int firstRow, int lastRow, int firstCol, int lastCol, CellType type) {
            this.name = name;
            this.firstRow = firstRow;
            this.lastRow = lastRow;
            this.firstCol = firstCol;
            this.lastCol = lastCol;
            this.type = type;
        }
    }

    static class DataRow {

        private final Row row;

        public DataRow(Row row) {
            this.row = row;
        }

        public String getItemCode() {
            Cell cell = this.row.getCell(FullCatalog.ITEM_CODE_INDEX);
            if (cell == null) {
                return "";
            }

            return cell.getStringCellValue();
        }

        private String getCellString(int column) {
            Cell cell = this.row.getCell(column);
            if (cell == null) {
                return "";
            }

            return cell.getStringCellValue();
        }

        public String getMarket() {
            return  this.getCellString(FullCatalog.MARKET_INDEX);
        }

        public String getName() {
            return  this.getCellString(FullCatalog.NAME_INDEX);
        }

        public String getIndustryCode() {
            return  this.getCellString(FullCatalog.INDUSTRY_INDEX);
        }

        public String getCategory() {
            return  this.getCellString(FullCatalog.CATEGORY_INDEX);
        }

        public String getComment() {
            return  this.getCellString(FullCatalog.COMMENT_INDEX);
        }
    }

    public static final String TITLE_ASSETS = "자산";
    public static final String TITLE_CAPITAL = "자본";
    public static final String TITLE_COMMENT= "설명";
    public static final String TITLE_INCOME = "순익";
    public static final String TITLE_LIABILITIES = "부채";
    public static final String TITLE_MARKET_CAPITALIZATION = "시총";
    public static final String TITLE_PBR = "PBR";
    public static final String TITLE_QUARTERLY_ROE = "ROE/4";
    public static final String TITLE_QUARTERLY_SALES = "SALES/4";
    public static final String TITLE_RATING = "등급";
    public static final String TITLE_ROE = "ROE";
    public static final String TITLE_YEARLY_SALES = "SALES";
    public static final String TITLE_SALES = "매출";
    public static final String TITLE_STOCK = "발행주식수";

    public static final String ITEM_ASSETS = "ASSETS";
    public static final String ITEM_CAPITAL = "CAPITAL";
    public static final String ITEM_COMMENT = "COMMENT";
    public static final String ITEM_INCOME = "INCOME";
    public static final String ITEM_LIABILITIES = "LIABILITIES";
    public static final String ITEM_MARKET_CAPITALIZATION = "MARKET_CAPITALIZATION";
    public static final String ITEM_PBR = "PBR";
    public static final String ITEM_QUARTERLY_ROE = "QUARTERLY_ROE";
    public static final String ITEM_QUARTERLY_SALES = "QUARTERLY_SALES";
    public static final String ITEM_QUARTERLY_INCOME= "QUARTERLY_INCOME";
    public static final String ITEM_RATING = "RATING";
    public static final String ITEM_ROE = "ROE";
    public static final String ITEM_YEARLY_SALES = "YEAR_SALES";
    public static final String ITEM_SALES = "SALES";
    public static final String ITEM_STOCK = "STOCK";

    public static final String CELL_STYLE_PBR = "pbr";
    public static final String CELL_STYLE_ROE = "roe";
    public static final String CELL_STYLE_CURRENCY = "currency";
    public static final String CELL_STYLE_NEGATIVE_CURRENCY = "negativeCurrency";
    public static final String CELL_STYLE_HEADER = "header";
    public static final String CELL_STYLE_NEGATIVE_ROE = "negativeRoe";

    private static final Map<String, String> accountMapper = Map.ofEntries(
            Map.entry(TITLE_ASSETS, ITEM_ASSETS),
            Map.entry(TITLE_CAPITAL, ITEM_CAPITAL),
            Map.entry(TITLE_COMMENT, ITEM_COMMENT),
            Map.entry(TITLE_INCOME, ITEM_INCOME),
            Map.entry(TITLE_LIABILITIES, ITEM_LIABILITIES),
            Map.entry(TITLE_MARKET_CAPITALIZATION, ITEM_MARKET_CAPITALIZATION),
            Map.entry(TITLE_QUARTERLY_ROE, ITEM_QUARTERLY_ROE),
            Map.entry(TITLE_QUARTERLY_SALES, ITEM_QUARTERLY_SALES),
            Map.entry(TITLE_RATING, ITEM_RATING),
            Map.entry(TITLE_ROE, ITEM_ROE),
            Map.entry(TITLE_YEARLY_SALES, ITEM_SALES),
            Map.entry(TITLE_SALES, ITEM_SALES)
    );

    private static final HeaderField[] headerFields = {
            new HeaderField("코드", 0, 2, 0 , 0, CellType.STRING),
            new HeaderField("이름", 0, 2, 1 , 1, CellType.STRING),
            new HeaderField("산업", 0, 2, 2 , 2, CellType.STRING),
            new HeaderField("산업코드", 0, 2, 3 , 3, CellType.STRING),
            new HeaderField("시장", 0, 2, 4 , 4, CellType.STRING),
            new HeaderField("설명", 0, 2, 5 , 5, CellType.STRING)
    };

    public static final int ITEM_CODE_INDEX = 0;
    public static final int NAME_INDEX = 1;
    public static final int CATEGORY_INDEX = 2;
    public static final int INDUSTRY_INDEX = 3;
    public static final int MARKET_INDEX = 4;
    public static final int COMMENT_INDEX = 5;
    public static final int START_OF_DATA_INDEX = 6;
    public static final String MARKET_KOSPI = "KS";
    public static final String MARKET_KOSDAQ = "KQ";

    private final Map<String, Company>    companies;

    private final String[] acceptedAccounts = {ITEM_RATING, ITEM_ROE, ITEM_YEARLY_SALES, ITEM_PBR,
        ITEM_QUARTERLY_ROE, ITEM_QUARTERLY_SALES, ITEM_MARKET_CAPITALIZATION,
        ITEM_INCOME, ITEM_CAPITAL, ITEM_SALES};

    private final long defaultUnit = 1000;

    public FullCatalog(Map<String, Company> companies) {
        this.companies = companies;
    }

    @Override
    protected Map<Integer, String> getHeader(Sheet sheet) {
        Map<Integer, String>  titles = new TreeMap<>();

        Row header1 = sheet.getRow(0);
        Row header2 = sheet.getRow(1);
        Row header3 = sheet.getRow(2);

        String [] titles2 = new String[header1.getLastCellNum()];
        header1.forEach(cell-> {
            String value;
            if ( cell.getCellType() == CellType.STRING) {
                value = cell.getStringCellValue();
            } else {
                value = String.valueOf((int)cell.getNumericCellValue());
            }

            if ((value.length() != 0) && !value.equals("0")) {
                titles2[cell.getColumnIndex()] = value;
            } else {
                titles2[cell.getColumnIndex()] = titles2[cell.getColumnIndex() - 1];
            }

        });

        String [] titles3 = new String[header2.getLastCellNum()];
        header2.forEach(cell -> {
            String title = this.toString(cell);
            if ((titles2[cell.getColumnIndex()].equals(TITLE_QUARTERLY_ROE)
                    || titles2[cell.getColumnIndex()].equals(TITLE_QUARTERLY_SALES)
                    || titles2[cell.getColumnIndex()].equals(TITLE_RATING))
                    && (title.length() == 0)) {
                titles3[cell.getColumnIndex()] = titles3[cell.getColumnIndex() - 1];
            } else {
                titles3[cell.getColumnIndex()] = title;
            }
        });


        String [] titles4 = new String[header3.getLastCellNum()];
        header3.forEach(cell -> titles4[cell.getColumnIndex()] = this.toString(cell));

        for(int i = 0 ; i < titles2.length ; i++) {
            if (titles2[i].matches("[0-9]+")) {
                titles.put(i, FullCatalog.accountMapper.get(titles4[i]) + "-" + titles2[i]);
            } else {
                String title = titles2[i];
                String quarter = "";

                if (titles3[i] != null) {
                    quarter += titles3[i];
                }

                if (titles4[i] != null) {
                    if (titles4[i].length() == 1) {
                        quarter += "0" + titles4[i];
                    } else {
                        quarter += titles4[i];
                    }
                }

                if (quarter.length() != 0) {
                    if (title.matches("\\d+")) {
                        title = FullCatalog.accountMapper.getOrDefault(quarter, quarter) + "-" + title;
                    } else {
                        title = FullCatalog.accountMapper.getOrDefault(title, title) + "-" + quarter;
                    }
                } else {
                    title = FullCatalog.accountMapper.getOrDefault(title, title);
                }

                titles.put(i, title);
            }
        }

        return  titles;
    }

    @Override
    protected boolean isHeader(Row row) {
        return  (row.getRowNum() < 3);
    }

    @Override
    protected boolean isData(Row row) {
        return  ((row.getFirstCellNum() == 0) && (row.getLastCellNum() > FullCatalog.MARKET_INDEX)
                && (row.getCell(FullCatalog.MARKET_INDEX).getCellType() == CellType.STRING)
                && (row.getCell(FullCatalog.MARKET_INDEX).getStringCellValue().equals(FullCatalog.MARKET_KOSDAQ)
                || (row.getCell(FullCatalog.MARKET_INDEX).getStringCellValue().equals(FullCatalog.MARKET_KOSPI))));
    }

    @Override
    protected void load(Row row) {
        DataRow dataRow = new DataRow(row);

        String itemCode = dataRow.getItemCode();

        Company company;

        if (this.companies.containsKey(itemCode)) {
            company = this.companies.get(itemCode);
            if (!company.getName().equals(dataRow.getName())) {
                company.setName(dataRow.getName());
            }
        } else {
            company = new Company(dataRow.getMarket(), itemCode, dataRow.getName(),
                    dataRow.getIndustryCode(), dataRow.getCategory());
            this.companies.put(itemCode, company);
        }

        company.setComment(dataRow.getComment());

        row.forEach(cell -> {
            if (FullCatalog.START_OF_DATA_INDEX <= cell.getColumnIndex() && header.containsKey(cell.getColumnIndex())) {
                String[] fields = header.get(cell.getColumnIndex()).split("-");
                if (Arrays.stream(this.acceptedAccounts).anyMatch(account -> account.equals(fields[0]))) {
                    try {
                        switch (fields[0]) {
                            case ITEM_SALES:
                            case ITEM_INCOME:
                            case ITEM_CAPITAL: {
                                if (!company.hasAccount(fields[0], Long.parseLong(fields[1]))) {
                                    Long sales = this.toLong(cell);
                                    if (sales != null)
                                        company.setAccount(fields[0], fields[1], sales);
                                }
                            }
                            break;

                            case ITEM_YEARLY_SALES:
                            case ITEM_QUARTERLY_SALES: {
                                if (!company.hasAccount(fields[0], Long.parseLong(fields[1]))) {
                                    Long sales = toDefaultUnit(this.toLong(cell));
                                    if (sales != null)
                                        company.setAccount(fields[0], fields[1], sales);
                                }
                            }
                            break;

                            case ITEM_ROE:
                            case ITEM_QUARTERLY_ROE: {
                                if (!company.hasAccount(fields[0], Long.parseLong(fields[1]))) {
                                    Double roe = this.toPercent(cell);
                                    if (roe != null) {
                                        company.setAccount(fields[0], fields[1], roe);
                                    }
                                }
                            }
                            break;

                            case ITEM_RATING: {
                                if (fields[1].matches("\\d\\d")) {
                                    fields[1] = String.valueOf((2020 + Long.parseLong(fields[1])/10)*100 +  Long.parseLong(fields[1]) % 10 * 3);
                                } else if (fields[1].matches("\\d\\d\\d\\d\\d")) {
                                    fields[1] = String.valueOf(Long.parseLong(fields[1])/10*100 +  Long.parseLong(fields[1]) % 10 * 3);
                                }

                                company.setAccount(fields[0], fields[1], cell.getStringCellValue());
                            }
                            break;

                            case ITEM_PBR: {
                                if (cell.getCellType() == CellType.NUMERIC)
                                    company.setPbr(cell.getNumericCellValue());
                            }
                            break;

                            case ITEM_MARKET_CAPITALIZATION: {
                                if (cell.getCellType() == CellType.NUMERIC)
                                    company.setMarketCapitalization((long) cell.getNumericCellValue());
                                else if (cell.getCellType() == CellType.STRING)
                                    company.setMarketCapitalization((long) Long.parseLong(cell.getStringCellValue().replaceAll(",", "")));
                            }
                            break;

                            default: {
                                Long value = this.toLong(cell);
                                if (value != null) {
                                    company.setAccount(fields[1], fields[0], value);
                                }
                            }
                        }
                    } catch (InvalidAttributeValueException | NumberFormatException e) {
                        this.logger.log(Level.ERROR, "Invalid value for {}-{}-{} : {}", itemCode, fields[0], fields[1], e);
                    } catch (Exception e) {
                        this.logger.log(Level.ERROR, "Invalid value for {}-{}-{} : {}", itemCode, fields[0], fields[1], e);
                    }
                }
            }
        });

        this.companies.put(itemCode, company);
    }

    public void calculate() {
        this.companies.forEach((code, company)->{
            company.calculate();
//
//            List<Long> quarters = company.getQuarterlies().stream().sorted().collect(Collectors.toList());
//            quarters.forEach(quarter -> {
//                try {
//                    Finance finance = company.getFinance(quarter / 100, quarter % 100);
//
//                    if ((quarter % 100 / 3) == 1) {
//                        if (!finance.hasAccount(ITEM_QUARTERLY_SALES) && finance.hasAccount(ITEM_SALES)) {
//                            finance.setAccount(ITEM_QUARTERLY_SALES, (long) finance.getAccount(ITEM_SALES) / 100000);
//                        }
//
//                        if (!finance.hasAccount(ITEM_QUARTERLY_INCOME) && finance.hasAccount(ITEM_INCOME)) {
//                            finance.setAccount(ITEM_QUARTERLY_INCOME, (long) finance.getAccount(ITEM_INCOME));
//                        }
//
//
//                    } else {
//                        long previousQuarterly = quarter - 3;
//
//                        Finance previousFinance = company.getFinance(previousQuarterly / 100, previousQuarterly % 100);
//
//                        if (!finance.hasAccount(ITEM_QUARTERLY_SALES)) {
//                            long sales = (long) finance.getAccount(ITEM_SALES);
//                            sales -= (long) previousFinance.getAccount(ITEM_SALES);
//                            finance.setAccount(ITEM_QUARTERLY_SALES, sales / 100000);
//                        }
//
//                        if (!finance.hasAccount(ITEM_QUARTERLY_INCOME)) {
//                            long income = (long) finance.getAccount(ITEM_INCOME);
//                            income -= (long) previousFinance.getAccount(ITEM_INCOME);
//                            finance.setAccount(ITEM_QUARTERLY_INCOME, income);
//                        }
//                    }
//
//                    if (!finance.hasAccount(ITEM_QUARTERLY_ROE)) {
//                        if ((Long)finance.getAccount(ITEM_QUARTERLY_INCOME) < 0L && (Long)finance.getAccount(ITEM_CAPITAL) < 0L) {
//                            finance.setAccount(ITEM_QUARTERLY_ROE, 0.0D);
//                        } else {
//                            finance.setAccount(ITEM_QUARTERLY_ROE,
//                                ((Long) finance.getAccount(ITEM_QUARTERLY_INCOME) * 1.0 / (Long) finance.getAccount(ITEM_CAPITAL)));
//                        }
//                    }
//
//                } catch (NullPointerException e) {
//                    this.logger.log(Level.WARN, e);
//                }
//            });
        });
    }

    public void save(String filename) {
        try(FileWriter output = new FileWriter(filename)) {
            output.write(this.toJson().toString(4));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void createColumn(XSheet xSheet, String title, int row, int rowSize, int column, int columnSize) {

        xSheet.addMergedRegion(new CellRangeAddress(row, row+rowSize - 1, column, column + columnSize - 1));
        for(int i = 0 ; i < rowSize ; i++) {
            xSheet.getRow(row + i).createCell(column, CellType.STRING).setCellValue(title);
        }
        xSheet.getRow(row).getCell(column).setCellStyle(xSheet.getXWorkbook().getCellStyle(CELL_STYLE_HEADER));
    }

    public void createYearlyColumns(XSheet xSheet, String title, int row, int column, Set<Long> years) {
        final int[] tempColumn = { column };

        xSheet.addMergedRegion(new CellRangeAddress(row, row, column, column + years.size() - 1));
        years.stream()
                .sorted(Comparator.reverseOrder())
                .forEach(year->{
                    xSheet.getRow(row).createCell(tempColumn[0], CellType.STRING).setCellValue(title);
                    xSheet.getRow(row).getCell(tempColumn[0]).setCellStyle(xSheet.getXWorkbook().getCellStyle(CELL_STYLE_HEADER));

                    xSheet.getRow(row+1).createCell(tempColumn[0], CellType.STRING).setCellValue(String.valueOf(year));
                    xSheet.getRow(row+1).getCell(tempColumn[0]).setCellStyle(xSheet.getXWorkbook().getCellStyle(CELL_STYLE_HEADER));
                    xSheet.addMergedRegion(new CellRangeAddress(row+1, row+2, tempColumn[0], tempColumn[0]));
                    tempColumn[0] = tempColumn[0] +1;
                });
    }

    public void createQuarterlyColumns(XSheet xSheet, String title, int row, int column, Set<Long> years) {
        final int[] tempColumn = { column };

        xSheet.addMergedRegion(new CellRangeAddress(row, row, column, column + years.size()*4 - 1));
        years.stream()
                .sorted(Comparator.reverseOrder())
                .forEach(year->{
                    xSheet.addMergedRegion(new CellRangeAddress(1, 1, tempColumn[0], tempColumn[0] + 3));
                    for(int quarter = 12 ; quarter > 0 ; quarter-=3) {
                        Cell cell0 = xSheet.getRow(row).createCell(tempColumn[0], CellType.STRING);
                        cell0.setCellValue(title);
                        cell0.setCellStyle(xSheet.getXWorkbook().getCellStyle(CELL_STYLE_HEADER));
                        Cell cell1 = xSheet.getRow(row+1).createCell(tempColumn[0], CellType.STRING);
                        cell1.setCellValue(String.valueOf(year));
                        cell1.setCellStyle(xSheet.getXWorkbook().getCellStyle(CELL_STYLE_HEADER));
                        Cell cell2 = xSheet.getRow(row+2).createCell(tempColumn[0], CellType.STRING);
                        cell2.setCellValue(String.valueOf(quarter));
                        cell2.setCellStyle(xSheet.getXWorkbook().getCellStyle(CELL_STYLE_HEADER));
                        tempColumn[0] = tempColumn[0] +1;
                    }
                });
    }

    public void saveToExcel(String filename) throws IOException {
        try (XWorkbook xWorkbook = new XWorkbook()) {

            CellStyle headerCellStyle = xWorkbook.createCellStyle(CELL_STYLE_HEADER);

            headerCellStyle.setAlignment(HorizontalAlignment.CENTER);
            headerCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            headerCellStyle.setFillBackgroundColor(new XSSFColor(new byte[]{(byte) 0, (byte) 0, (byte) 128}, null));
            headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerCellStyle.setBorderBottom(BorderStyle.THIN);
            headerCellStyle.setBottomBorderColor(IndexedColors.WHITE.getIndex());
            headerCellStyle.setBorderTop(BorderStyle.THIN);
            headerCellStyle.setTopBorderColor(IndexedColors.WHITE.getIndex());
            headerCellStyle.setBorderLeft(BorderStyle.THIN);
            headerCellStyle.setLeftBorderColor(IndexedColors.WHITE.getIndex());
            headerCellStyle.setBorderRight(BorderStyle.THIN);
            headerCellStyle.setRightBorderColor(IndexedColors.WHITE.getIndex());

            Font headerFont = xWorkbook.createFont(CELL_STYLE_HEADER);
            headerFont.setColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
            headerFont.setBold(true);
            headerCellStyle.setFont(headerFont);

            CellStyle roeCellStyle = xWorkbook.createCellStyle(CELL_STYLE_ROE);

            roeCellStyle.setBorderBottom(BorderStyle.THIN);
            roeCellStyle.setBottomBorderColor(IndexedColors.WHITE.getIndex());
            roeCellStyle.setBorderTop(BorderStyle.THIN);
            roeCellStyle.setTopBorderColor(IndexedColors.WHITE.getIndex());
            roeCellStyle.setBorderLeft(BorderStyle.THIN);
            roeCellStyle.setLeftBorderColor(IndexedColors.WHITE.getIndex());
            roeCellStyle.setBorderRight(BorderStyle.THIN);
            roeCellStyle.setRightBorderColor(IndexedColors.WHITE.getIndex());

            roeCellStyle.setDataFormat(xWorkbook.createDataFormat(CELL_STYLE_ROE, "0.00%"));

            CellStyle negativeRoeCellStyle = xWorkbook.createCellStyle(CELL_STYLE_NEGATIVE_ROE);

            negativeRoeCellStyle.setBorderBottom(BorderStyle.THIN);
            negativeRoeCellStyle.setBottomBorderColor(IndexedColors.WHITE.getIndex());
            negativeRoeCellStyle.setBorderTop(BorderStyle.THIN);
            negativeRoeCellStyle.setTopBorderColor(IndexedColors.WHITE.getIndex());
            negativeRoeCellStyle.setBorderLeft(BorderStyle.THIN);
            negativeRoeCellStyle.setLeftBorderColor(IndexedColors.WHITE.getIndex());
            negativeRoeCellStyle.setBorderRight(BorderStyle.THIN);
            negativeRoeCellStyle.setRightBorderColor(IndexedColors.WHITE.getIndex());

            Font roeFont = xWorkbook.createFont(CELL_STYLE_NEGATIVE_ROE);
            roeFont.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
            negativeRoeCellStyle.setFont(roeFont);
            negativeRoeCellStyle.setDataFormat(xWorkbook.createDataFormat(CELL_STYLE_NEGATIVE_ROE, "0.00%"));

            CellStyle pbrCellStyle = xWorkbook.createCellStyle(CELL_STYLE_PBR);
            pbrCellStyle.setDataFormat(xWorkbook.createDataFormat(CELL_STYLE_PBR, "0.00"));

            CellStyle currencyCellStyle = xWorkbook.createCellStyle(CELL_STYLE_CURRENCY);
            currencyCellStyle.setDataFormat(xWorkbook.createDataFormat(CELL_STYLE_CURRENCY, "#,##0"));

            CellStyle negativeCurrencyCellStyle = xWorkbook.createCellStyle(CELL_STYLE_NEGATIVE_CURRENCY);
            negativeCurrencyCellStyle.setDataFormat(xWorkbook.createDataFormat(CELL_STYLE_NEGATIVE_CURRENCY, "#,##0"));
            Font negativeCurrencyFont = xWorkbook.createFont(CELL_STYLE_NEGATIVE_CURRENCY);
            negativeCurrencyFont.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
            negativeCurrencyCellStyle.setFont(negativeCurrencyFont);

            XSheet xSheet = xWorkbook.createXSheet();
            Row[] header = {
                    xSheet.createRow(0),
                    xSheet.createRow(1),
                    xSheet.createRow(2)
            };

            Arrays.stream(FullCatalog.headerFields).forEach(field -> {
                xSheet.addMergedRegion(new CellRangeAddress(field.firstRow, field.lastRow, field.firstCol, field.lastCol));
                Cell cell = xSheet.getRow(field.firstRow).createCell(field.firstCol, field.type);
                cell.setCellValue(field.name);
                cell.setCellStyle(headerCellStyle);
            });

            Optional<Set<Long>> optionalYears = this.companies.values().stream().map(Company::getYears).reduce((x, y) -> {
                x.addAll(y);
                return x;
            });
            Optional<Set<Long>> optionalQuarterlies = this.companies.values().stream().map(Company::getQuarterlies).reduce((x, y) -> {
                x.addAll(y);
                return x;
            });

            if (optionalYears.isEmpty() || optionalQuarterlies.isEmpty()) {
                this.logger.log(Level.DEBUG, "저장할 데이터가 존재하지 않습니다.");
                return;
            }

            Set<Long> years = optionalYears.get();
            Set<Long> quarters = optionalQuarterlies.get();

            this.createQuarterlyColumns(xSheet, TITLE_RATING, 0, xSheet.getRow(0).getLastCellNum(), years);
            this.createYearlyColumns(xSheet, TITLE_ROE, 0, xSheet.getRow(0).getLastCellNum(), years);
            this.createColumn(xSheet, TITLE_PBR, 0, 3, xSheet.getRow(0).getLastCellNum(), 1);
            this.createYearlyColumns(xSheet, TITLE_YEARLY_SALES, 0, xSheet.getRow(0).getLastCellNum(), years);
            this.createQuarterlyColumns(xSheet, TITLE_QUARTERLY_ROE, 0, xSheet.getRow(0).getLastCellNum(), years);
            this.createQuarterlyColumns(xSheet, TITLE_QUARTERLY_SALES, 0, xSheet.getRow(0).getLastCellNum(), years);
            this.createColumn(xSheet, TITLE_MARKET_CAPITALIZATION, 0, 3, xSheet.getRow(0).getLastCellNum(), 1);

            quarters.stream()
                    .sorted(Comparator.reverseOrder())
                    .forEach(quarter -> {
                        xSheet.addMergedRegion(new CellRangeAddress(0, 1, header[0].getLastCellNum(), header[0].getLastCellNum() + 2));
                        Cell cell = header[0].createCell(header[0].getLastCellNum(), CellType.STRING);
                        cell.setCellValue(quarter);
                        header[0].createCell(header[0].getLastCellNum(), CellType.STRING).setCellValue(quarter);
                        header[0].createCell(header[0].getLastCellNum(), CellType.STRING).setCellValue(quarter);
                        header[1].createCell(header[1].getLastCellNum(), CellType.STRING).setCellValue(quarter);
                        header[1].createCell(header[1].getLastCellNum(), CellType.STRING).setCellValue(quarter);
                        header[1].createCell(header[1].getLastCellNum(), CellType.STRING).setCellValue(quarter);
                        header[2].createCell(header[2].getLastCellNum(), CellType.STRING).setCellValue(TITLE_SALES);
                        header[2].createCell(header[2].getLastCellNum(), CellType.STRING).setCellValue(TITLE_INCOME);
                        header[2].createCell(header[2].getLastCellNum(), CellType.STRING).setCellValue(TITLE_CAPITAL);
                    });


            xSheet.getRow(0).forEach(cell -> cell.setCellStyle(xSheet.xWorkbook.getCellStyle(CELL_STYLE_HEADER)));
            xSheet.getRow(1).forEach(cell -> cell.setCellStyle(xSheet.xWorkbook.getCellStyle(CELL_STYLE_HEADER)));
            xSheet.getRow(2).forEach(cell -> cell.setCellStyle(xSheet.xWorkbook.getCellStyle(CELL_STYLE_HEADER)));

            this.companies.values()
                    .stream()
                    .sorted(Comparator.comparingInt(x -> Integer.parseInt(x.getCode().substring(1))))
                    .forEach(company -> {
                        XRow xRow = xSheet.createXRow(xSheet.getLastRowNum() + 1);

                        xRow.createCell(xRow.getLastCellNum() + 1, CellType.STRING).setCellValue(company.getCode());
                        xRow.createCell(xRow.getLastCellNum(), CellType.STRING).setCellValue(company.getName());
                        xRow.createCell(xRow.getLastCellNum(), CellType.STRING).setCellValue(company.getCategories());
                        xRow.createCell(xRow.getLastCellNum(), CellType.STRING).setCellValue(company.getIndustryCode());
                        xRow.createCell(xRow.getLastCellNum(), CellType.STRING).setCellValue(company.getMarket());
                        xRow.createCell(xRow.getLastCellNum(), CellType.STRING).setCellValue(company.getComment());

                        long maxYear = years.stream().max(Long::compare).get();
                        long minYear = years.stream().min(Long::compare).get();

                        Set<Long> fullQuarterlies = new TreeSet<>();

                        years.stream().sorted().forEach(year->{
                            fullQuarterlies.add(year*100+3);
                            fullQuarterlies.add(year*100+6);
                            fullQuarterlies.add(year*100+9);
                            fullQuarterlies.add(year*100+12);
                        });

                        this.addQuarterlyRating(xRow, company, fullQuarterlies);
                        this.addYearlyRoe(xRow, company, minYear, maxYear);
                        this.addPbr(xRow, company);
                        this.addYearlySales(xRow, company, minYear, maxYear);
                        this.addQuarterlyRoe(xRow, company, fullQuarterlies);
                        this.addQuarterlySales(xRow, company, fullQuarterlies);
                        this.addMarketCapitalization(xRow, company);
                        this.addQuarterlyFinance(xRow, company, quarters);
                    });

            for(int i = START_OF_DATA_INDEX; i < xSheet.getRow(3).getLastCellNum(); i++) {
                xSheet.autoSizeColumn(i, true);
            }


            try (FileOutputStream output = new FileOutputStream(filename)) {
                xWorkbook.write(output);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    public void addQuarterlyRating(XRow xRow, Company company, Set<Long> fullQuarterlies) {
        Map<Long, Object> ratings = company.getQuarterlyItems(ITEM_RATING);
        fullQuarterlies.stream().sorted(Comparator.reverseOrder()).forEach(fullQuarterly->{
            Cell cell = xRow.createCell(xRow.getLastCellNum(), CellType.STRING);
            ratings.computeIfPresent(fullQuarterly, (fq, r)->{cell.setCellValue((String) r); return r; });
        });
    }


    public void addYearlySales(XRow xRow, Company company, long minYear, long maxYear) {
        for (long year = maxYear; year >= minYear; year--) {
            Cell cell = xRow.createCell(xRow.getLastCellNum(), CellType.NUMERIC);
            cell.setCellStyle(xRow.getXSheet().getXWorkbook().getCellStyle(CELL_STYLE_CURRENCY));
            try {
                cell.setCellValue((long)company.getAccount(ITEM_SALES, year));
            } catch (NullPointerException ignore) {
                Map<Long, Object> values = company.getValues(ITEM_QUARTERLY_SALES, year * 100 + 12, (year - 1) * 100 + 6);
                if (values.size() >= 4) {
                    Set<Long> quarterSales = values.keySet()
                            .stream()
                            .sorted(Comparator.reverseOrder())
                            .limit(4)
                            .collect(Collectors.toSet());

                    AtomicReference<Long> sum = new AtomicReference<>((long) 0);
                    quarterSales.forEach(quarter -> sum.updateAndGet(v -> (v + (long) values.get(quarter))));
                    cell.setCellValue(toHundredMillionUnit(sum.get()));
                }
            }
        }
    }


    public void addYearlyRoe(XRow xRow, Company company, long minYear, long maxYear) {
        for (long year = maxYear; year >= minYear; year--) {
            Cell cell = xRow.createCell(xRow.getLastCellNum(), CellType.NUMERIC);
            Double roe = null;
            try {
                if (year < 2020) {
                    roe = (Double) company.getAccount(ITEM_ROE, year);
                } else {
                    throw new NullPointerException();
                }
            } catch(NullPointerException ignore) {
                Map<Long, Object> values = company.getValues(ITEM_QUARTERLY_ROE, year * 100 + 12, (year - 1) * 100 + 6);
                if (values.size() >= 4) {
                    AtomicReference<Double> sum = new AtomicReference<>((double) 0);
                    values.keySet()
                            .stream()
                            .sorted(Comparator.reverseOrder())
                            .limit(4)
                            .forEach(quarter-> sum.updateAndGet(v -> (v + (double) values.get(quarter))));

                    roe = sum.get();
                }
            }

            if (roe != null) {
                cell.setCellValue(roe);
                if (roe >= 0) {
                    cell.setCellStyle(xRow.getXSheet().getXWorkbook().getCellStyle(CELL_STYLE_ROE));
                } else {
                    cell.setCellStyle(xRow.getXSheet().getXWorkbook().getCellStyle(CELL_STYLE_NEGATIVE_ROE));
                }
            }
        }
    }

    public void addPbr(XRow xRow, Company company) {
        Cell cell = xRow.createCell(xRow.getLastCellNum(), CellType.NUMERIC);
        cell.setCellValue(company.getPbr());
        cell.setCellStyle(xRow.getXSheet().getXWorkbook().getCellStyle(CELL_STYLE_PBR));
    }

    public void addMarketCapitalization(XRow xRow, Company company) {
        Cell cell = xRow.createCell(xRow.getLastCellNum(), CellType.NUMERIC);
        cell.setCellValue(toHundredMillionUnit(company.getMarketCapitalization()));
        cell.setCellStyle(xRow.getXSheet().getXWorkbook().getCellStyle(CELL_STYLE_CURRENCY));
    }

    public void addQuarterlyRoe(XRow xRow, Company company, Set<Long> fullQuarterlies) {
        fullQuarterlies.stream().sorted(Comparator.reverseOrder()).forEach(quarter-> {
            Cell cell = xRow.createCell(xRow.getLastCellNum(), CellType.NUMERIC);
            if (company.hasAccount(ITEM_QUARTERLY_ROE, quarter)) {
                double quarterlyRoe = (double)company.getAccount(ITEM_QUARTERLY_ROE, quarter);
                cell.setCellValue(quarterlyRoe);
                if (quarterlyRoe >= 0) {
                    cell.setCellStyle(xRow.getXSheet().getXWorkbook().getCellStyle(CELL_STYLE_ROE));
                } else {
                    cell.setCellStyle(xRow.getXSheet().getXWorkbook().getCellStyle(CELL_STYLE_NEGATIVE_ROE));
                }
            }
        });
    }

    public void addQuarterlySales(XRow xRow, Company company, Set<Long> fullQuarterlies) {
        Optional<Long> optionalMaxFullQuarterly = fullQuarterlies.stream().max(Long::compare);
        Optional<Long> optionalMinFullQuarterly = fullQuarterlies.stream().min(Long::compare);

        if (optionalMaxFullQuarterly.isPresent()&& optionalMinFullQuarterly.isPresent()) {
            fullQuarterlies.stream().sorted(Comparator.reverseOrder()).forEach(quarter -> {
                Cell cell = xRow.createCell(xRow.getLastCellNum(), CellType.NUMERIC);
                if (company.hasAccount(ITEM_QUARTERLY_SALES, quarter)) {
                    cell.setCellValue(toHundredMillionUnit((long) company.getAccount(ITEM_QUARTERLY_SALES, quarter)));
                    cell.setCellStyle(xRow.getXSheet().getXWorkbook().getCellStyle(CELL_STYLE_CURRENCY));
                }
                ;
            });
        }
    }

    public void addQuarterlyFinance(XRow xRow, Company company, Set<Long> quarters) {
        quarters.stream()
                .sorted(Comparator.reverseOrder())
                .forEach(quarter -> {
                    if (company.getQuarterlyFinancial().containsKey(quarter)) {
                        Finance finance = company.getQuarterlyFinancial().get(quarter);
                        xRow.createCell(xRow.getLastCellNum(), CellType.NUMERIC, xRow.getXSheet().getXWorkbook().getCellStyle(CELL_STYLE_CURRENCY))
                                .setCellValue((long) finance.getAccount(ITEM_SALES, 0L));

                        if ((long) finance.getAccount(ITEM_INCOME, 0L) >= 0 ) {
                            xRow.createCell(xRow.getLastCellNum(), CellType.NUMERIC, xRow.getXSheet().getXWorkbook().getCellStyle(CELL_STYLE_CURRENCY))
                                    .setCellValue((long) finance.getAccount(ITEM_INCOME, 0L));
                        } else {
                            xRow.createCell(xRow.getLastCellNum(), CellType.NUMERIC, xRow.getXSheet().getXWorkbook().getCellStyle(CELL_STYLE_NEGATIVE_CURRENCY))
                                    .setCellValue((long) finance.getAccount(ITEM_INCOME, 0L));
                        }

                        xRow.createCell(xRow.getLastCellNum(), CellType.NUMERIC, xRow.getXSheet().getXWorkbook().getCellStyle(CELL_STYLE_CURRENCY))
                                .setCellValue((long) finance.getAccount(ITEM_CAPITAL, 0L));
                    } else {
                        xRow.createCell(xRow.getLastCellNum(), CellType.NUMERIC).setCellValue(0);
                        xRow.createCell(xRow.getLastCellNum(), CellType.NUMERIC).setCellValue(0);
                        xRow.createCell(xRow.getLastCellNum(), CellType.NUMERIC).setCellValue(0);
                    }
                });
    }

    public long toHundredMillionUnit(long value) {
        return  value / 100000;
    }

    public long toDefaultUnit(long value) {
        return  value * 100000;
    }

    public JSONObject toJson() {
        JSONObject root = new JSONObject();

        this.companies.values()
                .stream()
                .sorted(Comparator.comparingInt(x -> Integer.parseInt(x.getCode().substring(1))))
                .forEach((company)->root.put(company.getCode(), company.toJson())) ;

        return  root;
    }
}
