import org.apache.logging.log4j.Level;
import org.apache.poi.ss.usermodel.*;
import org.json.JSONObject;

import javax.management.InvalidAttributeValueException;
import java.util.*;
import java.util.List;
import java.util.regex.Pattern;

public class FnGuide extends ExcelFile {
    public static final String TITLE_CURRENT_ASSETS = "유동자산";
    public static final String TITLE_NON_CURRENT_ASSETS = "비유동자산";
    public static final String TITLE_LIABILITIES = "부채총계";
    public static final String TITLE_CURRENT_LIABILITIES = "유동부채";
    public static final String TITLE_NON_CURRENT_LIABILITIES = "비유동부채";
    public static final String TITLE_ASSETS = "자산총계";
    public static final String TITLE_CAPITAL = "자본총계";
    public static final String TITLE_EQUITY = "지배기업주주지분";
    public static final String TITLE_NON_EQUITY = "비지배기업주주지분";
    public static final String TITLE_COST_OF_SALES = "매출원가";
    public static final String TITLE_GROSS_PROFIT_ON_SALES = "매출총이익";
    public static final String TITLE_BUSINESS_PROFIT = "영업이익";
    public static final String TITLE_CURRENT_PROFIT = "당기순이익";
    public static final String TITLE_SALES = "매출액";
    public static final String TITLE_QUARTERLY_SALES = "분기매출액";
    public static final String TITLE_FINANCIAL_CONST = "금융비용";

    private static final Map<String, Map<String, String>> accountMapper = Map.ofEntries(
            Map.entry("주재무제표", Map.ofEntries(
                    Map.entry(TITLE_CURRENT_LIABILITIES, "CURRENT_LIABILITIES"),
                    Map.entry(TITLE_ASSETS, "ASSETS"),
                    Map.entry(TITLE_CURRENT_PROFIT, "INCOME"),
                    Map.entry(TITLE_FINANCIAL_CONST, "FINANCIAL_COST"),
                    Map.entry(TITLE_EQUITY, "OWNERS_OF_PARENT_EQUITY"),
                    Map.entry(TITLE_NON_EQUITY, "OWNERS_OF_NON_PARENT_EQUITY"),
                    Map.entry(TITLE_COST_OF_SALES, "CONST_OF_SALES"),
                    Map.entry(TITLE_GROSS_PROFIT_ON_SALES, "GROSS_PROFIT_ON_SALES"),
                    Map.entry(TITLE_BUSINESS_PROFIT, "BUSINESS_PROFIT"),
                    Map.entry(TITLE_CURRENT_ASSETS, "CURRENT_ASSETS"),
                    Map.entry(TITLE_NON_CURRENT_ASSETS, "NON_CURRENT_ASSETS"),
                    Map.entry(TITLE_NON_CURRENT_LIABILITIES, "NON_CURRENT_LIABILITIES"),
                    Map.entry(TITLE_LIABILITIES, "LIABILITIES"),
                    Map.entry(TITLE_CAPITAL, "CAPITAL"),
                    Map.entry(TITLE_SALES, "SALES"),
                    Map.entry(TITLE_QUARTERLY_SALES, "QUARTERLY_SALES"))),
            Map.entry("제조", Map.ofEntries(
                    Map.entry(TITLE_CURRENT_LIABILITIES, "CURRENT_LIABILITIES"),
                    Map.entry(TITLE_ASSETS, "ASSETS"),
                    Map.entry(TITLE_CURRENT_PROFIT, "INCOME"),
                    Map.entry(TITLE_FINANCIAL_CONST, "FINANCIAL_COST"),
                    Map.entry(TITLE_EQUITY, "OWNERS_OF_PARENT_EQUITY"),
                    Map.entry(TITLE_NON_EQUITY, "OWNERS_OF_NON_PARENT_EQUITY"),
                    Map.entry(TITLE_COST_OF_SALES, "CONST_OF_SALES"),
                    Map.entry(TITLE_GROSS_PROFIT_ON_SALES, "GROSS_PROFIT_ON_SALES"),
                    Map.entry(TITLE_BUSINESS_PROFIT, "BUSINESS_PROFIT"),
                    Map.entry(TITLE_CURRENT_ASSETS, "CURRENT_ASSETS"),
                    Map.entry(TITLE_NON_CURRENT_ASSETS, "NON_CURRENT_ASSETS"),
                    Map.entry(TITLE_NON_CURRENT_LIABILITIES, "NON_CURRENT_LIABILITIES"),
                    Map.entry(TITLE_LIABILITIES, "LIABILITIES"),
                    Map.entry(TITLE_CAPITAL, "CAPITAL"),
                    Map.entry(TITLE_SALES, "SALES"),
                    Map.entry(TITLE_QUARTERLY_SALES, "QUARTERLY_SALES"))),
            Map.entry("제조(연결)", Map.ofEntries(
                    Map.entry(TITLE_CURRENT_LIABILITIES, "CURRENT_LIABILITIES"),
                    Map.entry(TITLE_ASSETS, "ASSETS"),
                    Map.entry(TITLE_CURRENT_PROFIT, "INCOME"),
                    Map.entry("(당기순이익귀속)\n지배기업주주지분", "INCOME"),
                    Map.entry(TITLE_FINANCIAL_CONST, "FINANCIAL_COST"),
                    Map.entry(TITLE_EQUITY, "OWNERS_OF_PARENT_EQUITY"),
                    Map.entry(TITLE_NON_EQUITY, "OWNERS_OF_NON_PARENT_EQUITY"),
                    Map.entry(TITLE_COST_OF_SALES, "CONST_OF_SALES"),
                    Map.entry(TITLE_GROSS_PROFIT_ON_SALES, "GROSS_PROFIT_ON_SALES"),
                    Map.entry(TITLE_BUSINESS_PROFIT, "BUSINESS_PROFIT"),
                    Map.entry(TITLE_CURRENT_ASSETS, "CURRENT_ASSETS"),
                    Map.entry(TITLE_NON_CURRENT_ASSETS, "NON_CURRENT_ASSETS"),
                    Map.entry(TITLE_NON_CURRENT_LIABILITIES, "NON_CURRENT_LIABILITIES"),
                    Map.entry(TITLE_LIABILITIES, "LIABILITIES"),
                    Map.entry(TITLE_CAPITAL, "CAPITAL"),
                    Map.entry(TITLE_SALES, "SALES"),
                    Map.entry(TITLE_QUARTERLY_SALES, "QUARTERLY_SALES"))),
            Map.entry("금용", Map.ofEntries(
                    Map.entry(TITLE_CURRENT_LIABILITIES, "CURRENT_LIABILITIES"),
                    Map.entry(TITLE_ASSETS, "SALES"),
                    Map.entry(TITLE_CURRENT_PROFIT, "INCOME"),
                    Map.entry(TITLE_FINANCIAL_CONST, "FINANCIAL_COST"),
                    Map.entry(TITLE_EQUITY, "OWNERS_OF_PARENT_EQUITY"),
                    Map.entry(TITLE_NON_EQUITY, "OWNERS_OF_NON_PARENT_EQUITY"),
                    Map.entry(TITLE_COST_OF_SALES, "CONST_OF_SALES"),
                    Map.entry(TITLE_GROSS_PROFIT_ON_SALES, "GROSS_PROFIT_ON_SALES"),
                    Map.entry(TITLE_BUSINESS_PROFIT, "BUSINESS_PROFIT"),
                    Map.entry(TITLE_CURRENT_ASSETS, "CURRENT_ASSETS"),
                    Map.entry(TITLE_NON_CURRENT_ASSETS, "NON_CURRENT_ASSETS"),
                    Map.entry(TITLE_NON_CURRENT_LIABILITIES, "NON_CURRENT_LIABILITIES"),
                    Map.entry(TITLE_LIABILITIES, "LIABILITIES"),
                    Map.entry(TITLE_CAPITAL, "CAPITAL"))),
            Map.entry("금용(연결)", Map.ofEntries(
                    Map.entry(TITLE_CURRENT_LIABILITIES, "CURRENT_LIABILITIES"),
                    Map.entry("(당기순이익귀속)\n지배기업주주지분", "INCOME"),
                    Map.entry(TITLE_FINANCIAL_CONST, "FINANCIAL_COST"),
                    Map.entry(TITLE_EQUITY, "OWNERS_OF_PARENT_EQUITY"),
                    Map.entry(TITLE_NON_EQUITY, "OWNERS_OF_NON_PARENT_EQUITY"),
                    Map.entry(TITLE_COST_OF_SALES, "CONST_OF_SALES"),
                    Map.entry(TITLE_GROSS_PROFIT_ON_SALES, "GROSS_PROFIT_ON_SALES"),
                    Map.entry(TITLE_BUSINESS_PROFIT, "BUSINESS_PROFIT"),
                    Map.entry(TITLE_CURRENT_ASSETS, "CURRENT_ASSETS"),
                    Map.entry(TITLE_NON_CURRENT_ASSETS, "NON_CURRENT_ASSETS"),
                    Map.entry(TITLE_NON_CURRENT_LIABILITIES, "NON_CURRENT_LIABILITIES"),
                    Map.entry(TITLE_LIABILITIES, "LIABILITIES"),
                    Map.entry(TITLE_CAPITAL, "CAPITAL"),
                    Map.entry(TITLE_ASSETS, "SALES")))
    );

    public static final int MARKET_INDEX = 0;
    public static final int ITEM_CODE_INDEX = 1;
    public static final int NAME_INDEX = 2;
    public static final int INDUSTRY_INDEX = 4;
    public static final int CATEGORY_INDEX = 5;
    public static final int START_OF_DATA_INDEX = 7;
    public static final String MARKET_KOSPI = "KS";
    public static final String MARKET_KOSDAQ = "KQ";

    private final Map<String, Company> companies;

    public FnGuide(Map<String, Company> companies) {
        super();
        this.companies = companies;
    }

    @Override
    protected boolean isData(Row row) {
        return  ((row.getFirstCellNum() == 0) && (row.getCell(0).getCellType() == CellType.STRING)
                && (row.getCell(FnGuide.MARKET_INDEX).getStringCellValue().equals(FnGuide.MARKET_KOSDAQ)
                || (row.getCell(FnGuide.MARKET_INDEX).getStringCellValue().equals(FnGuide.MARKET_KOSPI))));
    }

    protected String getCellToString(Row row, int index) {
        String value = "";
        if (row.getCell(index).getCellType() == CellType.STRING) {
            value = row.getCell(index).getStringCellValue();
        } else if (row.getCell(index).getCellType() == CellType.NUMERIC) {
            value = String.valueOf((long)row.getCell(index).getNumericCellValue());
        }

        return  value;
    }

    @Override
    protected boolean isHeader(Row row) {
        return  !((row.getFirstCellNum() == 0) && (row.getCell(0).getCellType() == CellType.STRING)
                && (row.getCell(FnGuide.MARKET_INDEX).getStringCellValue().equals(FnGuide.MARKET_KOSDAQ)
                || (row.getCell(FnGuide.MARKET_INDEX).getStringCellValue().equals(FnGuide.MARKET_KOSPI))));
    }

    @Override
    protected void load(Row row) {
        String itemCode = row.getCell(FnGuide.ITEM_CODE_INDEX).toString();

        Company company = this.companies.computeIfAbsent(itemCode, code-> {
            String market = this.getCellToString(row, FnGuide.MARKET_INDEX);
            String name = this.getCellToString(row, FnGuide.NAME_INDEX);
            String industryCode = this.getCellToString(row, FnGuide.INDUSTRY_INDEX);
            String categories = this.getCellToString(row, FnGuide.CATEGORY_INDEX);

            this.logger.log(Level.DEBUG, "{} : {}, {}, {}, {}", row.getRowNum(), itemCode, name, industryCode, categories);
            return  new Company( market, itemCode, name, industryCode, categories);
        });

        row.forEach(cell -> {
            if (FnGuide.START_OF_DATA_INDEX <= cell.getColumnIndex() && header.containsKey(cell.getColumnIndex()) && cell.getCellType() == CellType.NUMERIC) {
                String[] fields = header.get(cell.getColumnIndex()).split("-");
                if (Pattern.compile("\\d+/누적", Pattern.CANON_EQ).matcher(fields[1]).matches()) {
                    fields[1] = fields[1].substring(0,6);
                }
                try {
                    company.setAccount(fields[0], fields[1], (long) cell.getNumericCellValue());
                } catch (InvalidAttributeValueException ignore) {}
            }
        });
    }

    @Override
    protected Map<Integer, String> getHeader(Sheet sheet) {
        Map<Integer, String>  titles = new TreeMap<>();

        Row fieldRow = null;

        for(Row row : sheet) {
            if ((row.getFirstCellNum() == FnGuide.MARKET_INDEX)
                    && (row.getCell(FnGuide.MARKET_INDEX).toString().equals("시장"))
                    && (sheet.getRow(row.getRowNum()+1) != null)) {
                fieldRow = row;
                break;
            }
        }

        if (fieldRow == null) {
            return Collections.emptyMap();
        }

        Row quarterRow = sheet.getRow(fieldRow.getRowNum() + 1);

        fieldRow.forEach(cell -> {
            String quarter;
            if (quarterRow.getCell(cell.getColumnIndex()).getCellType() == CellType.NUMERIC) {
                quarter = String.valueOf((long) quarterRow.getCell(cell.getColumnIndex()).getNumericCellValue());
            } else {
                quarter = quarterRow.getCell(cell.getColumnIndex()).getStringCellValue();
            }


            if (FnGuide.accountMapper.containsKey(sheet.getSheetName())
                    && FnGuide.accountMapper.get(sheet.getSheetName()).containsKey(cell.toString())) {
                String title = FnGuide.accountMapper.get(sheet.getSheetName()).get(cell.toString());
                if (quarter.length() != 0) {
                    title += "-" + quarter;
                }

                titles.put(cell.getColumnIndex(), title);
                this.logger.log(Level.DEBUG, "{} : {}", cell.getCellType(), title);
            }
        });

        System.out.println(titles);
        return  titles;
    }

    public List<String> getCompanyNames() {
        List<String> companyNames = new ArrayList<>();

        this.companies.forEach((code, company)-> companyNames.add(company.getName()));

        return  companyNames;
    }


    public JSONObject toJson() {
        JSONObject root = new JSONObject();

        this.companies.forEach((code, company)-> root.put(code, company.toJson()));

        return  root;
    }

    public String toString() {
        return  this.toJson().toString();
    }
}
