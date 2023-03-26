import org.apache.poi.ss.usermodel.*;
import org.json.JSONObject;

import java.util.*;

public class ShoppingGuide extends ExcelFile {
    public static final int NAME_INDEX = 1;
    public static final int ITEM_CODE_INDEX = 2;
    public static final int MARKET_INDEX = 3;
    public static final int MARKET_CAPITALIZATION = 7;

    private final Map<String, Company> companies;

    public ShoppingGuide(Map<String, Company> companies) {
        super();
        this.companies = companies;
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
    protected void load(Row row) {
        String itemCode = row.getCell(ShoppingGuide.ITEM_CODE_INDEX).toString();

        if (!companies.containsKey(itemCode)) {
            return;
        }
        Company  company = companies.get(itemCode);

        try {
            long marketCapitalization = Long.parseLong(row.getCell(ShoppingGuide.MARKET_CAPITALIZATION).toString());

            company.setMarketCapitalization(marketCapitalization * 100000000);
        } catch(NumberFormatException ignore) {
        }
    }


    public JSONObject toJson() {
        JSONObject root = new JSONObject();

        this.companies.forEach((code, company)-> root.put(code, company.toJson()));

        return  root;
    }

    public String toString() {
        return  this.toJson().toString();
    }

    @Override
    protected boolean isHeader(Row row) {
        return  (row.getRowNum() < 3) ;
    }

    @Override
    protected boolean isData(Row row) {
        return  (row.getRowNum() >= 3) ;
    }

    @Override
    protected Map<Integer, String> getHeader(Sheet sheet) {
        // TODO Auto-generated method stub
        throw new UnsupportedOperationException("Unimplemented method 'getHeader'");
    }
}
