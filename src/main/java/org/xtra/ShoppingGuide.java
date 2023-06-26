package org.xtra;

import org.apache.logging.log4j.Level;
import org.apache.poi.ss.usermodel.*;
import org.json.JSONObject;

import java.util.*;

public class ShoppingGuide extends ExcelFile {
    public static final int RATING_INDEX = 0;
    public static final int NAME_INDEX = 1;
    public static final int ITEM_CODE_INDEX = 2;
    public static final int MARKET_INDEX = 3;
    public static final int MARKET_CAPITALIZATION = 7;
    public static final int PBR_INDEX = 11;

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
        String itemCode = row.getCell(ShoppingGuide.ITEM_CODE_INDEX).toString().trim();
        if (itemCode.matches("[0-9]+")) {
            itemCode = "A" + itemCode;
        } else if (!itemCode.matches("A[0-9]+")) {
            return;
        }

        if (!companies.containsKey(itemCode)) {
            return;
        }
        Company  company = companies.get(itemCode);

        try {
            Long capitalization = toLong(row.getCell(ShoppingGuide.MARKET_CAPITALIZATION));
            if (capitalization != null) {
                company.setMarketCapitalization(capitalization * 100000);
            }

            Double pbr = toDouble(row.getCell(ShoppingGuide.PBR_INDEX));
            if (pbr != null) {
                company.setPbr(pbr);
            }
        } catch(IllegalStateException | NumberFormatException ignore) {
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
        Map<Integer, String>  titles = new TreeMap<>();

        Row fieldRow = null;

        for(Row row : sheet) {
            try {
                if (row.getCell(ShoppingGuide.MARKET_INDEX).toString().equals("시장")) {
                    fieldRow = row;
                    break;
                }
            } catch (Exception ignore) {
            }
        }

        if (fieldRow != null) {

            fieldRow.forEach(cell -> {
                titles.put(cell.getColumnIndex(), cell.toString());
                this.logger.log(Level.DEBUG, "{} : {}", cell.getCellType(), cell.toString());
            });

            System.out.println(titles);
        }
        return  titles;

    }
}
