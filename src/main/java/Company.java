import org.json.JSONObject;

import javax.management.InvalidAttributeValueException;
import java.util.*;
import java.util.stream.Collectors;

class Company {
    private final String code;
    private       String name;
    private final String industryCode;
    private final String categories;
    private final String market;
    private       String comment;
    private       double pbr;
    private       long   marketCapitalization;
    private final Map<Long, Finance> yearlyFinancial;
    private final Map<Long, Finance> quarterlyFinancial;

    Company(String market, String code, String name, String industryCode, String categories) {
        this(market, code, name, industryCode, categories, "") ;
    }

    Company(String market, String code, String name, String industryCode, String categories, String comment) {
        this.market = market;
        this.code = code;
        this.name = name;
        this.industryCode = industryCode;
        this.categories = categories;
        this.comment = comment;
        this.pbr = 0.0;
        this.yearlyFinancial = new HashMap<>();
        this.quarterlyFinancial = new HashMap<>();
    }

    public String getCode() {
        return  this.code;
    }

    public String getName() {
        return  this.name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getIndustryCode() {
        return  this.industryCode;
    }

    public String getCategories() {
        return  this.categories;
    }

    public String getMarket() {
        return  this.market;
    }

    public String getComment() {
        return  this.comment;
    }

    public void setComment(String comment) {
        this.comment = comment;
    }

    public double getPbr() { return this.pbr; }
    public void setPbr(double pbr) { this.pbr = pbr; }

    public double getMarketCapitalization() { return this.marketCapitalization; }
    public void setMarketCapitalization(long marketCapitalization) { this.marketCapitalization = marketCapitalization; }

    public Map<Long, Finance> getQuarterlyFinancial() {
        return  this.quarterlyFinancial;
    }

    public Set<Long> getYears() {
        Set<Long> years = this.yearlyFinancial.keySet().stream().collect(Collectors.toCollection(LinkedHashSet::new));
        years.addAll(this.quarterlyFinancial.keySet().stream().map(q->q/100).sorted().collect(Collectors.toCollection(LinkedHashSet::new)));

        return  years;
    }

    public Set<Long> getQuarterlies() {
        return new LinkedHashSet<>(this.quarterlyFinancial.keySet());
    }

    public Map<Long, Object> getQuarterlyItems(String item) {
        Map<Long, Object> quarterItems = new TreeMap<>();
        Set<Long> quarters = this.quarterlyFinancial.keySet();

        for(Long quarter : quarters) {
            try {
                quarterItems.put(quarter, this.quarterlyFinancial.get(quarter).getAccount(item));
            } catch (NullPointerException ignore) {}
        }

        return  quarterItems;
    }

    public Long getFirstQuarterly() {
        Optional<Long> optionalResult = quarterlyFinancial.keySet().stream().min(Long::compare);
        if (optionalResult.isPresent()) {
            return optionalResult.get();
        }

        throw new NullPointerException();
    }

    public Long getLastQuarterly() {
        Optional<Long> optionalResult = quarterlyFinancial.keySet().stream().max(Long::compare);
        if (optionalResult.isPresent()) {
            return optionalResult.get();
        }

        throw new NullPointerException();
    }

    public Map<Long, Object> getValues(String item, Long firstQuarterly, Long lastQuarterly) {
        Map<Long, Object> values = new HashMap<>();

        this.quarterlyFinancial.forEach((quarter, finance)-> {
            if ((firstQuarterly <= quarter && quarter <= lastQuarterly) || (lastQuarterly <= quarter && quarter <= firstQuarterly))  {
                try {
                    values.put(quarter, finance.getAccount(item));
                } catch(NullPointerException ignore) {}
            }
        });

        return  values;
    }

    public Finance getFinance(Long year, Long quarter) throws NullPointerException {
        Long key = year*100 + quarter;
        if (this.quarterlyFinancial.containsKey(key)) {
            return  this.quarterlyFinancial.get(key);
        }

        throw new NullPointerException("Finance[" + year + ":Q" + quarter + " is not exist");
    }

    public boolean hasAccount(String title, Long session) {
        if (session > 9999) {
            if (this.quarterlyFinancial.containsKey(session)) {
                return  this.quarterlyFinancial.get(session).hasAccount(title);
            }
        }
        else {
            if (this.yearlyFinancial.containsKey(session)) {
                return  this.yearlyFinancial.get(session).hasAccount(title);
            }
        }

        return  false;
    }

    public Object getAccount(String title, Long session) {
        if (session > 9999) {
            if (this.quarterlyFinancial.containsKey(session)) {
                return ((Finance) this.quarterlyFinancial.get(session)).getAccount(title);
            }
        }
        else {
            if (this.quarterlyFinancial.containsKey(session)) {
                return ((Finance) this.quarterlyFinancial.get(session)).getAccount(title);
            }
        }

       throw new NullPointerException(title + " of " + session + " is not exist");
    }

    public void setAccount(String title, Long year, Object value) throws InvalidAttributeValueException {
        if (value != null) {
            if ((value instanceof String) && (((String) value).length() == 0)) {
                throw new InvalidAttributeValueException();
            }

            if (this.yearlyFinancial.containsKey(year)) {
                this.yearlyFinancial.get(year).setAccount(title, value);
            } else {
                Finance finance = new Finance();
                finance.setAccount(title, value);
                this.yearlyFinancial.put(year, finance);
            }
        }
    }

    public void setAccount(String title, Long year, Long quarter, Object value) throws InvalidAttributeValueException {
        if (value != null) {
            if ((value instanceof String) && (((String) value).length() == 0)) {
                throw new InvalidAttributeValueException();
            }

            Long key = year * 100 + quarter;
            if (this.quarterlyFinancial.containsKey(key)) {
                this.quarterlyFinancial.get(key).setAccount(title, value);
            } else {
                Finance finance = new Finance();
                finance.setAccount(title, value);
                this.quarterlyFinancial.put(key, finance);
            }
        }
    }


    public void setAccount(String title, String fullQuarterly, Object value) throws InvalidAttributeValueException {
        if (value != null) {
            if (fullQuarterly.matches("\\d\\d\\d\\d")) {
                this.setAccount(title,
                        Long.parseLong(fullQuarterly.substring(0, 4)),
                        value);
            } else if (fullQuarterly.matches("\\d\\d\\d\\d\\d\\d")) {
                this.setAccount(title,
                        Long.parseLong(fullQuarterly.substring(0, 4)),
                        Long.parseLong(fullQuarterly.substring(4, 6)),
                        value);
            }
        }
    }


    public JSONObject toJson() {
        JSONObject root = new JSONObject();
        JSONObject information = new JSONObject();

        information.put("시장", this.market);
        information.put("종목코드", this.code);
        information.put("이름", this.name);
        information.put("산업코드", this.industryCode);
        information.put("업종", this.categories);
        information.put("설명", this.comment);
        information.put("PBR", this.pbr);
        information.put("시총", this.marketCapitalization);
        root.put("기본정보", information);
        JSONObject yearFinance = new JSONObject();
        this.yearlyFinancial.forEach((key, value)-> yearFinance.put(String.valueOf(key), value.toJson()));
        JSONObject quarterFinance = new JSONObject();
        this.quarterlyFinancial.forEach((key, value)-> quarterFinance.put(String.valueOf(key), value.toJson()));

        root.put("회기재무", yearFinance);
        root.put("분기재무", quarterFinance);
        return  root;
    }

    public String toString() {
        return  this.toJson().toString();
    }
}

