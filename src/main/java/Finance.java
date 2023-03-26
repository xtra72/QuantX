import org.json.JSONObject;
import java.util.HashMap;
import java.util.Map;

class Finance {

    private final Map<String, Object> values;

    public Finance() {
        this.values = new HashMap<>();

    }

    public Object getAccount(String title) throws NullPointerException {
        if (this.values.containsKey(title)) {
            return this.values.get(title);
        }

        throw new NullPointerException();
    }

    public Object getAccount(String title, Object defaultValue) throws NullPointerException {
        if (this.values.containsKey(title)) {
            return this.values.get(title);
        }

        return  defaultValue;
    }

    public void setAccount(String title, Object value) {
        this.values.put(title, value);
    }

    public boolean hasAccount(String title) {
        return  this.values.containsKey(title);
    }

    public JSONObject toJson() {
        JSONObject root = new JSONObject();
        this.values.forEach(root::put);
        return  root;
    }

    public String toString() {
        return  this.toJson().toString();
    }
}