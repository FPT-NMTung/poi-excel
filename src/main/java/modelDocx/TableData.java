package modelDocx;

import io.vertx.core.json.JsonObject;

import java.util.ArrayList;
import java.util.HashSet;

public class TableData {
    private String name;
    private HashSet<String> key;
    private ArrayList<JsonObject> rows;

    public TableData(String name) {
        this.name = name;
        this.key = new HashSet<>();
        this.rows = new ArrayList<>();
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public HashSet<String> getKey() {
        return key;
    }

    public void setKey(HashSet<String> key) {
        this.key = key;
    }

    public ArrayList<JsonObject> getRows() {
        return rows;
    }

    public void setRows(ArrayList<JsonObject> rows) {
        this.rows = rows;
    }
}
