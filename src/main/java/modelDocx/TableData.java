package modelDocx;

import io.vertx.core.json.JsonObject;

import java.util.ArrayList;
import java.util.HashSet;

public class TableData {
    private String name;
    private ArrayList<RowData> rows;

    public TableData(String name) {
        this.name = name;
        this.rows = new ArrayList<>();
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public ArrayList<RowData> getRows() {
        return rows;
    }

    public void setRows(ArrayList<RowData> rows) {
        this.rows = rows;
    }
}
