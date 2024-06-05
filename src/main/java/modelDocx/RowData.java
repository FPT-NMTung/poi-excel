package modelDocx;

import io.vertx.core.json.JsonObject;

import java.util.ArrayList;

public class RowData {
    private JsonObject data;
    private ArrayList<RowData> childRow;

    public RowData(JsonObject data) {
        this.data = data;
        this.childRow = new ArrayList<>();
    }

    public RowData(JsonObject data, ArrayList<RowData> childRow) {
        this.data = data;
        this.childRow = childRow;
    }

    public JsonObject getData() {
        return data;
    }

    public void setData(JsonObject data) {
        this.data = data;
    }

    public ArrayList<RowData> getChildRow() {
        return childRow;
    }

    public void setChildRow(ArrayList<RowData> childRow) {
        this.childRow = childRow;
    }
}
