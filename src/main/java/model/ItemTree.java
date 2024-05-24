package model;

import io.vertx.core.json.JsonObject;

public class ItemTree {
    private String key;
    private JsonObject value;
    private ChildTree child;

    public ItemTree() {

    }

    public ItemTree(JsonObject value, ChildTree child) {
        this.value = value;
        this.child = child;
    }

    public JsonObject getValue() {
        return value;
    }

    public void setValue(JsonObject value) {
        this.value = value;
    }

    public ChildTree getChild() {
        return child;
    }

    public void setChild(ChildTree child) {
        this.child = child;
    }

    public String getKey() {
        return key;
    }

    public void setKey(String key) {
        this.key = key;
    }
}
