package modelDocx;

public class TableConfig {
    private String tableName;
    private RowConfig rowConfig;

    public TableConfig() {
    }

    public TableConfig(String tableName, RowConfig rowConfig) {
        this.tableName = tableName;
        this.rowConfig = rowConfig;
    }

    public String getTableName() {
        return tableName;
    }

    public void setTableName(String tableName) {
        this.tableName = tableName;
    }

    public RowConfig getRowConfig() {
        return rowConfig;
    }

    public void setRowConfig(RowConfig rowConfig) {
        this.rowConfig = rowConfig;
    }
}
