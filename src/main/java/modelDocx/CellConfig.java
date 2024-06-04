package modelDocx;

public class CellConfig {
    private String name;
    private String data;
    private String format;

    public CellConfig(String name, String data, String format) {
        this.name = name;
        this.data = data == null ? name : data;
        this.format = format;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getData() {
        return data;
    }

    public void setData(String data) {
        this.data = data;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }
}
