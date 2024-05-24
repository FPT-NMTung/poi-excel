package model;

import java.util.*;

public class ChildTree {
    private int level;
    private ArrayList<ItemTree> data;
    private HashSet<String> hashSetKey;

    public ChildTree() {

    }

    public ChildTree(int level) {
        this.level = level;
        this.data = new ArrayList<>();
        this.hashSetKey = new HashSet<>();
    }

    public ChildTree(int level, ArrayList<ItemTree> data) {
        this.level = level;
        this.data = data;
    }

    public int getLevel() {
        return level;
    }

    public void setLevel(int level) {
        this.level = level;
    }

    public ArrayList<ItemTree> getData() {
        return data;
    }

    public void setData(ArrayList<ItemTree> data) {
        this.data = data;
    }

    public HashSet<String> getHashSetKey() {
        return hashSetKey;
    }

    public void setHashSetKey(HashSet<String> hashSetKey) {
        this.hashSetKey = hashSetKey;
    }

    public boolean isContainKey (String key) {
        return this.hashSetKey.contains(key);
    }
}