package org.example;

import io.vertx.core.json.JsonObject;

public class Test {
    public static void main(String[] args) throws Exception {
        JsonObject resultData = new JsonObject();

        resultData.put("a", new JsonObject());

        JsonObject a = resultData.getJsonObject("a");
        a.put("b", new JsonObject());

        JsonObject c = resultData.getJsonObject("c");

        String aaaaaa = resultData.encode();

        System.out.println("123");
    }
}
