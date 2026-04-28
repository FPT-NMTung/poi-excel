package org.example;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.security.InvalidKeyException;
import java.security.NoSuchAlgorithmException;
import java.security.Signature;
import java.security.spec.InvalidKeySpecException;

public class TestSecurity {
    public static void main(String[] args) throws Exception {
        Path privateKeyFile = Paths.get("D:/SourceBuild/poi-excel/id_rsa");
        String privateKey = Files.readString(privateKeyFile);

        System.out.println(privateKey);

        Security security = new Security();

        Signature a = security.generateSignatureSign("PKCS8", "RSA", privateKey);
        String aa = security.sign(a, "asdawdawd");

        System.out.println(aa);
    }
}
