package org.example;

import java.nio.charset.StandardCharsets;
import java.security.InvalidKeyException;
import java.security.KeyFactory;
import java.security.NoSuchAlgorithmException;
import java.security.PrivateKey;
import java.security.PublicKey;
import java.security.Signature;
import java.security.spec.InvalidKeySpecException;
import java.security.spec.PKCS8EncodedKeySpec;
import java.security.spec.X509EncodedKeySpec;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Base64;
import java.util.List;

public class Security {
    public Security() {
    }

    private byte[] toBytes(String value) {
        return (value == null ? "" : value).getBytes(StandardCharsets.UTF_8);
    }

    public byte[] getKey(String key) {
        String newKey = key
                .replace("-----BEGIN OPENSSH PUBLIC KEY-----", "")
                .replace("-----BEGIN OPENSSH PRIVATE KEY-----", "")
                .replace("-----END OPENSSH PRIVATE KEY-----", "")
                .replace("-----END OPENSSH PUBLIC KEY-----", "")
                .replace("\t", "")
                .replace(" ", "")
                .replaceAll("\n", "");

        System.out.println(newKey);

        return Base64.getDecoder().decode(newKey);
    }

    public PublicKey generatePublic(String algorithm, String key) throws Exception {
        byte[] newKey = this.getKey(key);
        X509EncodedKeySpec keySpec = new X509EncodedKeySpec(newKey);
        KeyFactory keyFactory = KeyFactory.getInstance(algorithm);
        return keyFactory.generatePublic(keySpec);
    }

    public PrivateKey generatePrivate(String algorithm, String key) throws NoSuchAlgorithmException, InvalidKeySpecException {
        byte[] newKey = this.getKey(key);
        PKCS8EncodedKeySpec keySpec = new PKCS8EncodedKeySpec(newKey);
        KeyFactory keyFactory = KeyFactory.getInstance(algorithm);
        return keyFactory.generatePrivate(keySpec);
    }

    public Signature generateSignatureVerify(String signatureAlgorithm, String keyAlgorithm, String publicKeyString) throws Exception {
        PublicKey publicKey = this.generatePublic(keyAlgorithm, publicKeyString);
        Signature signature = Signature.getInstance(signatureAlgorithm);
        signature.initVerify(publicKey);
        return signature;
    }

    public Signature generateSignatureSign(String signatureAlgorithm, String keyAlgorithm, String privateKeyString) throws NoSuchAlgorithmException, InvalidKeyException, InvalidKeySpecException {
        PrivateKey privateKey = this.generatePrivate(keyAlgorithm, privateKeyString);
        Signature signature = Signature.getInstance(signatureAlgorithm);
        signature.initSign(privateKey);
        return signature;
    }

    public synchronized String sign(Signature signature, String dataSign) throws Exception {
        signature.update(this.toBytes(dataSign));
        return new String(Base64.getEncoder().encode(signature.sign()));
    }

    public synchronized boolean verify(Signature signatureInstance, String dataSign, String signature) throws Exception {
        signatureInstance.update(this.toBytes(dataSign));
        return signatureInstance.verify(Base64.getDecoder().decode(this.toBytes(signature)));
    }
}
