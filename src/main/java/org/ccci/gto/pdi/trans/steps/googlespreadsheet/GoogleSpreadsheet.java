package org.ccci.gto.pdi.trans.steps.googlespreadsheet;

import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.http.GenericUrl;
import com.google.api.client.http.HttpRequest;
import com.google.api.client.http.HttpRequestFactory;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.Base64;
import com.google.gdata.client.spreadsheet.FeedURLFactory;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.security.GeneralSecurityException;
import java.security.KeyStore;
import java.security.PrivateKey;
import java.util.Arrays;
import java.util.List;

public class GoogleSpreadsheet {
    public static final char[] SECRET = "notasecret".toCharArray();

    private static final HttpTransport HTTP_TRANSPORT = new NetHttpTransport();
    private static final JsonFactory JSON_FACTORY = new JacksonFactory();
    private static final List<String> SCOPES = Arrays.asList("https://spreadsheets.google.com/feeds",
            "https://docs.google.com/feeds");
    private static final FeedURLFactory FEED_URL_FACTORY = FeedURLFactory.getDefault();

    public static String base64EncodePrivateKeyStore(KeyStore pks) throws GeneralSecurityException, IOException {
        if (pks != null && pks.containsAlias("privatekey")) {
            ByteArrayOutputStream privateKeyStream = new ByteArrayOutputStream();
            pks.store(privateKeyStream, GoogleSpreadsheet.SECRET);
            return Base64.encodeBase64String(privateKeyStream.toByteArray());
        }
        return "";
    }

    public static KeyStore base64DecodePrivateKeyStore(String pks) throws GeneralSecurityException, IOException {
        if (pks != null && !pks.equals("")) {
            ByteArrayInputStream privateKeyStream = new ByteArrayInputStream(Base64.decodeBase64(pks));
            if (privateKeyStream.available() > 0) {
                KeyStore privateKeyStore = KeyStore.getInstance("PKCS12");
                privateKeyStore.load(privateKeyStream, GoogleSpreadsheet.SECRET);
                if (privateKeyStore.containsAlias("privatekey")) {
                    return privateKeyStore;
                }
            }
        }
        return null;
    }

    private static PrivateKey getPrivateKey(KeyStore pks) {
        try {
            if (pks != null) {
                return (PrivateKey) pks.getKey("privatekey", GoogleSpreadsheet.SECRET);
            }
        } catch (Exception e) {
        }
        return null;
    }

    public static String getAccessToken(String email, KeyStore pks) throws Exception {
        PrivateKey pk = getPrivateKey(pks);
        if (pk != null && !email.equals("")) {
            try {
                GoogleCredential credential = new GoogleCredential.Builder().setTransport(GoogleSpreadsheet.HTTP_TRANSPORT)
                        .setJsonFactory(GoogleSpreadsheet.JSON_FACTORY).setServiceAccountScopes(GoogleSpreadsheet.SCOPES).setServiceAccountId(email)
                        .setServiceAccountPrivateKey(pk).build();

                HttpRequestFactory requestFactory = GoogleSpreadsheet.HTTP_TRANSPORT.createRequestFactory(credential);
                GenericUrl url = new GenericUrl(GoogleSpreadsheet.getSpreadsheetFeedURL().toString());
                HttpRequest request = requestFactory.buildGetRequest(url);
                request.execute();
                return credential.getAccessToken();
            } catch (Exception e) {
                throw new Exception("Error fetching Access Token", e);
            }
        }
        return null;
    }

    private static URL getSpreadsheetFeedURL() {
        return GoogleSpreadsheet.FEED_URL_FACTORY.getSpreadsheetsFeedUrl();
    }

}
