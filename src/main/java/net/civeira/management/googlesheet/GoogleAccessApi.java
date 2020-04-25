package net.civeira.management.googlesheet;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.List;
import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;

public class GoogleAccessApi {
  
  private final GoogleClientSecrets clientSecrets;

  public GoogleAccessApi(InputStream key) throws IOException {
    clientSecrets = GoogleClientSecrets.load(JacksonFactory.getDefaultInstance(), 
        new InputStreamReader(key) );
  }
  
  public Credential authorize(List<String> scopes) throws IOException, GeneralSecurityException {
    GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(GoogleNetHttpTransport.newTrustedTransport(), 
        JacksonFactory.getDefaultInstance(), clientSecrets, scopes)
        .setDataStoreFactory(new FileDataStoreFactory(new File("tokens")))
        .setAccessType("offline")
        .build();
    return new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
  }  
}
