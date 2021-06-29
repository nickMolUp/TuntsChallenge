import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.Collections;
import java.util.List;

public class SheetsQuickstart {
    private static final String APPLICATION_NAME = "Challenge";
    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "tokens";

    /**
     * Global instance of the scopes required by this quickstart.
     * If modifying these scopes, delete your previously saved tokens/ folder.
     */
    private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS);
    private static final String CREDENTIALS_FILE_PATH = "/credentials.json";

    /**
     * Creates an authorized Credential object.
     * @param HTTP_TRANSPORT The network HTTP Transport.
     * @return An authorized Credential object.
     * @throws IOException If the credentials.json file cannot be found.
     */
    private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
        // Load client secrets.
        InputStream in = SheetsQuickstart.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
        if (in == null) {
            throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
        }
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline")
                .build();
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }

    /**
     * Copied spreadsheet
     * https://docs.google.com/spreadsheets/d/1I9mV17Tb-OkR1UJvlY-775QGQFD7uHXlK3ZkwMfFN3U/edit?usp=sharing
     */
    public static void main(String... args) throws IOException, GeneralSecurityException {
        // Build a new authorized API client service.
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        final String spreadsheetId = "1I9mV17Tb-OkR1UJvlY-775QGQFD7uHXlK3ZkwMfFN3U"; //Sheet ID
        final String rangeRead = "engenharia_de_software!A1:H27"; //Range witch contains data
		final String rangeWrite = "engenharia_de_software!G4:H27";//Range to be inserted data
        Sheets service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                .setApplicationName(APPLICATION_NAME)
                .build();
        ValueRange response = service.spreadsheets().values()
                .get(spreadsheetId, rangeRead)
                .execute();
		
		List<List<Object>> values = response.getValues();//Read sheet range in list of lists
        if (values == null || values.isEmpty()) {//Verifies if there's any data
            System.out.println("No data found.");
        } else {
			String strTotalClasses = values.get(1).toString();//Row witch contains total classes per semester
			int totalClasses = Integer.parseInt(strTotalClasses.replaceAll("[^0-9.]", ""));//Total Classes in integer for calc			
			values.remove(0);//Removes extra rows
			values.remove(0);//Removes extra rows
			values.remove(0);//Removes extra rows
			values.forEach(row -> {//Reads and writes to "values"
				int missedClasses= Integer.parseInt(row.get(2).toString());//Gets student total missed classes
				if(missedClasses > (totalClasses/4)){//in case student haven't attended to enough classes
					row.removeAll(row);
					row.add("Reprovado por Falta");
					row.add("0");
				}else{				
				float m = (float) (((Float.parseFloat(row.get(3).toString())) + 
									(Float.parseFloat(row.get(4).toString()))+
									(Float.parseFloat(row.get(5).toString())))/3.0);//Calculates the student average
				m = (int) Math.ceil(m);//Rounds average 
				
				row.removeAll(row);
				if(m<50){
					row.add("Reprovado por Nota");
					row.add("0");
				}
				else if((50<=m)&&(m<70)){
					row.add("Exame final");
					row.add((m-10));
				}
				else if(m>70){
					row.add("Aprovado");
					row.add("0");
				}
				}				
			});
			
			ValueRange body = new ValueRange().setValues(values);
			UpdateValuesResponse result =
						service.spreadsheets().values().update(spreadsheetId, rangeWrite, body)
						.setValueInputOption("RAW")
						.execute();
			
        }
    }
}