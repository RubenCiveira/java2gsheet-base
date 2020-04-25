package net.civeira.management.googlesheet;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.List;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.AppendValuesResponse;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.DeleteDimensionRequest;
import com.google.api.services.sheets.v4.model.DimensionRange;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.SheetProperties;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.ValueRange;

public class GoogleSheetApi {
  private final String sheetId;
  private final Sheets sheet;

  public GoogleSheetApi(String appName, GoogleAccessApi api, String sheetId)
      throws GeneralSecurityException, IOException {
    super();
    this.sheetId = sheetId;
    this.sheet = new Sheets.Builder(GoogleNetHttpTransport.newTrustedTransport(),
        JacksonFactory.getDefaultInstance(),
        api.authorize(Arrays.asList(SheetsScopes.SPREADSHEETS))).setApplicationName(appName)
            .build();
  }

  public List<List<Object>> read(String page, String fromColumn, int fromRow, String toColumn,
      int toRow) throws IOException {
    return this.sheet.spreadsheets().values()
        .get(sheetId, page + "!" + fromColumn + fromRow + ":" + toColumn + toRow).execute()
        .getValues();
  }

  public int append(String page, List<List<Object>> rows) throws IOException {
    AppendValuesResponse execute = this.sheet.spreadsheets().values()
        .append(sheetId, page, new ValueRange().setValues(rows)).setValueInputOption("USER_ENTERED")
        .setIncludeValuesInResponse(true).setInsertDataOption("INSERT_ROWS").execute();
    String range = execute.getTableRange();
    return Integer.parseInt(range.substring(range.lastIndexOf(':') + 1).replaceAll("[^\\d.]", ""))
        + 1;
  }

  public void update(String page, String fromColumn, int fromRow, List<List<Object>> rows)
      throws IOException {
    this.sheet.spreadsheets().values()
        .update(sheetId, page + "!" + fromColumn + fromRow, new ValueRange().setValues(rows))
        .setValueInputOption("USER_ENTERED").setIncludeValuesInResponse(true).execute();
  }

  public void deleteRow(String page, int row) throws IOException {
    deleteRows(page, row, row);
  }
  
  public void deleteRows(String page, int fromRow, int toRow) throws IOException {
    DeleteDimensionRequest deleteRequest =
        new DeleteDimensionRequest().setRange(new DimensionRange().setSheetId(sheetId(page))
            .setDimension("ROWS").setStartIndex(fromRow).setEndIndex(toRow+1));
    this.sheet.spreadsheets()
        .batchUpdate(sheetId,
            new BatchUpdateSpreadsheetRequest()
                .setRequests(Arrays.asList(new Request().setDeleteDimension(deleteRequest))))
        .execute();
  }

  public void deleteColumn(String page, int column) throws IOException {
    deleteColumns(page, column, column);
  }
  public void deleteColumns(String page, int fromColumn, int toColumn) throws IOException {
    DeleteDimensionRequest deleteRequest =
        new DeleteDimensionRequest().setRange(new DimensionRange().setSheetId(sheetId(page))
            .setDimension("COLUMNS").setStartIndex(fromColumn).setEndIndex(toColumn+1));
    this.sheet.spreadsheets()
        .batchUpdate(sheetId,
            new BatchUpdateSpreadsheetRequest()
                .setRequests(Arrays.asList(new Request().setDeleteDimension(deleteRequest))))
        .execute();
  }

  private int sheetId(String name) throws IOException {
    int result = -1;
    Spreadsheet response1 =
        this.sheet.spreadsheets().get(sheetId).setIncludeGridData(false).execute();
    for (Sheet sheet2 : response1.getSheets()) {
      SheetProperties properties = sheet2.getProperties();
      if (properties.getTitle().equals(name)) {
        result = properties.getSheetId();
      }
    }
    if (-1 == result) {
      throw new IOException("No sheet named " + name);
    }
    return result;
  }
}
