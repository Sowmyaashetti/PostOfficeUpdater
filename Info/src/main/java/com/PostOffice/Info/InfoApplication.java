package com.PostOffice.Info;

import org.json.JSONException;
import org.json.JSONObject;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;

import org.json.JSONArray;


import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
@SpringBootApplication
public class InfoApplication {


	public static void main(String[] args) {
		SpringApplication.run(InfoApplication.class, args);
		// Read Excel file
		try (Workbook workbook = new XSSFWorkbook(new FileInputStream("C:\\Users\\soumy\\OneDrive\\PostOffice.xlsx"))) {
			Sheet sheet = workbook.getSheetAt(0); // Get the first sheet
			boolean isFirstRow = true; // Flag to identify the first row

			for (Row row : sheet) {
				if (isFirstRow) {
					isFirstRow = false;
					continue; // Skip the first row
				}

				Cell cityNameCell = row.getCell(0);
				int cityCode = (int) cityNameCell.getNumericCellValue();
				String cityName = String.valueOf(cityCode);

				// Prepare the API URL
				String apiUrl = "https://api.postalpincode.in/postoffice/" + cityName;

				// Make API request and process the response
				HttpClient httpClient = HttpClient.newHttpClient();
				HttpRequest httpRequest = HttpRequest.newBuilder()
						.uri(URI.create(apiUrl))
						.build();

				System.out.println(httpRequest);
//
				try {
					HttpResponse<String> httpResponse = httpClient.send(httpRequest, HttpResponse.BodyHandlers.ofString());
					String jsonResponse = httpResponse.body();


					// Extract the number of post offices from the API response
					int numberOfPostOffices = extractNumberOfPostOffices(jsonResponse);

					// Update the Excel sheet with the number of post offices
					Cell numberOfPostOfficesCell = row.createCell(1); // Assuming the number of post offices column is at index 1 (column B)
					numberOfPostOfficesCell.setCellValue(numberOfPostOffices);
				} catch (IOException | InterruptedException e) {
					// Handle API request error
					System.out.println(e);
				}
			}

			// Save the updated Excel file
			try (FileOutputStream outputStream = new FileOutputStream("C:\\Users\\soumy\\OneDrive\\PostOffice.xlsx")) {
				workbook.write(outputStream);
				System.out.println("Excel file updated successfully.");
			} catch (IOException e) {
				// Handle file writing error
				System.out.println(e);
			}
		} catch (IOException e) {
			// Handle file reading error
			System.out.println(e);
		}
	}

	// Method to extract the number of post offices from the JSON response
	private static int extractNumberOfPostOffices(String jsonResponse) {
		try {
			JSONArray jsonArray = new JSONArray(jsonResponse);
			JSONObject jsonObject = jsonArray.getJSONObject(0);
			String message = jsonObject.getString("Message");
			String[] parts = message.split(":");
			if (parts.length >= 2) {
				String countString = parts[1].trim();
				return Integer.parseInt(countString);
			}
		} catch (JSONException e) {
			System.out.println(e);
		}
		return 0;
	}
}