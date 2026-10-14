package com.eit.automation.core;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Properties;

public class LocatorMapper {
    private static final Properties props = new Properties();

    static {
        // 🚀 Dynamic Centralized Excel Locator Loading from Root Directory
        String rootPath = System.getProperty("user.dir");
        File locatorWorkbook = new File(rootPath + File.separator + "Global_Locators.xlsx");

        System.out.println("📂 Looking for Centralized Locator Workbook: " + locatorWorkbook.getAbsolutePath());

        if (locatorWorkbook.exists() && locatorWorkbook.isFile()) {
            try (FileInputStream fis = new FileInputStream(locatorWorkbook);
                 Workbook workbook = new XSSFWorkbook(fis)) {

                int totalSheets = workbook.getNumberOfSheets();
                int totalMappings = 0;

                for (int i = 0; i < totalSheets; i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    String sheetName = sheet.getSheetName();
                    int sheetMappings = 0;

                    for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                        Row row = sheet.getRow(rowIndex);
                        if (row == null) continue;

                        Cell cell = row.getCell(0); // Column A (index 0)
                        if (cell == null) continue;

                        String cellValue = cell.getStringCellValue().trim();

                        // Skip empty rows, header labels, and comment lines starting with #
                        if (cellValue.isEmpty() || cellValue.equalsIgnoreCase("locators") || cellValue.startsWith("#")) {
                            continue;
                        }

                        int equalsIndex = cellValue.indexOf('=');
                        if (equalsIndex > 0) {
                            String key = cellValue.substring(0, equalsIndex).trim();
                            String value = cellValue.substring(equalsIndex + 1).trim();

                            props.setProperty(key, value);
                            sheetMappings++;
                            totalMappings++;
                        }
                    }

                    System.out.println("   ↳ Tab [" + sheetName + "]: Loaded " + sheetMappings + " locator mapping(s)");
                }

                System.out.println("✅ Centralized Object Repository Initialized. Total merged mappings: " + totalMappings);

            } catch (Exception e) {
                System.out.println("❌ Error reading centralized locator workbook " + locatorWorkbook.getName() + ": " + e.getMessage());
            }
        } else {
            System.out.println("⚠️ Warning: 'Global_Locators.xlsx' not found at project root directory.");
        }
    }

    /**
     * Resolves an Excel/CSV locator element.
     *
     * @param label The logical key (e.g., 'mobile.driver.agree.btn') or a raw XPath/ID.
     * @return The real target locator string.
     */
    public static String getXPath(String label) {
        if (label == null || label.trim().isEmpty()) {
            return "";
        }

        // Strip out any accidental wrapping quotes added by Excel/CSV sheet formatting
        String cleanLabel = label.replace("\"", "").trim();

        // Fallback: Return the property value if mapped, otherwise return the original sheet text
        return props.getProperty(cleanLabel, label);
    }
}