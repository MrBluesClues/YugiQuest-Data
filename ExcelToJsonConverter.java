package com.example.excelconverter;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelToJsonConverter {
    public static void main(String[] args) {
        // Define file paths for pack data
        String packExcelFilePath = "C:\\Users\\Mateo\\Documents\\Aden\\commonpack.xlsx";
        String packJsonFilePath = "D:\\stuff\\New folder\\Forge-Yugi-1.20.X\\src\\main\\resources\\assets\\yugiquest\\data\\common_pack.json";

        // Generate pack JSON from pack Excel
        generatePackJson(packExcelFilePath, packJsonFilePath);
    }

    private static void generatePackJson(String excelFilePath, String jsonFilePath) {
        List<Map<String, Object>> packData = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }

                Cell idCell = row.getCell(0);
                Cell chanceCell = row.getCell(1);
                int id = 0;
                double chance = 0.0;

                if (idCell != null) {
                    if (idCell.getCellType() == CellType.NUMERIC) {
                        id = (int) idCell.getNumericCellValue();
                    } else if (idCell.getCellType() == CellType.STRING) {
                        id = Integer.parseInt(idCell.getStringCellValue());
                    }
                }

                if (chanceCell != null) {
                    if (chanceCell.getCellType() == CellType.NUMERIC) {
                        chance = chanceCell.getNumericCellValue();
                    } else if (chanceCell.getCellType() == CellType.STRING) {
                        chance = Double.parseDouble(chanceCell.getStringCellValue());
                    }
                }

                Map<String, Object> card = new HashMap<>();
                card.put("id", id);
                card.put("chance", chance);
                packData.add(card);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        Gson gson = new GsonBuilder().setPrettyPrinting().create();
        try {
            Files.createDirectories(Paths.get(jsonFilePath).getParent());
            try (FileWriter writer = new FileWriter(jsonFilePath)) {
                gson.toJson(packData, writer);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Pack JSON file created successfully.");
    }
}
