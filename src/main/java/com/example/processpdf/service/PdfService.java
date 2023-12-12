package com.example.processpdf.service;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class PdfService {
    public List<String> processPdfFile(MultipartFile file) throws IOException {
        String uploadedFilePath = "D:\\codehaus\\upload" + file.getOriginalFilename();
        file.transferTo(new File(uploadedFilePath));

        List<String> numbersList = extractNumbersFromPdf(uploadedFilePath);
        convertPdfToExcel(uploadedFilePath, "D:\\codehaus\\save_excel");
        return numbersList;
    }

    private List<String> extractNumbersFromPdf(String filePath) throws IOException {
        List<String> numbersList = new ArrayList<>();

        try (PDDocument document = PDDocument.load(new File(filePath))) {
            PDFTextStripper pdfStripper = new PDFTextStripper();
            String text = pdfStripper.getText(document);

            Pattern pattern = Pattern.compile("\\b\\d+\\b");
            Matcher matcher = pattern.matcher(text);
            while (matcher.find()) {
                numbersList.add(matcher.group());
            }
        }

        return numbersList;
    }

    private void convertPdfToExcel(String pdfFilePath, String excelFolderPath) throws IOException {
        String pdfFileName = new File(pdfFilePath).getName();
        String excelFilePath = excelFolderPath + File.separator + pdfFileName.replace(".pdf", ".xlsx");

        try (PDDocument document = PDDocument.load(new File(pdfFilePath));
             Workbook workbook = new XSSFWorkbook()) {

            PDFTextStripper pdfStripper = new PDFTextStripper();
            String text = pdfStripper.getText(document);

            Sheet sheet = workbook.createSheet("PDF Data");
            String[] lines = text.split("\\r?\\n");
            int rowNum = 0;
            for (String line : lines) {
                Row row = sheet.createRow(rowNum++);
                String[] cells = line.split("\\s+");
                int cellNum = 0;
                for (String cellData : cells) {
                    Cell cell = row.createCell(cellNum++);
                    cell.setCellValue(cellData);
                }
            }

            try (FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {
                workbook.write(fileOut);
            }
        }
    }

}
