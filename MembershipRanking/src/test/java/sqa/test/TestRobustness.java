package sqa.test;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import sqa.main.Ranking;

import java.io.*;

class TestRobustness {
    private final Ranking ranking = new Ranking();

    @Test
    void testCalculateMembershipRankFromExcel() throws Exception {
        String filePath = "src/main/resources/TestCasesRobustness.xlsx";
        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        StringBuilder errors = new StringBuilder();

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null ||
                row.getCell(0) == null ||
                row.getCell(1) == null ||
                row.getCell(2) == null ||
                row.getCell(3) == null ||
                row.getCell(0).getCellType() != CellType.NUMERIC ||
                row.getCell(1).getCellType() != CellType.NUMERIC ||
                row.getCell(2).getCellType() != CellType.NUMERIC ||
                row.getCell(3).getCellType() != CellType.STRING) {
                continue;
            }

            int spending = (int) row.getCell(0).getNumericCellValue();
            int visits = (int) row.getCell(1).getNumericCellValue();
            int points = (int) row.getCell(2).getNumericCellValue();
            String expectedRank = row.getCell(3).getStringCellValue();

            String actualRank = ranking.CalculateMembershipRank(spending, visits, points);

            Cell actualCell = row.createCell(4);
            actualCell.setCellValue(actualRank);

            Cell resultCell = row.createCell(5);
            String result = actualRank.equals(expectedRank) ? "PASS" : "FAIL";
            resultCell.setCellValue(result);

            if (!result.equals("PASS")) {
                errors.append("Row ").append(i)
                      .append(" failed: expected ")
                      .append(expectedRank).append(", but got ")
                      .append(actualRank).append("\n");
            }
        }


        fis.close();

        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            workbook.write(fos);
        }
        workbook.close();

        if (errors.length() > 0) {
            throw new AssertionError("Some tests failed:\n" + errors);
        }
    }

}
