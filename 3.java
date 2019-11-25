package io.alexeylevin.writer;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public class Writer {

    // 1495 count start 926264 for 927759
    // 357 count start 930014 for 930371
    public static void main(String[] args) throws IOException {
        List<Integer> talons1495 = generateTalonNumbers(1495, 926264, 927759);
        log.info(talons1495.toString());
        List<Integer> talons357 = generateTalonNumbers(357, 930014, 930371);
        log.info(talons357.toString());
        ArrayList<Integer> allTalons = new ArrayList<>(talons357);
        allTalons.addAll(talons1495);
        write(talons357, "talons357");
        write(talons1495, "talons1495");
        write(allTalons, "allTalons");
    }

    public static List<Integer> generateTalonNumbers(int size, int start, int end) {
        List<Integer> talons = new ArrayList<>(size + 9);
        for (int i = start; i <= end; i++) {
            talons.add(i);
        }
        return talons;
    }

    public static void write(List<Integer> talonNumbers, String name) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(name);

        for (int i = 0; i < talonNumbers.size(); i += 3) {
            Row row = sheet.createRow(i / 3);
            int iter = i;
            try {
                row.createCell(0).setCellValue(createCellFromTemplate(workbook, talonNumbers.get(iter)));
                iter = i + 1;
                row.createCell(1).setCellValue(createCellFromTemplate(workbook, talonNumbers.get(iter)));
                iter = i + 2;
                row.createCell(2).setCellValue(createCellFromTemplate(workbook, talonNumbers.get(iter)));
            } catch (IndexOutOfBoundsException e) {
                log.info(String.format("Кончился массив %d %d", iter, talonNumbers.size()));
            }
        }
        for (int i = 0; i < 3; i++) {
            sheet.autoSizeColumn(i);
        }

        FileOutputStream fileOut = new FileOutputStream(name + ".xlsx");
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
    }

    private static XSSFRichTextString createCellFromTemplate(XSSFWorkbook workbook, int number) {
        XSSFFont fontBold = workbook.createFont();
        fontBold.setBold(true);
        XSSFFont fontUnderline = workbook.createFont();
        fontUnderline.setUnderline((byte)1);

        XSSFRichTextString cellValue = new XSSFRichTextString();
        cellValue.append("\n");
        cellValue.append("        ООО «ЭКОСТРОЙПРОГРЕСС»\n", fontBold);
        cellValue.append("                      ТАЛОН\n");
        cellValue.append("                    № ");
        cellValue.append(number + "\n");
        cellValue.append("\n");
        cellValue.append("\n");
        cellValue.append("  Московская область, Люберецкий\n");
        cellValue.append("   район, п.Красково, д.Машково,\n");
        cellValue.append("    Кореневский тупик, в районе полей\n");
        cellValue.append("       аэрации, возле дома №5\n");
        cellValue.append("          Экземпляр заказчика\n");
        cellValue.append("\n");
        cellValue.append("  Вид отхода: ");
        cellValue.append("Древесные отходы\n", fontUnderline);
        cellValue.append("  Объем:        ");
        cellValue.append("         5 тонн       \n", fontUnderline);
        cellValue.append("\n");
        cellValue.append("\n");
        cellValue.append("             М.П.\n");
        cellValue.append("\n");
        cellValue.append("\n");
        return cellValue;
    }
}
