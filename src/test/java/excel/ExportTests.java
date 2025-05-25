package excel;

import com.dwl.excel.exporter.ExcelExporter;
import com.dwl.excel.exporter.SXSSFExcelExporter;
import com.dwl.excel.exporter.writer.CellWriter;
import com.dwl.excel.style.CellStyleApplier;
import com.dwl.excel.style.enums.BorderStyleValues;
import com.dwl.excel.style.enums.HorizontalAlignmentValues;
import com.dwl.excel.style.enums.IndexedColorValues;
import com.dwl.excel.style.enums.VerticalAlignmentValues;
import com.dwl.excel.style.font.FontStyleApplier;
import creator.DataCreator;
import dto.MemberDto;
import dto.TravelDto;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Order;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.time.LocalDateTime;
import java.util.List;

public class ExportTests {

  @Test
  @Order(1)
  @DisplayName("엑셀 커스텀 객체 및 병합 테스트1")
  public void cellCustomHeaderBodyTest1() throws Exception {
    //given
    int memberCount = 10, travelCount = 10;
    String[] groupHeader = new String[]{"회원", "여행"};
    Class<MemberDto> memberClass = MemberDto.class;
    Class<TravelDto> travelClass = TravelDto.class;
    List<MemberDto> memberTravels = DataCreator.findMemberAndTravels(memberCount, travelCount);

    FontStyleApplier fontStyleApplier = FontStyleApplier.builder()
        .font(c ->
            c.bold().sizeInPoints((short) 12)
        )
        .build();

    CellStyleApplier defaultHeaderCellStyleApplier = CellStyleApplier.builder()
        .wrap(c -> c.enable())
        .border(c -> c.styleAll(BorderStyleValues.THIN))
        .foreground(c -> c.color(IndexedColorValues.PALE_BLUE))
        .alignment(c ->
            c.horizontal(HorizontalAlignmentValues.CENTER)
                .vertical(VerticalAlignmentValues.CENTER)
        )
        .fontStyleApplier(fontStyleApplier)
        .build();

    ExcelExporter exporter = SXSSFExcelExporter.builder()
        .defaultHeaderCellStyleApplier(defaultHeaderCellStyleApplier)
        .defaultBodyCellStyleApplier(null)
        .flushCount(100)
        .build();

    exporter.createSheet("시트1");

    exporter
        .mergeCell(0,0,0,2)
        .mergeCell(0,0,3,6)
        .createHeader().writeCells(groupHeader);

    exporter.createHeader()
        .writeTargetHeaders(memberClass)
        .writeTargetHeaders(travelClass);

    for(MemberDto member : memberTravels){
      List<TravelDto> travels = member.getTravels();

      int travelSize = CollectionUtils.size(travels);
      int currentRow = exporter.getCurrentRow();

      int mergeStartRow = currentRow;
      int mergeEndRow = mergeStartRow+travelSize;

      TravelDto firstTravel = null;
      if(travelSize > 0){
        firstTravel = travels.get(0);
        mergeEndRow--;
      }

      exporter
          .mergeCell(mergeStartRow, mergeEndRow, 0, 0)
          .mergeCell(mergeStartRow, mergeEndRow, 1, 1)
          .mergeCell(mergeStartRow, mergeEndRow, 2, 2)
          .createRow()
          .writeTargetCells(member)
          .writeTargetCells(firstTravel);

      for(int i = 1 ; i < travelSize ; i++){
        exporter.createRow().writeTargetCells(travels.get(i));
      }
    }

    File file = new File("export.xlsx");
    OutputStream stream = new FileOutputStream(file);
    exporter.export(stream);

    //when
    FileInputStream fis = new FileInputStream(file);
    Workbook workbook = new XSSFWorkbook(fis);

    int sheetCount = workbook.getNumberOfSheets();
    int styleCount = workbook.getNumCellStyles();
    int fontCount = workbook.getNumberOfFonts();

    Sheet sheet = workbook.getSheetAt(0);
    int lastRowNum = sheet.getPhysicalNumberOfRows();
    int mergedRegionCount = sheet.getNumMergedRegions();

    CellStyle cellStyle = workbook.getCellStyleAt(1);
    Font font = workbook.getFontAt(1);

    file.delete();

    //then
    int expectedSheetCount = 1;
    int expectedStyleCount = 2;
    int expectedFontCount = 2;
    int expectedLastRowNum = 102;
    int expectedMergedRegionCount = 32;
    short expectedFontSizeInPoints = 12;

    short expectedBorderTop = BorderStyleValues.THIN.getPoiBorderStyle().getCode();
    short expectedBorderLeft = BorderStyleValues.THIN.getPoiBorderStyle().getCode();
    short expectedBorderBottom = BorderStyleValues.THIN.getPoiBorderStyle().getCode();
    short expectedBorderRight = BorderStyleValues.THIN.getPoiBorderStyle().getCode();
    short expectedForegroundColor = IndexedColorValues.PALE_BLUE.getPoiIndexedColors().getIndex();
    short expectedHorizontalCode = HorizontalAlignmentValues.CENTER.getPoiHorizontal().getCode();
    short expectedVerticalCode = VerticalAlignmentValues.CENTER.getPoiVertical().getCode();

    Assertions.assertEquals(expectedSheetCount, sheetCount);
    Assertions.assertEquals(expectedStyleCount, styleCount);
    Assertions.assertEquals(expectedFontCount, fontCount);
    Assertions.assertEquals(expectedLastRowNum, lastRowNum);
    Assertions.assertEquals(expectedMergedRegionCount, mergedRegionCount);

    Assertions.assertTrue(font.getBold());
    Assertions.assertEquals(expectedFontSizeInPoints, font.getFontHeightInPoints());

    Assertions.assertTrue(cellStyle.getWrapText());
    Assertions.assertEquals(expectedBorderTop, cellStyle.getBorderTop().getCode());
    Assertions.assertEquals(expectedBorderLeft, cellStyle.getBorderLeft().getCode());
    Assertions.assertEquals(expectedBorderBottom, cellStyle.getBorderBottom().getCode());
    Assertions.assertEquals(expectedBorderRight, cellStyle.getBorderRight().getCode());
    Assertions.assertEquals(expectedForegroundColor, cellStyle.getFillForegroundColor());
    Assertions.assertEquals(expectedHorizontalCode, cellStyle.getAlignment().getCode());
    Assertions.assertEquals(expectedVerticalCode, cellStyle.getVerticalAlignment().getCode());
  }

  @Test
  @Order(2)
  @DisplayName("엑셀 비정형 커스텀 객체 및 병합 테스트2")
  public void cellCustomHeaderBodyTest2() throws Exception {
    //given
    int memberCount = 10, travelCount = 100;
    String[] groupHeader = new String[]{"회원", "여행"};
    Class<MemberDto> memberClass = MemberDto.class;
    Class<TravelDto> travelClass = TravelDto.class;
    List<MemberDto> members = DataCreator.findMembers(memberCount);
    List<TravelDto> travels = DataCreator.findTravels(travelCount);

    FontStyleApplier fontStyleApplier = FontStyleApplier.builder()
        .font(c ->
            c.bold().sizeInPoints((short) 12)
        )
        .build();

    CellStyleApplier defaultHeaderCellStyleApplier = CellStyleApplier.builder()
        .wrap(c -> c.enable())
        .border(c -> c.styleAll(BorderStyleValues.THIN))
        .foreground(c -> c.color(IndexedColorValues.PALE_BLUE))
        .alignment(c ->
            c.horizontal(HorizontalAlignmentValues.CENTER)
                .vertical(VerticalAlignmentValues.CENTER)
        )
        .fontStyleApplier(fontStyleApplier)
        .build();

    CellStyleApplier redFontStyleApplier = CellStyleApplier.builder()
        .foreground(c -> c.color(IndexedColorValues.GREY_80_PERCENT))
        .fontStyleApplier(
            FontStyleApplier.builder()
              .font(c -> c.bold().color(IndexedColorValues.RED))
              .build()
        )
        .build();

    CellStyleApplier blueFontStyleApplier = CellStyleApplier.builder()
        .foreground(c -> c.color(IndexedColorValues.GREY_80_PERCENT))
        .fontStyleApplier(
            FontStyleApplier.builder()
                .font(c -> c.bold().color(IndexedColorValues.BLUE))
                .build()
        )
        .build();

    ExcelExporter exporter = SXSSFExcelExporter.builder()
        .defaultHeaderCellStyleApplier(defaultHeaderCellStyleApplier)
        .flushCount(100)
        .build();

    exporter.createSheet("시트1");

    exporter
        .createCellStyle("GREY_RED", redFontStyleApplier)
        .createCellStyle("GREY_BLUE", blueFontStyleApplier);

    exporter
        .mergeCell(0, 0, 0, 2)
        .mergeCell(0, 0, 4, 7)
        .createHeader()
        .writeCell(groupHeader[0])
        .nextCol()
        .writeCell(groupHeader[1]);

    exporter.createHeader()
        .writeTargetHeaders(memberClass)
        .nextCol()
        .writeTargetHeaders(travelClass);

    int size = Math.max(members.size(), travels.size());
    LocalDateTime today = LocalDateTime.now();
    for (int i = 0; i < size; i++){
      CellWriter writer = exporter.createRow();
      if (i < members.size()){
        writer.writeTargetCells(members.get(i)).nextCol();
      }
      else {
        writer.nextCol(4);
      }

      if (i < travels.size()) {
        TravelDto travel = travels.get(i);
        writer.writeCell(travel.getTravelName());
        String styleKey = null;
        if(today.isBefore(travel.getStartDate())) {
          styleKey = "GREY_BLUE";
        }
        else if(today.isAfter(travel.getEndDate())) {
          styleKey = "GREY_RED";
        }
        writer.writeCell(styleKey, travel.getStartDate())
            .writeCell(styleKey, travel.getEndDate())
            .writeCell(travel.getCreateDate());

        writer.writeTargetCells(travels.get(i));
      }
    }

    File file = new File("export2.xlsx");
    OutputStream stream = new FileOutputStream(file);
    exporter.export(stream);

    //when
    FileInputStream fis = new FileInputStream(file);
    Workbook workbook = new XSSFWorkbook(fis);

    int sheetCount = workbook.getNumberOfSheets();
    int styleCount = workbook.getNumCellStyles();
    int fontCount = workbook.getNumberOfFonts();

    Sheet sheet = workbook.getSheetAt(0);
    int lastRowNum = sheet.getPhysicalNumberOfRows();
    int mergedRegionCount = sheet.getNumMergedRegions();

    CellStyle defaultCellStyle = workbook.getCellStyleAt(1);
    Font defaultFont = workbook.getFontAt(1);

    CellStyle redCellStyle = workbook.getCellStyleAt(2);
    Font redFont = workbook.getFontAt(2);

    CellStyle blueCellStyle = workbook.getCellStyleAt(3);
    Font blueFont = workbook.getFontAt(3);

    file.delete();

    //then
    int expectedSheetCount = 1;
    int expectedStyleCount = 4;
    int expectedFontCount = 4;
    int expectedLastRowNum = 102;
    int expectedMergedRegionCount = 2;
    short expectedFontSizeInPoints = 12;

    short expectedBorderTop = BorderStyle.THIN.getCode();
    short expectedBorderLeft = BorderStyle.THIN.getCode();
    short expectedBorderBottom = BorderStyle.THIN.getCode();
    short expectedBorderRight = BorderStyle.THIN.getCode();
    short expectedForegroundColor = IndexedColorValues.PALE_BLUE.getPoiIndexedColors().getIndex();
    short expectedHorizontalCode = HorizontalAlignmentValues.CENTER.getPoiHorizontal().getCode();
    short expectedVerticalCode = VerticalAlignmentValues.CENTER.getPoiVertical().getCode();

    short expectedForegroundColor2 = IndexedColorValues.GREY_80_PERCENT.getPoiIndexedColors().getIndex();
    short expectedFontColor2 = IndexedColorValues.RED.getPoiIndexedColors().getIndex();

    short expectedForegroundColor3 = IndexedColorValues.GREY_80_PERCENT.getPoiIndexedColors().getIndex();
    short expectedFontColor3 = IndexedColorValues.BLUE.getPoiIndexedColors().getIndex();

    Assertions.assertEquals(expectedSheetCount, sheetCount);
    Assertions.assertEquals(expectedStyleCount, styleCount);
    Assertions.assertEquals(expectedFontCount, fontCount);
    Assertions.assertEquals(expectedLastRowNum, lastRowNum);
    Assertions.assertEquals(expectedMergedRegionCount, mergedRegionCount);

    Assertions.assertTrue(defaultFont.getBold());
    Assertions.assertEquals(expectedFontSizeInPoints, defaultFont.getFontHeightInPoints());

    Assertions.assertTrue(defaultCellStyle.getWrapText());
    Assertions.assertEquals(expectedBorderTop, defaultCellStyle.getBorderTop().getCode());
    Assertions.assertEquals(expectedBorderLeft, defaultCellStyle.getBorderLeft().getCode());
    Assertions.assertEquals(expectedBorderBottom, defaultCellStyle.getBorderBottom().getCode());
    Assertions.assertEquals(expectedBorderRight, defaultCellStyle.getBorderRight().getCode());
    Assertions.assertEquals(expectedForegroundColor, defaultCellStyle.getFillForegroundColor());
    Assertions.assertEquals(expectedHorizontalCode, defaultCellStyle.getAlignment().getCode());
    Assertions.assertEquals(expectedVerticalCode, defaultCellStyle.getVerticalAlignment().getCode());

    Assertions.assertEquals(expectedForegroundColor2, redCellStyle.getFillForegroundColor());
    Assertions.assertEquals(expectedFontColor2, redFont.getColor());
    Assertions.assertEquals(expectedForegroundColor3, blueCellStyle.getFillForegroundColor());
    Assertions.assertEquals(expectedFontColor3, blueFont.getColor());
  }
}
