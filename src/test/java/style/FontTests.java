package style;

import com.dwl.excel.style.font.FontStyleApplier;
import com.dwl.excel.style.font.enums.FontCharsetValues;
import com.dwl.excel.style.font.enums.FontUnderlineValues;
import org.apache.poi.common.usermodel.fonts.FontCharset;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Order;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;

@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
public class FontTests {

  @Test
  @Order(1)
  @DisplayName("폰트 색상 테스트")
  public void fontColorTest(){
    //given
    Workbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), -1);
    Font font = workbook.createFont();

    //when
    FontStyleApplier fontStyler = FontStyleApplier.builder()
        .font(c -> c.color(0, 0, 255, 200))
        .build();

    fontStyler.apply(font);

    //then
    XSSFColor expectedFontColor = new XSSFColor(new java.awt.Color(0, 0, 255, 200), null);

    Assertions.assertInstanceOf(XSSFFont.class, font);
    Assertions.assertEquals(expectedFontColor, ((XSSFFont) font).getXSSFColor());
  }

  @Test
  @Order(2)
  @DisplayName("폰트 굵기 테스트")
  public void fontBoldTest(){
    //given
    Workbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), -1);
    Font font = workbook.createFont();

    //when
    FontStyleApplier fontStyler = FontStyleApplier.builder()
        .font(c -> c.bold())
        .build();

    fontStyler.apply(font);

    //then
    Assertions.assertInstanceOf(XSSFFont.class, font);
    Assertions.assertTrue(font.getBold());
  }

  @Test
  @Order(3)
  @DisplayName("폰트 기울기 테스트")
  public void fontItalicTest(){
    //given
    Workbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), -1);
    Font font = workbook.createFont();

    //when
    FontStyleApplier fontStyler = FontStyleApplier.builder()
        .font(c -> c.italic())
        .build();

    fontStyler.apply(font);

    //then
    Assertions.assertInstanceOf(XSSFFont.class, font);
    Assertions.assertTrue(font.getItalic());
  }

  @Test
  @Order(4)
  @DisplayName("폰트 취소선 테스트")
  public void fontStrikeOutTest(){
    //given
    Workbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), -1);
    Font font = workbook.createFont();

    //when
    FontStyleApplier fontStyler = FontStyleApplier.builder()
        .font(c -> c.strikeOut())
        .build();

    fontStyler.apply(font);

    //then
    Assertions.assertInstanceOf(XSSFFont.class, font);
    Assertions.assertTrue(font.getStrikeout());
  }

  @Test
  @Order(5)
  @DisplayName("폰트 밑줄 테스트")
  public void fontUnderlineTest(){
    //given
    Workbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), -1);
    Font font = workbook.createFont();

    //when
    FontStyleApplier fontStyler = FontStyleApplier.builder()
        .font(c -> c.underline(FontUnderlineValues.SINGLE))
        .build();

    fontStyler.apply(font);

    //then
    byte expectedFontUnderline = FontUnderline.SINGLE.getByteValue();

    Assertions.assertInstanceOf(XSSFFont.class, font);
    Assertions.assertEquals(expectedFontUnderline, font.getUnderline());
  }

  @Test
  @Order(6)
  @DisplayName("폰트 크기 테스트")
  public void fontSizeTest(){
    //given
    Workbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), -1);
    Font font = workbook.createFont();

    //when
    FontStyleApplier fontStyler = FontStyleApplier.builder()
        .font(c -> c.sizeInPoints((short)12))
        .build();

    fontStyler.apply(font);

    //then
    short expectedSizeInPoints = (short)12;

    Assertions.assertInstanceOf(XSSFFont.class, font);
    Assertions.assertEquals(expectedSizeInPoints, font.getFontHeightInPoints());
  }

  @Test
  @Order(7)
  @DisplayName("폰트 CharSet(글꼴 언어/지역) 테스트")
  public void fontCharSetTest(){
    //given
    Workbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), -1);
    Font font = workbook.createFont();

    //when
    FontStyleApplier fontStyler = FontStyleApplier.builder()
        .font(c -> c.charSet(FontCharsetValues.ANSI))
        .build();

    fontStyler.apply(font);

    //then
    int expectedCharSet = FontCharset.ANSI.getNativeId();
    Assertions.assertInstanceOf(XSSFFont.class, font);
    Assertions.assertEquals(expectedCharSet, font.getCharSet());
  }

  @Test
  @Order(8)
  @DisplayName("폰트 전체 테스트")
  public void fontTotalTest(){
    //given
    Workbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), -1);
    Font font = workbook.createFont();

    //when
    FontStyleApplier fontStyler = FontStyleApplier.builder()
        .font(c ->
            c.color(150, 100, 255, 50)
                .bold()
                .italic()
                .strikeOut()
                .underline(FontUnderlineValues.SINGLE)
                .sizeInPoints((short)12)
                .charSet(FontCharsetValues.HANGUL)
        )
        .build();

    fontStyler.apply(font);

    //then
    Assertions.assertInstanceOf(XSSFFont.class, font);

    XSSFColor expectedFontColor = new XSSFColor(new java.awt.Color(150, 100, 255, 50), null);
    Assertions.assertEquals(expectedFontColor, ((XSSFFont) font).getXSSFColor());

    short expectedSizeInPoints = (short)12;
    Assertions.assertEquals(expectedSizeInPoints, font.getFontHeightInPoints());

    Assertions.assertTrue(font.getBold());

    Assertions.assertTrue(font.getItalic());

    byte expectedFontUnderline = FontUnderline.SINGLE.getByteValue();
    Assertions.assertEquals(expectedFontUnderline, font.getUnderline());

    Assertions.assertTrue(font.getStrikeout());

    int expectedCharSet = FontCharset.HANGUL.getNativeId();
    Assertions.assertEquals(expectedCharSet, font.getCharSet());
  }
}
