package com.dwl.excel.style.font;

import com.dwl.excel.style.font.configurer.FontConfigurer;
import com.dwl.excel.style.font.functional.FontCustomizer;
import org.apache.poi.ss.usermodel.Font;

import java.util.Objects;

public class FontStyleApplier {

  private FontCustomizer fontCustomizer;

  public static FontStyleApplier.Builder builder() {
    return new FontStyleApplier.Builder();
  }
  public static class Builder {
    private FontCustomizer fontCustomizer;

    public Builder font(FontCustomizer fontCustomizer) {
      this.fontCustomizer = fontCustomizer;
      return this;
    }

    public FontStyleApplier build() {
      return new FontStyleApplier()
          .setFontCustomizer(fontCustomizer);
    }
  }

  public FontStyleApplier setFontCustomizer(FontCustomizer fontCustomizer) {
    this.fontCustomizer = fontCustomizer;
    return this;
  }

  public void apply(Font font) {
    Objects.requireNonNull(font, "cellStyle can not be null");
    if (this.fontCustomizer != null) {
      FontConfigurer fontConfigurator = new FontConfigurer();
      this.fontCustomizer.customize(fontConfigurator);
      fontConfigurator.apply(font);
    }
  }
}
