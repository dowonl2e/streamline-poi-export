package com.dwl.excel.style;

import com.dwl.excel.style.configurer.AlignmentConfigurer;
import com.dwl.excel.style.configurer.BorderConfigurer;
import com.dwl.excel.style.configurer.ForegroundConfigurer;
import com.dwl.excel.style.configurer.WrapConfigurer;
import com.dwl.excel.style.font.FontStyleApplier;
import com.dwl.excel.style.functional.AlignmentCustomizer;
import com.dwl.excel.style.functional.BorderCustomizer;
import com.dwl.excel.style.functional.ForegroundCustomizer;
import com.dwl.excel.style.functional.WrapCustomizer;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import java.util.Objects;

public class CellStyleApplier {

  private AlignmentCustomizer alignmentCustomizer;
  private BorderCustomizer borderCustomizer;
  private ForegroundCustomizer foregroundCustomizer;
  private WrapCustomizer wrapCustomizer;
  private FontStyleApplier fontStyleApplier;

  public static Builder builder() {
    return new Builder();
  }

  public static class Builder {
    private AlignmentCustomizer alignmentCustomizer;
    private BorderCustomizer borderCustomizer;
    private ForegroundCustomizer foregroundCustomizer;
    private WrapCustomizer wrapCustomizer;
    private FontStyleApplier fontStyleApplier;

    public Builder alignment(AlignmentCustomizer alignmentCustomizer) {
      this.alignmentCustomizer = alignmentCustomizer;
      return this;
    }

    public Builder border(BorderCustomizer borderCustomizer) {
      this.borderCustomizer = borderCustomizer;
      return this;
    }

    public Builder foreground(ForegroundCustomizer foregroundCustomizer) {
      this.foregroundCustomizer = foregroundCustomizer;
      return this;
    }

    public Builder wrap(WrapCustomizer wrapCustomizer) {
      this.wrapCustomizer = wrapCustomizer;
      return this;
    }

    public Builder fontStyleApplier(FontStyleApplier fontStyleApplier) {
      this.fontStyleApplier = fontStyleApplier;
      return this;
    }

    public CellStyleApplier build() {
      return new CellStyleApplier()
          .alignment(alignmentCustomizer)
          .border(borderCustomizer)
          .foreground(foregroundCustomizer)
          .wrap(wrapCustomizer)
          .fontStyleApplier(fontStyleApplier);
    }
  }

  public FontStyleApplier getFontStyleApplier(){
    return this.fontStyleApplier;
  }

  public CellStyleApplier alignment(AlignmentCustomizer alignmentCustomizer) {
    this.alignmentCustomizer = alignmentCustomizer;
    return this;
  }

  public CellStyleApplier border(BorderCustomizer borderCustomizer) {
    this.borderCustomizer = borderCustomizer;
    return this;
  }

  public CellStyleApplier foreground(ForegroundCustomizer foregroundCustomizer) {
    this.foregroundCustomizer = foregroundCustomizer;
    return this;
  }

  public CellStyleApplier wrap(WrapCustomizer wrapCustomizer) {
    this.wrapCustomizer = wrapCustomizer;
    return this;
  }

  public CellStyleApplier fontStyleApplier(FontStyleApplier fontStyleApplier) {
    this.fontStyleApplier = fontStyleApplier;
    return this;
  }

  public void apply(Font font, CellStyle cellStyle) {
    Objects.requireNonNull(cellStyle, "cellStyle can not be null");
    if (this.alignmentCustomizer != null) {
      AlignmentConfigurer alignmentConfigurator = new AlignmentConfigurer();
      this.alignmentCustomizer.customize(alignmentConfigurator);
      alignmentConfigurator.apply(cellStyle);
    }
    if (this.borderCustomizer != null) {
      BorderConfigurer borderConfigurator = new BorderConfigurer();
      this.borderCustomizer.customize(borderConfigurator);
      borderConfigurator.apply(cellStyle);
    }
    if (this.foregroundCustomizer != null) {
      ForegroundConfigurer foregroundConfigurator = new ForegroundConfigurer();
      this.foregroundCustomizer.customize(foregroundConfigurator);
      foregroundConfigurator.apply(cellStyle);
    }
    if (this.wrapCustomizer != null) {
      WrapConfigurer wrapConfigurator = new WrapConfigurer();
      this.wrapCustomizer.customize(wrapConfigurator);
      wrapConfigurator.apply(cellStyle);
    }
    if (this.fontStyleApplier != null && font != null) {
      this.fontStyleApplier.apply(font);
      cellStyle.setFont(font);
    }
  }
}
