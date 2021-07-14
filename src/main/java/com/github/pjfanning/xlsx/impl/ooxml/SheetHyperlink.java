package com.github.pjfanning.xlsx.impl.ooxml;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.sl.usermodel.Hyperlink;
import org.apache.poi.sl.usermodel.Slide;

public class SheetHyperlink implements Hyperlink {

  private String label;

  @Override
  public void linkToEmail(String s) {

  }

  @Override
  public void linkToUrl(String s) {

  }

  @Override
  public void linkToSlide(Slide slide) {

  }

  @Override
  public void linkToNextSlide() {

  }

  @Override
  public void linkToPreviousSlide() {

  }

  @Override
  public void linkToFirstSlide() {

  }

  @Override
  public void linkToLastSlide() {

  }

  @Override
  public String getAddress() {
    return null;
  }

  /**
   * update operations are not supported
   */
  @Override
  public void setAddress(String s) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  @Override
  public String getLabel() {
    return label;
  }

  /**
   * update operations are not supported
   */
  @Override
  public void setLabel(String s) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  @Override
  public HyperlinkType getType() {
    return null;
  }
}
