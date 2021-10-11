package com.github.pjfanning.xlsx.impl;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

public class ReadOnlyComment implements Comment {

  private final XSSFComment xssfComment;

  ReadOnlyComment(XSSFComment xssfComment) {
    this.xssfComment = xssfComment;
  }

  @Override
  public boolean isVisible() {
    return xssfComment.isVisible();
  }

  @Override
  public CellAddress getAddress() {
    return xssfComment.getAddress();
  }

  @Override
  public int getRow() {
    return xssfComment.getRow();
  }

  @Override
  public int getColumn() {
    return xssfComment.getColumn();
  }

  @Override
  public String getAuthor() {
    return xssfComment.getAuthor();
  }

  @Override
  public ClientAnchor getClientAnchor() {
    return xssfComment.getClientAnchor();
  }

  @Override
  public XSSFRichTextString getString() {
    XSSFRichTextString rts = xssfComment.getString();
    String text = rts.getString();
    if(rts.getString().contains("Your version of Excel allows you to read this threaded comment")) {
      String splitText = "Comment:";
      int pos = text.lastIndexOf(splitText);
      if (pos != -1) {
        return new XSSFRichTextString(ltrim(text.substring(pos + splitText.length())));
      }
    }
    return rts;
  }

  @Override
  public void setAddress(CellAddress addr) {

  }

  @Override
  public void setVisible(boolean visible) {

  }

  @Override
  public void setAddress(int row, int col) {

  }

  @Override
  public void setRow(int row) {

  }

  @Override
  public void setColumn(int col) {

  }

  @Override
  public void setAuthor(String author) {

  }

  @Override
  public void setString(RichTextString string) {

  }

  private String ltrim(String s) {
    int i = 0;
    while (i < s.length() && Character.isWhitespace(s.charAt(i))) {
      i++;
    }
    return s.substring(i);
  }
}
