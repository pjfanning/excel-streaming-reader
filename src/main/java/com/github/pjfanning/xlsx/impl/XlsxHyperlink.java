package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.xlsx.impl.ooxml.HyperlinkData;
import org.apache.poi.common.Duplicatable;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;

import java.net.URI;
import java.util.Objects;

/**
 * A read-only implementation of Hyperlink
 */
public class XlsxHyperlink implements Hyperlink, Duplicatable {
  private final HyperlinkType _type;
  private final PackageRelationship _externalRel;
  private final HyperlinkData hyperlinkData; //contains a reference to the cell where the hyperlink is anchored, getRef()
  private String _address; //what the hyperlink refers to

  /**
   * Create a XlsxHyperlink and initialize it from the supplied HyperlinkData bean and package relationship
   *
   * @param hyperlinkData the bean containing xml properties
   * @param hyperlinkRel the relationship in the underlying OPC package which stores the actual link's address
   */
  XlsxHyperlink(HyperlinkData hyperlinkData, PackageRelationship hyperlinkRel) {
    this.hyperlinkData = hyperlinkData;
    _externalRel = hyperlinkRel;

    // Figure out the Hyperlink type and destination

    if (_externalRel == null) {
      // If it has a location, it's internal
      if (hyperlinkData.getLocation() != null) {
        _type = HyperlinkType.DOCUMENT;
        _address = hyperlinkData.getLocation();
      } else if (hyperlinkData.getId() != null) {
        throw new IllegalStateException("The hyperlink for cell "
                + hyperlinkData.getRef() + " references relation "
                + hyperlinkData.getId() + ", but that didn't exist!");
      } else {
        // hyperlink is internal and is not related to other parts
        _type = HyperlinkType.DOCUMENT;
      }
    } else {
      URI target = _externalRel.getTargetURI();
      _address = target.toString();
      if (hyperlinkData.getLocation() != null) {
        // URI fragment
        _address += "#" + hyperlinkData.getLocation();
      }

      // Try to figure out the type
      if (_address.startsWith("http://") || _address.startsWith("https://")
              || _address.startsWith("ftp://")) {
        _type = HyperlinkType.URL;
      } else if (_address.startsWith("mailto:")) {
        _type = HyperlinkType.EMAIL;
      } else {
        _type = HyperlinkType.FILE;
      }
    }

  }

  /**
   * Return the type of this hyperlink
   *
   * @return the type of this hyperlink
   */
  @Override
  public HyperlinkType getType() {
    return _type;
  }

  /**
   * Get the address of the cell this hyperlink applies to, e.g. A55
   */
  public String getCellRef() {
    return hyperlinkData.getRef();
  }

  /**
   * Hyperlink address. Depending on the hyperlink type it can be URL, e-mail, path to a file.
   * This is the hyperlink target.
   *
   * @return the address of this hyperlink
   */
  @Override
  public String getAddress() {
    return _address;
  }

  private String getAddressWithoutLocation() {
    String addr = _address;
    String locn = getLocation();
    if (addr != null && !addr.equals(locn) && addr.endsWith(locn)) {
      return addr.substring(0, addr.length() - locn.length() - 1);
    }
    return addr;
  }

  /**
   * Return text label for this hyperlink
   *
   * @return text to display
   */
  @Override
  public String getLabel() {
    return hyperlinkData.getDisplay();
  }

  /**
   * Location within target. If target is a workbook (or this workbook) this shall refer to a
   * sheet and cell or a defined name. Can also be an HTML anchor if target is HTML file.
   *
   * @return location
   */
  public String getLocation() {
    return hyperlinkData.getLocation();
  }

  private CellReference buildFirstCellReference() {
    return buildCellReference(false);
  }

  private CellReference buildLastCellReference() {
    return buildCellReference(true);
  }

  private CellReference buildCellReference(boolean lastCell) {
    String ref = hyperlinkData.getRef();
    if (ref == null) {
      ref = "A1";
    }
    if (ref.contains(":")) {
      AreaReference area = new AreaReference(ref, SpreadsheetVersion.EXCEL2007);
      return lastCell ? area.getLastCell() : area.getFirstCell();
    }
    return new CellReference(ref);
  }

  /**
   * Return the column of the first cell that contains the hyperlink
   *
   * @return the 0-based column of the first cell that contains the hyperlink
   */
  @Override
  public int getFirstColumn() {
    return buildFirstCellReference().getCol();
  }


  /**
   * Return the column of the last cell that contains the hyperlink
   *
   * @return the 0-based column of the last cell that contains the hyperlink
   */
  @Override
  public int getLastColumn() {
    return buildLastCellReference().getCol();
  }

  /**
   * Return the row of the first cell that contains the hyperlink
   *
   * @return the 0-based row of the cell that contains the hyperlink
   */
  @Override
  public int getFirstRow() {
    return buildFirstCellReference().getRow();
  }


  /**
   * Return the row of the last cell that contains the hyperlink
   *
   * @return the 0-based row of the last cell that contains the hyperlink
   */
  @Override
  public int getLastRow() {
    return buildLastCellReference().getRow();
  }

  /**
   * @return additional text to help the user understand more about the hyperlink
   */
  public String getTooltip() {
    return hyperlinkData.getTooltip();
  }

  /**
   * @throws UnsupportedOperationException
   */
  @Override
  public void setAddress(String address) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * @throws UnsupportedOperationException
   */
  @Override
  public void setLabel(String label) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * @throws UnsupportedOperationException
   */
  @Override
  public void setFirstColumn(int col) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * @throws UnsupportedOperationException
   */
  @Override
  public void setLastColumn(int col) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * @throws UnsupportedOperationException
   */
  @Override
  public void setFirstRow(int row) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * @throws UnsupportedOperationException
   */
  @Override
  public void setLastRow(int row) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  @Override
  public boolean equals(Object o) {
    if (this == o) return true;
    if (o == null || getClass() != o.getClass()) return false;
    XlsxHyperlink that = (XlsxHyperlink) o;
    return _type == that._type && Objects.equals(_externalRel, that._externalRel) && Objects.equals(hyperlinkData, that.hyperlinkData) && Objects.equals(_address, that._address);
  }

  @Override
  public int hashCode() {
    return Objects.hash(_type, _externalRel, hyperlinkData, _address);
  }

  /**
   * @return a copy of this XlsxHyperlink instance
   * @since 4.0.0
   */
  @Override
  public Duplicatable copy() {
    return new XlsxHyperlink(hyperlinkData, _externalRel);
  }

  /**
   * @return a copy of this XlsxHyperlink instance but as a XSSFHyperlink
   * @since 4.0.0
   */
  public XSSFHyperlink createXSSFHyperlink(){
    XSSFHyperlink xssfHyperlink = new XSSFHyperlink(getType()) {};
    xssfHyperlink.setFirstRow(getFirstRow());
    xssfHyperlink.setLastRow(getLastRow());
    xssfHyperlink.setFirstColumn(getFirstColumn());
    xssfHyperlink.setLastColumn(getLastColumn());
    xssfHyperlink.setLabel(getLabel());
    xssfHyperlink.setTooltip(getTooltip());
    xssfHyperlink.setAddress(getAddressWithoutLocation());
    xssfHyperlink.setLocation(getLocation());
    return xssfHyperlink;
  }
}
