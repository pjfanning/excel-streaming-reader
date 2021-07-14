package com.github.pjfanning.xlsx.impl.ooxml;

public class HyperLinkData {

  private final String id;
  private final String ref;
  private final String location;
  private final String display;

  public HyperLinkData(String id, String ref, String location, String display) {
    this.id = id;
    this.ref = ref;
    this.location = location;
    this.display = display;
  }

  public String getId() {
    return id;
  }

  public String getRef() {
    return ref;
  }

  public String getLocation() {
    return location;
  }
}
