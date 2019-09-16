package com.supwisdom.spreadsheet.mapper.o2w.converter;

import org.joda.time.LocalDateTime;

/**
 * local date time text value with supplied pattern converter
 * Created by hanwen on 5/3/16.
 */
public class JodaLocalDateTimePropertyStringifier
    extends PropertyStringifierTemplate<LocalDateTime, JodaLocalDateTimePropertyStringifier> {

  private String pattern;

  public JodaLocalDateTimePropertyStringifier(String pattern) {
    this.pattern = pattern;
  }

  @Override
  public String convertProperty(Object property) {

    if (!LocalDateTime.class.equals(property.getClass())) {
      throw new IllegalArgumentException(
          "Not a Joda LocalDateTime: " + property.getClass().getName() + " value: " + property.toString());
    }

    return ((LocalDateTime) property).toString(pattern);
  }

}
