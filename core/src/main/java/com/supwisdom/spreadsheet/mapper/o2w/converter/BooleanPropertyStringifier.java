package com.supwisdom.spreadsheet.mapper.o2w.converter;

/**
 * boolean readable text value converter
 * Created by hanwen on 16/3/18.
 */
public class BooleanPropertyStringifier extends PropertyStringifierTemplate<Object, BooleanPropertyStringifier> {

  private String trueString;

  private String falseString;

  public BooleanPropertyStringifier trueString(String trueString) {
    this.trueString = trueString;
    return this;
  }

  public BooleanPropertyStringifier falseString(String falseString) {
    this.falseString = falseString;
    return this;
  }

  @Override
  public String convertProperty(Object property) {

    Class<?> propertyClass = property.getClass();
    if (!Boolean.class.equals(propertyClass) && !Boolean.TYPE.equals(propertyClass)) {
      throw new IllegalArgumentException("Not a boolean: " + propertyClass.getName() + " value: " + property.toString());
    }
    return Boolean.TRUE.equals(property) ? trueString : falseString;

  }

}
