package com.supwisdom.spreadsheet.mapper.o2w.converter;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.time.LocalDate;

import static org.testng.Assert.assertEquals;

/**
 * Created by qianjia on 2017/2/14.
 */
public class LocalDatePropertyStringifierTest {

  @DataProvider
  public Object[][] testDoConvertParam() {
    LocalDate localDate = LocalDate.of(1984, 11, 22);
    return new Object[][] {
        new Object[] { localDate, "yyyy-MM-dd", "1984-11-22" },
        new Object[] { localDate, "yyyy-MM", "1984-11" },
        new Object[] { localDate, "yyyy", "1984" }
    };
  }

  @Test(dataProvider = "testDoConvertParam")
  public void testDoConvert(LocalDate localDate, String pattern, String expected) throws Exception {

    LocalDatePropertyStringifier converter = new LocalDatePropertyStringifier(pattern);
    assertEquals(converter.convertProperty(localDate), expected);

  }

}
