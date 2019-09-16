package com.supwisdom.spreadsheet.mapper.o2w.converter;

import org.joda.time.LocalDateTime;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import static org.testng.Assert.assertEquals;

/**
 * Created by qianjia on 2017/2/14.
 */
public class JodaLocalDateTimePropertyStringifierTest {

  @DataProvider
  public Object[][] testDoConvertParam() {
    LocalDateTime localDateTime = new LocalDateTime(1984, 11, 22, 0, 0, 0);
    return new Object[][] {
        new Object[] { localDateTime, "yyyy-MM-dd HH:mm:ss", "1984-11-22 00:00:00" },
        new Object[] { localDateTime, "yyyy-MM-dd", "1984-11-22" },
        new Object[] { localDateTime, "yyyy", "1984" }
    };
  }

  @Test(dataProvider = "testDoConvertParam")
  public void testDoConvert(LocalDateTime localDateTime, String pattern, String expected) throws Exception {

    JodaLocalDateTimePropertyStringifier converter = new JodaLocalDateTimePropertyStringifier(pattern);
    assertEquals(converter.convertProperty(localDateTime), expected);

  }
}
