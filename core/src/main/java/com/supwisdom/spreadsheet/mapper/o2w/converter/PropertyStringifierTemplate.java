package com.supwisdom.spreadsheet.mapper.o2w.converter;

import com.supwisdom.spreadsheet.mapper.bean.BeanHelper;
import com.supwisdom.spreadsheet.mapper.bean.BeanHelperBean;
import com.supwisdom.spreadsheet.mapper.model.core.Row;
import com.supwisdom.spreadsheet.mapper.model.meta.FieldMeta;
import com.supwisdom.spreadsheet.mapper.o2w.Object2WorkbookComposeException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @param <T> 要转换成{@link Row}的Object的类型
 * @param <V> 本类的类型
 */
public abstract class PropertyStringifierTemplate<T, V extends PropertyStringifierTemplate>
    implements PropertyStringifier<T> {

  protected final Logger logger = LoggerFactory.getLogger(getClass());

  protected String matchField;

  protected String nullString;

  protected BeanHelper beanHelper = new BeanHelperBean();

  /**
   * 匹配哪个{@link FieldMeta#getName()}
   *
   * @param matchField 匹配的{@link FieldMeta}
   * @return 自己
   */
  public V matchField(String matchField) {
    this.matchField = matchField;
    return (V) this;
  }

  /**
   * 如果object的property为null时，返回怎样的字符串的设置
   *
   * @param nullString 字符串
   * @return 自己
   */
  public V nullString(String nullString) {
    this.nullString = nullString;
    return (V) this;
  }

  @Override
  public String getMatchField() {
    return matchField;
  }

  @Override
  public String getPropertyString(T object, FieldMeta fieldMeta) {

    String propertyPath = fieldMeta.getName();
    Object propertyValue;
    try {
      propertyValue = beanHelper.getProperty(object, propertyPath);
    } catch (Exception e) {
      throw new Object2WorkbookComposeException("Sheet compose error", e);
    }
    if (propertyValue != null) {
      return convertProperty(propertyValue);
    }
    return nullString;

  }

  /**
   * 将object的某个property转换成String
   *
   * @param property 要转换成String的某个property，不会为null
   * @return 字符串
   */
  protected abstract String convertProperty(Object property);

}
