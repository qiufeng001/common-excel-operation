# spreadsheet-mapper
[![Build Status](https://travis-ci.org/supwisdom/spreadsheet-mapper.svg?branch=develop)](https://travis-ci.org/supwisdom/spreadsheet-mapper)
[![codecov](https://codecov.io/gh/supwisdom/spreadsheet-mapper/branch/develop/graph/badge.svg)](https://codecov.io/gh/supwisdom/spreadsheet-mapper)
[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.supwisdom/spreadsheet-mapper-core/badge.svg)](https://maven-badges.herokuapp.com/maven-central/com.supwisdom/spreadsheet-mapper-core)
[![License](https://img.shields.io/badge/license-Apache%202-4EB1BA.svg)](https://www.apache.org/licenses/LICENSE-2.0.html)

## Maven

```xml
<dependency>
    <groupId>com.supwisdom</groupId>
    <artifactId>spreadsheet-mapper-core</artifactId>
    <version>1.0.9</version>
</dependency>

<!-- （可选）对joda的支持 -->
<dependency>
    <groupId>com.supwisdom</groupId>
    <artifactId>spreadsheet-mapper-joda</artifactId>
    <version>1.0.9</version>
</dependency>

<!-- （可选）对java8的支持 -->
<dependency>
    <groupId>com.supwisdom</groupId>
    <artifactId>spreadsheet-mapper-java8</artifactId>
    <version>1.0.9</version>
</dependency>

```

## Gradle

``` groovy
compile 'com.supwisdom:spreadsheet-mapper-core:1.0.9'
compile 'com.supwisdom:spreadsheet-mapper-java8:1.0.9'
compile 'com.supwisdom:spreadsheet-mapper-joda:1.0.9'
```

## 项目一览
 
用excel导入、导出数据是企业软件里很常见的需求，但是要实现好导入、导出并非易事，主要原因在于excel导入或导出并非简单的对数据库表的操作，往往牵涉到校验、转换等工作。
本项目希望提供一套方便好用的库，能够使得此类需求实现起来更为便捷。

### spreadsheet-mapper-core

本包提供了excel文件导入、导出的基础工具。

和导入相关的工具：

1. package f2w：读取excel文件并生成Workbook。
1. package validation：Workbook的校验工具。
1. package w2o：将Workbook转换成Object的工具。

和导出相关的工具：

1. package o2w：将Object转换成Workbook。
1. package w2f：将Workbook写到excel文件。
1. package m2f：将Message（一般是校验失败消息）追加到excel文件。

### spreadsheet-mapper-joda

在core的基础上，提供

1. package o2w：添加joda-time数据类型的支持。
1. package w2o：添加joda-time数据类型的支持。

### spreadsheet-mapper-java8

在core的技术上，提供对

1. package o2w：添加java 8 time数据类型的支持。
1. package w2o：添加java 8 time数据类型的支持。
1. package validation：添加java 8 lambda表达式的支持。

### 例子

1. 读取excel文件并生成Workbook的例子：[WorkbookReaderExample](examples/src/main/java/f2w/WorkbookReaderExample.java)
1. 校验workbook的例子：[WorkbookValidatorExample](examples/src/main/java/validation/WorkbookValidatorExample.java) [WorkbookValidatorWithValidatorBuilderExample](examples/src/main/java/validation/WorkbookValidatorWithValidatorBuilderExample.java)
1. 将Workbook转换成Object的例子：[Workbook2ObjectComposerExample](examples/src/main/java/w2o/Workbook2ObjectComposerExample.java)
1. 将Object转换成Workbook的例子：[Object2WorkbookExample](examples/src/main/java/o2w/Object2WorkbookExample.java)
1. 将Workbook写到excel文件的例子：[Workbook2ExcelWriterExample](examples/src/main/java/w2f/Workbook2ExcelWriterExample.java)
1. 将Message写到excel文件的例子：[Message2ExcelWriterExample](examples/src/main/java/m2f/Message2ExcelWriterExample.java)


## License

Released under the [Apache 2.0 license](license).
