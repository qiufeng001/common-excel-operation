package validation;

import com.supwisdom.spreadsheet.mapper.f2w.WorkbookReader;
import com.supwisdom.spreadsheet.mapper.f2w.excel.Excel2WorkbookReader;
import com.supwisdom.spreadsheet.mapper.m2f.MessageWriter;
import com.supwisdom.spreadsheet.mapper.m2f.excel.ExcelMessageWriter;
import com.supwisdom.spreadsheet.mapper.model.core.Workbook;
import com.supwisdom.spreadsheet.mapper.model.meta.FieldMetaBean;
import com.supwisdom.spreadsheet.mapper.model.meta.SheetMetaBean;
import com.supwisdom.spreadsheet.mapper.model.meta.WorkbookMetaBean;
import com.supwisdom.spreadsheet.mapper.model.msg.Message;
import com.supwisdom.spreadsheet.mapper.validation.DefaultSheetValidationJob;
import com.supwisdom.spreadsheet.mapper.validation.DefaultWorkbookValidationJob;
import com.supwisdom.spreadsheet.mapper.validation.validator.cell.RequireValidator;
import com.supwisdom.spreadsheet.mapper.validation.validator.cell.UniqueValidator;
import com.supwisdom.spreadsheet.mapper.validation.validator.cell.ValueScopeValidator;
import com.supwisdom.spreadsheet.mapper.validation.validator.unioncell.LambdaUnionCellValidator;
import com.supwisdom.spreadsheet.mapper.validation.validator.workbook.SheetAmountValidator;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * 校验excel的例子。
 * 这个例子里面使用了{@link Excel2WorkbookReader}，{@link ExcelMessageWriter}
 * Created by qianjia on 2017/3/10.
 */
public class WorkbookValidatorExample {

  private static final Logger LOGGER = LoggerFactory.getLogger(WorkbookValidatorExample.class);

  public static void main(String[] args) throws IOException {

    // 先从excel文件中读取成 Workbook
    WorkbookReader workbookReader = new Excel2WorkbookReader();
    Workbook workbook;
    try (InputStream inputStream = WorkbookValidatorExample.class.getResourceAsStream("test.xlsx")) {
      workbook = workbookReader.read(inputStream);
    }

    // 构造Workbook元信息，包含信息：有几个sheet，每个sheet里有包含有哪些field，数据是从哪行开始的
    WorkbookMetaBean workbookMeta = new WorkbookMetaBean();

    // 第一个Sheet的元信息，这里定义了真正的数据是从第4行开始的
    // 这样在校验的时候会跳过前面3行，因为前面3行是表头
    SheetMetaBean sheetMeta = new SheetMetaBean(4);

    // 添加Field元信息，field的名字可以和excel中的名字不一样
    // 不过为了便于理解，一般来说都是一样的
    // field的名字在后面注册校验器的时候有用
    sheetMeta.addFieldMeta(new FieldMetaBean("code", 1));
    sheetMeta.addFieldMeta(new FieldMetaBean("name", 2));
    sheetMeta.addFieldMeta(new FieldMetaBean("gender", 3));
    sheetMeta.addFieldMeta(new FieldMetaBean("xueli", 4));
    sheetMeta.addFieldMeta(new FieldMetaBean("identityCardType", 5));
    sheetMeta.addFieldMeta(new FieldMetaBean("identityCardNo", 6));
    sheetMeta.addFieldMeta(new FieldMetaBean("grade", 7));
    sheetMeta.addFieldMeta(new FieldMetaBean("college", 8));
    sheetMeta.addFieldMeta(new FieldMetaBean("major", 9));
    sheetMeta.addFieldMeta(new FieldMetaBean("adminclass", 10));

    workbookMeta.addSheetMeta(sheetMeta);

    // 构建Workbook校验工作
    DefaultWorkbookValidationJob workbookValidationJob = new DefaultWorkbookValidationJob();
    // 校验workbook的sheet数量等于1
    workbookValidationJob.addValidator(new SheetAmountValidator(1));

    DefaultSheetValidationJob sheetValidationJob = new DefaultSheetValidationJob();
    workbookValidationJob.addSheetValidationJob(sheetValidationJob);

    sheetValidationJob.addValidator(new RequireValidator().matchField("code").group("code").errorMessage("必填"));
    sheetValidationJob.addValidator(new RequireValidator().matchField("gender").group("gender").errorMessage("必填"));
    sheetValidationJob.addValidator(new RequireValidator().matchField("xueli").group("xueli").errorMessage("必填"));
    sheetValidationJob.addValidator(
        new RequireValidator().matchField("identityCardType").group("identityCardType").errorMessage("必填"));
    sheetValidationJob
        .addValidator(new RequireValidator().matchField("identityCardNo").group("identityCardNo").errorMessage("必填"));
    sheetValidationJob.addValidator(new RequireValidator().matchField("grade").group("grade").errorMessage("必填"));
    sheetValidationJob.addValidator(new RequireValidator().matchField("college").group("college").errorMessage("必填"));
    sheetValidationJob.addValidator(new RequireValidator().matchField("major").group("major").errorMessage("必填"));
    sheetValidationJob
        .addValidator(new RequireValidator().matchField("adminclass").group("adminclass").errorMessage("必填"));

    sheetValidationJob.addValidator(
        new ValueScopeValidator(new String[] { "男", "女" }).matchField("gender").group("gender")
            .errorMessage("只能填写：男、女"));
    sheetValidationJob.addValidator(
        new ValueScopeValidator(new String[] { "是", "否" }).matchField("xueli").group("xueli").errorMessage("只能填写：是、否"));
    sheetValidationJob.addValidator(
        new ValueScopeValidator(new String[] { "居民身份证", "护照", "港澳通行证" }).matchField("identityCardType")
            .group("identityCardType").errorMessage("只能填写。居民身份证、护照、港澳通行证"));
    sheetValidationJob.addValidator(
        new LambdaUnionCellValidator((cells, metas) -> {
          String idcardType = cells.get(0).getValue();
          String idcardNo = cells.get(1).getValue();
          if (idcardType.equals("居民身份证")) {
            return idcardNo.length() == 18;
          }
          return true;
        })
            .group("identity")
            .errorMessage("证件号码长度不符合")
            .matchFields("identityCardType", "identityCardNo")
            .dependsOn("identityCardType")
    );

    sheetValidationJob.addValidator(new UniqueValidator().matchField("code").group("code").errorMessage("不唯一"));
    sheetValidationJob
        .addValidator(new UniqueValidator().matchField("identityCardNo").group("identityCardNo").errorMessage("不唯一"));

    // 开始校验
    boolean valid = workbookValidationJob.validate(workbook, workbookMeta);
    List<Message> workbookErrors = workbookValidationJob.getErrorMessages();

    MessageWriter messageWriter = null;
    try (InputStream resourceAsStream = WorkbookValidatorExample.class.getResourceAsStream("test.xlsx")) {
      // 读取excel内容
      messageWriter = new ExcelMessageWriter(resourceAsStream);

      File tempFile = File.createTempFile("test-valid-result", ".xlsx");

      try (FileOutputStream outputStream = new FileOutputStream(tempFile)) {
        messageWriter.write(workbookErrors, outputStream);
      }
      LOGGER.info("test.xlsx validation result: " + valid + ". Output file: " + tempFile.getAbsolutePath());

    }

  }

}
