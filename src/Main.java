import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class Main {
  private static final String fileName = "C:\\Users\\cao.zm\\Desktop\\PurchaseRecord.xlsx";
  private static final String CMS_DATABASE_URL = "jdbc:mysql://60.205.142.85:3307/cms?useUnicode=true&autoReconnect=true&rewriteBatchedStatements=true&allowMultiQueries=true";
  private static final String USER_NAME = "root";
  private static final String PASSWORD = "abcd4321";
  public static void main(String[] args) {
    Statement statement;
    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
    Calendar today = Calendar.getInstance();
    Calendar yesterday = Calendar.getInstance();
    yesterday.add(Calendar.DATE, -1);
    String sql_1 = "SELECT fca.USER_NAME AS '商户名称', prr.buyer AS '用户ID', "
            + "ffp.FONT_NAME AS '字体名称', fo.FINAL_FEE AS '购买价格/分', "
            + "prr.create_date AS '购买日期' "
            + "FROM pay_result_record prr "
            + "LEFT JOIN fs_order fo ON fo.ORDER_ID = prr.ORDERID "
            + "LEFT JOIN fs_customer_app fca ON fca.APP_KEY = fo.APP_KEY "
            + "LEFT JOIN fs_order_item foi ON foi.ORDER_ID = prr.ORDERID "
            + "LEFT JOIN fs_font_pool ffp ON ffp.id = foi.ITEM_ID "
            + "WHERE prr.create_date >= '" + sdf.format(yesterday.getTime()) + "' "
            + "AND prr.create_date < '" + sdf.format(today.getTime()) + "' "
            + "ORDER BY prr.CREATE_DATE DESC";
    System.out.println(sql_1);
    List<PurchaseRecord> list = new ArrayList<>();
    try {
      Connection connection = DriverManager.getConnection(CMS_DATABASE_URL, USER_NAME, PASSWORD);
      statement = connection.createStatement();
      ResultSet resultSet = statement.executeQuery(sql_1);
      while (resultSet.next()) {
        PurchaseRecord purchaseRecord = new PurchaseRecord();
        purchaseRecord.setAppName(resultSet.getString("商户名称"));
        purchaseRecord.setUserId(resultSet.getString("用户ID"));
        purchaseRecord.setFontName(resultSet.getString("字体名称"));
        purchaseRecord.setPrice(resultSet.getInt("购买价格/分"));
        purchaseRecord.setDate(resultSet.getDate("购买日期"));
        list.add(purchaseRecord);
      }
      FileInputStream is = new FileInputStream(fileName);
      XSSFWorkbook workbook = new XSSFWorkbook(is);
      FileOutputStream fileOutputStream = new FileOutputStream(fileName);
      XSSFSheet sheet;
      XSSFRow row;
      XSSFCell cell;
      sheet = workbook.getSheetAt(0);
      sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 1));
      row = sheet.getRow(0);
      cell = row.createCell(0);
      cell.setCellValue(sdf.format(yesterday.getTime()));
      workbook.write(fileOutputStream);
      workbook.close();
    } catch (SQLException e) {
      e.printStackTrace();
    } catch (IOException e) {
      e.printStackTrace();
    }
  }
}

class PurchaseRecord {
  private String appName;
  private String userId;
  private String fontName;
  private Integer price;
  private Date date;

  public String getAppName() {
    return appName;
  }

  void setAppName(String appName) {
    this.appName = appName;
  }

  public String getUserId() {
    return userId;
  }

  void setUserId(String userId) {
    this.userId = userId;
  }

  public String getFontName() {
    return fontName;
  }

  void setFontName(String fontName) {
    this.fontName = fontName;
  }

  public Integer getPrice() {
    return price;
  }

  void setPrice(Integer price) {
    this.price = price;
  }

  public Date getDate() {
    return date;
  }

  void setDate(Date date) {
    this.date = date;
  }
}