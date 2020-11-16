import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.*;

import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.io.*;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;

public class GenerateCSharp {

    public static void main(String[] args) throws IOException {
        Scanner sc = new Scanner(System.in);
        String filePath = null;
        File file;
        do {
            System.out.println("请输入Word文件路径:");
            file = new File(sc.nextLine());
            if (!file.exists() || !file.isFile())
                System.out.println("文件路径无效！");
            else
                filePath = file.getPath();
        } while (filePath == null);
        do {
            String tableName;
            System.out.println("当前Word文件路径：" + filePath);
            System.out.println("请输入表名");
            tableName = sc.nextLine();
            testWord(filePath, tableName);
            System.out.println("温馨提示：输入Q退出，回车键继续");
        } while (!sc.nextLine().equalsIgnoreCase("q"));
    }

    public static void testWord(String filePath, String tableName) {
        try {
            if (StringUtils.isBlank(tableName)) {
                System.out.println("表名不能为空！");
                return;
            }
            if (filePath.toLowerCase().endsWith("docx")) {
                //word 2007 图片不会被读取， 表格中的数据会被放在字符串的最后
                FileInputStream in = new FileInputStream(filePath);//载入文档 //如果是office2007  docx格式
                XWPFDocument xwpf = new XWPFDocument(in);//得到word文档的信息
                List<XWPFParagraph> listParagraphs = xwpf.getParagraphs();//得到段落信息
                Iterator<XWPFParagraph> it = listParagraphs.iterator();
                int pPos = -1;
                while (it.hasNext()) {
                    XWPFParagraph p = it.next();
                    if (p.getText().contains(tableName)) {
                        pPos = xwpf.getPosOfParagraph(p);
                        break;
                    }
                }

                int tPos = -1;
                if (pPos >= 0) {
                    tPos = xwpf.getTablePos(pPos + 1);
                    if (tPos < 0)
                        tPos = xwpf.getTablePos(pPos + 2);
                }

                if (tPos < 0) {
                    System.out.println("未找到对应的表格");
                    return;
                }

                XWPFTable table = xwpf.getTableArray(tPos);
                List<XWPFTableRow> rows = table.getRows();
                int int32MaxLen = String.format("%d", Integer.MAX_VALUE).length();
                StringBuilder sb = new StringBuilder();
                sb.append(String.format("using System;%nnamespace Neptune.HealthBureau.WuXi.Data.Poco%n{%npublic class %s : UpdateRecordBase%n{%n", tableName));
                //读取每一行数据
                for (int i = 1; i < rows.size(); i++) {
                    XWPFTableRow row = rows.get(i);
                    //读取每一列数据
                    List<XWPFTableCell> cells = row.getTableCells();
                    String colName = cells.get(1).getText().trim();
                    if (!StringUtils.isEmpty(colName) && !StringUtils.isBlank(colName) && !StringUtils.isWhitespace(colName)) {
                        String comment = cells.get(2).getText();
                        String commentEx = cells.get(5).getText();
                        String commentEx2 = cells.get(6).getText();
                        boolean notNull = cells.get(4).getText().equals("必填");
                        String colType = cells.get(3).getText().replace("（", "(").replace("）", ")").trim();
                        if (colType.equalsIgnoreCase("date"))
                            colType = "DateTime" + (notNull ? "" : "?");
                        else if (StringUtils.startsWithIgnoreCase(colType, "VARCHAR"))
                            colType = "String";
                        else if (StringUtils.startsWithIgnoreCase(colType, "NUMBER") || StringUtils.startsWithIgnoreCase(colType, "numeric")) {
                            String numberLen = StringUtils.substringBetween(colType, "(", ")");
                            if (numberLen == null)
                                colType = "int";
                            else {
                                numberLen = numberLen.replace('，', ',');
                                String[] numberLenArr = numberLen.split(",");
                                if (numberLenArr.length > 0) {
                                    //整形
                                    if (numberLenArr.length == 1) {
                                        long numLen = Long.parseLong(numberLenArr[0]);
                                        if (numLen < int32MaxLen)
                                            colType = "int";
                                        else
                                            colType = "long";
                                    } else
                                        colType = "decimal";
                                }
                            }
                            colType = colType + (notNull ? "" : "?");
                        } else if (colType.equalsIgnoreCase("blob")) {
                            colType = "Byte[]";
                        } else if (colType.equalsIgnoreCase("INTEGER")) {
                            colType = "int";
                        }

                        sb.append(String.format("/// <summary>%n/// %s%n", comment));
                        if (commentEx != null && StringUtils.isNotEmpty(commentEx) && StringUtils.isNotBlank(commentEx))
                            sb.append(String.format("/// %s%n", commentEx));
                        if (commentEx2 != null && StringUtils.isNotEmpty(commentEx2) && StringUtils.isNotBlank(commentEx2))
                            sb.append(String.format("/// %s%n", commentEx2));
                        sb.append(String.format("/// </summary>%n"));
                        sb.append(String.format("public %s %s {get; set;}%n", colType, colName));
                    }
                }
                sb.append(String.format("}%n}"));
                System.out.println();
                setIntoClipboard(sb.toString());
                System.out.println(sb);
                System.out.println();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void setIntoClipboard(String data) {
        Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
        clipboard.setContents(new StringSelection(data), null);
    }
}
