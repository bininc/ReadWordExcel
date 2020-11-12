import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.impl.CTPImpl;

import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.io.*;
import java.util.Iterator;
import java.util.List;

public class GenerateMysql {

    public static void main(String[] args) throws IOException {
        do {
            String filePath, tableName;
            int tableIndex = -1;
            if (args.length < 3) {
                System.out.println("请输入Word文件路径:");
//            filePath = new BufferedReader(new InputStreamReader(System.in)).readLine();
                filePath = "/Volumes/Data/Document/卫监/全民健康信息平台建设及投资运营项目平台共享数据集V1.5.docx";
                //System.out.println("请输入Word文件中表格索引");
                //tableIndex = Integer.parseInt(new BufferedReader(new InputStreamReader(System.in)).readLine());
                System.out.println("请输入表名");
                tableName = new BufferedReader(new InputStreamReader(System.in)).readLine();
            } else {
                filePath = args[0];
                tableIndex = Integer.parseInt(args[1]);
                tableName = args[2];
            }
            testWord(filePath, tableIndex, tableName);
            System.out.println("温馨提示：按Q退出，任意键继续");
        } while (!new BufferedReader(new InputStreamReader(System.in)).readLine().equalsIgnoreCase("q"));
    }

    public static void testWord(String filePath, int tableIndex, String tableName) {
        try {
            FileInputStream in = new FileInputStream(filePath);//载入文档 //如果是office2007  docx格式
            if (filePath.toLowerCase().endsWith("docx")) {
                //word 2007 图片不会被读取， 表格中的数据会被放在字符串的最后
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
                StringBuilder sb = new StringBuilder();
                sb.append(String.format("CREATE TABLE %s (%n", tableName));
                //读取每一行数据
                for (int i = 1; i < rows.size(); i++) {
                    XWPFTableRow row = rows.get(i);
                    //读取每一列数据
                    List<XWPFTableCell> cells = row.getTableCells();
                    String colName = cells.get(1).getText().trim();
                    if (!StringUtils.isEmpty(colName) && !StringUtils.isBlank(colName) && !StringUtils.isWhitespace(colName)) {
                        String comment = cells.get(2).getText();
                        String colType = cells.get(3).getText().replace("（", "(").replace("）", ")").trim();
                        if (colType.equalsIgnoreCase("date"))
                            colType = "datetime";
                        if (StringUtils.startsWithIgnoreCase(colType, "VARCHAR2"))
                            colType = StringUtils.replaceIgnoreCase(colType, "VARCHAR2", "VARCHAR");
                        if (StringUtils.startsWithIgnoreCase(colType, "NUMBER"))
                            colType = StringUtils.replaceIgnoreCase(colType, "NUMBER", "NUMERIC");
                        boolean notNull = cells.get(4).getText().equals("必填");
                        sb.append(String.format("`%s` %s %s NULL COMMENT '%s',%n", colName, colType, notNull ? "NOT" : "", comment));
                    }
                }
                String sql = sb.substring(0, sb.length() - 2) + ");";
                System.out.println();
                setIntoClipboard(sql);
                System.out.println(sql);
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
