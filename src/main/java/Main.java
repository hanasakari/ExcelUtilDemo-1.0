import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.*;

/**
 * Created by Hanasakari on 3/27/2018.
 */
public class Main {
    public static void main(String[] args) {
        List<Map<Integer, Map<Integer, Object>>> list = Main.excelToListMap(new File(Main.class.getResource("Untitled1.xlsx").getPath()));
        System.out.println(list);
    }
    public static List<Map<Integer, Map<Integer, Object>>> excelToListMap(File file) {
        // 将导入的excel转换为list<map>
        List<Map<Integer, Map<Integer, Object>>> rtn = new ArrayList<Map<Integer, Map<Integer, Object>>>();
        int Row = 0;
        int Col = 0;

        InputStream is = null;
        try {
            is = new FileInputStream(file);

            org.apache.poi.ss.usermodel.Workbook workbook = WorkbookFactory.create(is);

            Sheet sheet = workbook.getSheetAt(0);
            Iterator<org.apache.poi.ss.usermodel.Row> rows = sheet.rowIterator();

            Integer colNum = 0;
            while (rows.hasNext()) {
                Row row = rows.next();
                // 过滤表头
                if (row.getRowNum() == 0) {
                    continue;
                }
                int rowNum = row.getPhysicalNumberOfCells();
                Map<Integer, Map<Integer, Object>> rowMap = new HashMap<Integer, Map<Integer, Object>>();
                Map<Integer, Object> map = new HashMap<Integer, Object>();
                for (int i = 0; i < rowNum; i++) {
                    Cell cell = row.getCell(i);
                    Integer keyNum = cell.getColumnIndex();
                    String data = cell.toString();
                    Row = keyNum;
                    Col = colNum;
                    map.put(keyNum, data);
                }
                rowMap.put(colNum,map);
                rtn.add(rowMap);
                colNum +=1 ;
            }
            return rtn;
        } catch (Exception e) {
            e.printStackTrace();
            System.err.println(Col+"col"+Row+"row"+"is null");
        }
        return rtn;
    }
}
