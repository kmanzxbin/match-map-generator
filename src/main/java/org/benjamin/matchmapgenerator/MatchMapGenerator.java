package org.benjamin.matchmapgenerator;

import lombok.Data;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

@Data
public class MatchMapGenerator {

    File file;
    File scoringPaper;
    String category;
    /**
     * 分组数量
     */
    int groupMemberNumber;

    public static void main(String[] args) throws Exception {
        new MatchMapGenerator().makeSnake(4,4);
    }
    public static void start(String[] args) throws Exception {

        if (args.length < 4) {
            throw new IllegalArgumentException("missing args\n" +
                    "1st is name list file path\n" +
                    "2nd is score list output file path\n" +
                    "3rd is category\n" +
                    "4th is group member number\n");
        }

        MatchMapGenerator matchMapGenerator = new MatchMapGenerator();
        matchMapGenerator.file = new File(args[0]);
        matchMapGenerator.scoringPaper = new File(args[1]);
        matchMapGenerator.category = args[2];
        matchMapGenerator.groupMemberNumber = Integer.parseInt(args[3]);

//        matchMapGenerator.makeRoundOrder(Arrays.asList(new String[]{"1", "2", "3"}));
//        badminton.matchOrder(Arrays.asList(new String[]{"1", "2", "3", "4", "5", "6", "7"}));

        matchMapGenerator.generator();
    }

    public void generator(List<String> list) {
        try{
        List<List<String>> groupLists = getGroupList(list,
                list.size() / groupMemberNumber + (list.size() % groupMemberNumber == 0 ? 0 : 1));
        outputScoringPaper(groupLists, scoringPaper);
    } catch (Exception e) {
        e.printStackTrace();
    }
        System.out.println();
    }

    public void generator() {
        List<String> list = null;
        try {
            list = getList(category);
            if (list.size() <= 0) {
                System.out.println(String.format("no members found in file: %s", file));
                return;
            }
            if (list.size() < 3) {
                System.out.println(String.format("members num less than 3"));
                return;
            }

            if (list.size() / groupMemberNumber < 1) {
                System.out.println(String.format("groupNumber %s is not suitable, total: %s", groupMemberNumber, list.size()));
                return;
            }

            List<List<String>> groupLists = getGroupList(list,
                    list.size() / groupMemberNumber + (list.size() % groupMemberNumber == 0 ? 0 : 1));
            outputScoringPaper(groupLists, scoringPaper);
        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println();
    }

    /**
     * 得到参赛人名单
     *
     * @param category
     * @return
     * @throws IOException
     */
    public List<String> getList(String category) {
        List<String> lines = null;
        try {
            lines = FileUtils.readLines(file, "utf-8");
        } catch (IOException e) {
            e.printStackTrace();
            return lines;
        }
        lines = lines.stream().filter(StringUtils::isNotBlank)
                .map(String::trim)
                .collect(Collectors.toList());
        System.out.println("原始记录数: " + lines.size());
        List<Map.Entry<String, List<String>>> entrys = lines.stream().collect(Collectors.groupingBy(Function.identity())).entrySet().stream()
                .filter(t -> t.getValue().size() > 1).collect(Collectors.toList());

        if (entrys.size() > 0) {
            System.out.println("重复记录: ");
            entrys.stream().forEach(t -> System.out.println(t.getKey()));
            System.out.println("重复记录输出完毕。。。。。。");
            lines = lines.stream()
                    .distinct().collect(Collectors.toList());

            System.out.println("过滤输出开始");
            lines.stream().forEach(t -> System.out.println(t));
            System.out.println("过滤输出结束。。。。。。");
        }


        lines = lines.stream()
                .distinct()
                .filter(t -> t.contains(category))
                .collect(Collectors.toList());
        System.out.println(category + "有效记录数: " + lines.size());
        System.out.println(StringUtils.join(lines, "\n"));
        lines = lines.stream()
                .map(t -> {
                    t = t.split(" ")[0].trim();
//                    if (t.length() == 2) {
//                        t = t.charAt(0) + " " + t.charAt(1);
//                    }
                    return t;
                })
                .distinct()
                .collect(Collectors.toList());
        System.out.println("————————");
        System.out.println(category + "组参赛名单（" + lines.size() + "人）:");
        System.out.println(StringUtils.join(lines, "\n"));
        System.out.println("");
        return lines;

    }

    /**
     * 处理成蛇形分布
     *
     * @param list
     * @param groupNumber 分组数量
     * @return
     */
    List<List<String>> getGroupList(List<String> list, int groupNumber) {
        int width = list.size() / groupNumber + (list.size() % groupNumber == 0 ? 0 : 1);
        List<List<String>> groupLists = new ArrayList<>();
        int[][] snake = makeSnake(width, groupNumber);
        for (int i = 0; i <snake.length ; i++) {
            List<String> groupList = new ArrayList<>();
            groupLists.add(groupList);
            for (int j = 0; j <snake[i].length ; j++) {
                if (snake[i][j] < list.size()) {
                    groupList.add(list.get(snake[i][j]));
                }
            }
        }

//        for (int i = 0; i < groupNumber; i++) {
//            List<String> groupList = new ArrayList<>();
//            groupLists.add(groupList);
//            for (int j = i; j < list.size(); j += groupNumber) {
//                groupList.add(list.get(j));
//            }
//        }
        System.out.println(String.format("%s分组名单:", category));
        for (int i = 0; i < groupLists.size(); i++) {
            System.out.println(String.format("%s组: %s", getGroupId(i), groupLists.get(i)));
        }
        System.out.println();
        return groupLists;
    }

    int[][] makeSnake(int width, int high) {
        int[][] matrix = new int[high][];
        for (int i = 0; i < high; i++) {
            matrix[i] = new int[width];
        }
        // 0 下沉 1 上升
        boolean goDown = true;
        int index = 0;
        for (int i = 0; i < width; i++) {
            if (goDown) {
                for (int j = 0; j < high; j++) {
                    matrix[j][i] = index;
                    index++;
                }
            } else {
                for (int j = high - 1; j >= 0; j--) {
                    matrix[j][i] = index;
                    index++;
                }
            }
            goDown = !goDown;
        }
        for (int i = 0; i < matrix.length; i++) {
            ArrayList list = new ArrayList();
            for (int j = 0; j < matrix[i].length; j++) {
                list.add(matrix[i][j]);
            }
            System.out.println(list);
        }
        return matrix;
    }


    /**
     * 输出excel记分表
     *
     * @param groupLists
     * @throws Exception
     */
    void outputScoringPaper(List<List<String>> groupLists, File file) throws Exception {

        // 创建工作薄
        HSSFWorkbook workbook = new HSSFWorkbook();

        for (int i = 0; i < groupLists.size(); i++) {
            makeSheet(workbook, getGroupId(i) + "组", groupLists.get(i));
        }

        FileOutputStream xlsStream = new FileOutputStream(file);
        workbook.write(xlsStream);
    }

    String getGroupId(int i) {
        return (char) (65 + i) + "";
    }

    CellStyle getCellStyle(HSSFWorkbook workbook, HorizontalAlignment alignment) {
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellStyle.setBorderTop(BorderStyle.THIN);//上边框
        cellStyle.setBorderRight(BorderStyle.THIN);//右边框
        cellStyle.setAlignment(alignment);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return cellStyle;
    }

    void makeSheet(HSSFWorkbook workbook, String sheetName, List<String> list) {
        // 创建工作表
        HSSFSheet sheet = workbook.createSheet(sheetName);
        sheet.setDefaultColumnWidth(16);
        HSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 10);
        sheet.setDefaultRowHeight((short) (1.2 * 256));
        sheet.setVerticallyCenter(true);
        HSSFFont labelFont = workbook.createFont();
        labelFont.setFontHeightInPoints((short) 12);
        labelFont.setBold(true);

        CellStyle rightStyle = getCellStyle(workbook, HorizontalAlignment.RIGHT);
        CellStyle centerStyle = getCellStyle(workbook, HorizontalAlignment.CENTER);
        CellStyle leftStyle = getCellStyle(workbook, HorizontalAlignment.LEFT);

        int rowBase = 0;
        rowBase = addMergedCells(sheet, rowBase, list.size(), category + sheetName + "分组计分表", centerStyle, labelFont);

        // 构造矩阵
        List<String> cols = new ArrayList<>(list.subList(1, list.size()));
        Collections.reverse(cols);
        List<String> rowList = list.subList(0, list.size() - 1);
        // 名称横表头
        HSSFRow sheetRow = sheet.createRow(rowBase++);
        HSSFCell cell = sheetRow.createCell(0);
        cell.setCellStyle(centerStyle);
        setRichString("", font, cell);
        for (int i = 0; i < cols.size(); i++) {
            cell = sheetRow.createCell(i + 1);
            cell.setCellStyle(rightStyle);
            setRichString(cols.get(i) + " .", font, cell);
        }
        // 构造竖表和内容
        for (int i = 0; i < rowList.size(); i++) {
            sheetRow = sheet.createRow(rowBase++);
            cell = sheetRow.createCell(0);
            cell.setCellStyle(centerStyle);
            setRichString(rowList.get(i), font, cell);
            for (int col = 0; col < cols.size(); col++) {
                // 向工作表中添加数据
                cell = sheetRow.createCell(col + 1);
                cell.setCellStyle(centerStyle);
//                if (col == cols.size() - 1) {
//                    cell.setCellValue("——");
//                } else {
//                    cell.setCellValue(":");
//                }
                setRichString(":", font, cell);
            }
            cols.remove(cols.size() - 1);
        }

        rowBase++;

        // 输出每轮对阵表
        System.out.println(String.format("————%s对阵表————", category + sheetName));
        List<List<String[]>> roundOrders = makeRoundOrder(list);

        rowBase = addMergedCells(sheet, rowBase, roundOrders.stream().mapToInt(List::size).max().getAsInt() + 1
                , "对阵表", centerStyle, labelFont);
        for (int i = 0; i < roundOrders.size(); i++) {
            sheetRow = sheet.createRow(rowBase++);
            List<String[]> order = roundOrders.get(i);

            cell = sheetRow.createCell(0);
            cell.setCellStyle(centerStyle);
            setRichString(String.format("第%s轮", i + 1), font, cell);
            for (int j = 0; j < order.size(); j++) {
                cell = sheetRow.createCell(j + 1);
                cell.setCellStyle(centerStyle);
                setRichString(String.format("%s VS %s", order.get(j)[0], order.get(j)[1]), font, cell);
            }
        }

          rowBase ++;
        rowBase = addMergedCells(sheet, rowBase, list.size(), "积分表", centerStyle, labelFont);
        rowBase = addPointsTable(sheet, rowBase, list, leftStyle, font);

          rowBase ++;
        rowBase = addMergedCells(sheet, rowBase, list.size(), "排名表", centerStyle, labelFont);
        rowBase = addRankingTable(sheet, rowBase, list, leftStyle, font);
    }


    int addRankingTable( HSSFSheet sheet, int rowBase, List<String> list, CellStyle cellStyle, HSSFFont font) {
        HSSFRow row = sheet.createRow(rowBase++);
        for (int i = 0; i < list.size(); i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellStyle(cellStyle);
            setRichString(String.format("第%s名：", i+1), font, cell);
        }
        return rowBase;
    }

    int addPointsTable( HSSFSheet sheet, int rowBase, List<String> list, CellStyle cellStyle, HSSFFont font) {
        HSSFRow row = sheet.createRow(rowBase++);
        for (int i = 0; i < list.size(); i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellStyle(cellStyle);
            setRichString(list.get(i) + "：", font, cell);
        }
        return rowBase;
    }


    int addMergedCells( HSSFSheet sheet, int rowBase, int length, String value, CellStyle cellStyle, HSSFFont font) {
        // 不能在此处对rowBase++ 因为下面要合并这一行
        HSSFRow row = sheet.createRow(rowBase);
        for (int i = 0; i < length; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellStyle(cellStyle);
            setRichString(value, font, cell);
        }

        sheet.addMergedRegion(new CellRangeAddress(
                rowBase, //first row (0-based)
                rowBase, //last row (0-based)
                0, //first column (0-based)
                length - 1//last column (0-based)
        ));
        return ++rowBase;
    }

    void setRichString(String textValue, HSSFFont font, HSSFCell cell) {
        HSSFRichTextString richString = new HSSFRichTextString(textValue); //textValue是要设置大小的单元格数据
        richString.applyFont(font);
        cell.setCellValue(richString);
    }


    //  排序
    List<List<String[]>> makeRoundOrder(List<String> list) {
        list = new ArrayList<String>(list);
        if (list.size() % 2 != 0) {
            list.add("轮空");
        }
        List<String> others = list.subList(1, list.size());
//        Collections.reverse(others);
        List<List<String[]>> rounds = new ArrayList<>();
        for (int i = 0; i < others.size(); i++) {
            List<String[]> round = new ArrayList<>();
            rounds.add(round);

            round.add(new String[]{list.get(0), others.get(others.size() - 1)});
            for (int j = 0; j < (others.size() / 2); ) {
//                System.out.println(others.get(j++) + " VS " + others.get(others.size() - j - 1));
                round.add(new String[]{others.get(j++), others.get(others.size() - j - 1)});
            }
            Collections.rotate(others, 1);
        }


        for (int i = 0; i < rounds.size(); i++) {
            System.out.println(String.format("——第%s轮——", i + 1));
            List<String[]> order = rounds.get(i);
            for (int j = 0; j < order.size(); j++) {
                System.out.println(String.format("%s VS %s", order.get(j)[0], order.get(j)[1]));
            }
        }
        System.out.println();
        return rounds;
    }

}
