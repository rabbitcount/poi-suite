//package com.ocelot.poi;
//
//import com.monitorjbl.xlsx.StreamingReader;
//import com.monitorjbl.xlsx.impl.StreamingCell;
//import com.ocelot.poi.annotation.CellInfo;
//import com.ocelot.poi.constant.PropertyTypeEnum;
//import com.ocelot.poi.excp.PoiSuiteException;
//import lombok.AllArgsConstructor;
//import lombok.Data;
//import lombok.NoArgsConstructor;
//import lombok.extern.slf4j.Slf4j;
//import org.apache.commons.collections4.CollectionUtils;
//import org.apache.commons.collections4.MapUtils;
//import org.apache.commons.lang3.StringUtils;
//import org.apache.commons.lang3.reflect.MethodUtils;
//import org.apache.poi.ss.usermodel.CellType;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.Workbook;
//
//import java.io.ByteArrayInputStream;
//import java.io.IOException;
//import java.io.InputStream;
//import java.lang.reflect.Field;
//import java.lang.reflect.InvocationTargetException;
//import java.lang.reflect.ParameterizedType;
//import java.lang.reflect.Type;
//import java.math.BigDecimal;
//import java.text.ParseException;
//import java.text.SimpleDateFormat;
//import java.util.*;
//import java.util.regex.Matcher;
//import java.util.regex.Pattern;
//import java.util.stream.Collectors;
//
///**
// * @description:
// * @author: 盗梦的虎猫
// * @date: 2018/6/28
// * @time: 6:46 PM
// * Copyright (C) 2018 AnOcelot
// * All rights reserved
// */
//@Slf4j
//public class CustomExcelUtil {
//
//    public static final String DEFAULT_MONTH_PATTERN = "yyyy-MM";
//
//    /**
//     * @param contentBytes
//     * @return
//     */
//    public static Workbook buildStreamWorkbook(byte[] contentBytes) throws IOException {
//        InputStream stream = new ByteArrayInputStream(contentBytes);
//        return StreamingReader.builder()
//                .rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
//                .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
//                .open(stream);
//    }
//
//    /**
//     * 读取excel sheet中的数据（不抛出异常）
//     *
//     * @param sheet
//     * @param clazz
//     * @param <T>
//     * @return
//     */
//    public static <T> List<T> readSheetSafe(Sheet sheet, Class<T> clazz) {
//        try {
//            return readSheetStream(sheet, clazz);
//        } catch (PoiSuiteException e) {
//            throw e;
//        } catch (Exception e) {
//            e.printStackTrace();
//            return null;
//        }
//    }
//
//    /**
//     * 读取excel sheet中的数据
//     *
//     * @param sheet
//     * @param clazz
//     * @param <T>
//     * @return
//     * @throws InvocationTargetException
//     * @throws NoSuchMethodException
//     * @throws InstantiationException
//     * @throws IllegalAccessException
//     */
//    private static <T> List<T> readSheetStream(Sheet sheet, Class<T> clazz)
//            throws InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException {
//        Map<String, Integer> headers = new HashMap<>();
//        int rowIndex = 1;
//        List<T> result = new LinkedList<>();
//        ExcelEntity excelEntity = null;
//        for (Row row : sheet) {
//            if (row == null) {
//                continue;
//            }
//            // 调整 仅读取文件头
//            if (MapUtils.isEmpty(headers)) {
//                // 仅读取一次头文件
//                int j = 0;
//                for (org.apache.poi.ss.usermodel.Cell cell : row) {
//                    String cellValue = getCellValue(cell);
//                    if (StringUtils.isNotEmpty(cellValue)) {
//                        headers.put(cellValue, j);
//                    }
//                    ++j;
//                }
//                excelEntity = convertPropertyExcelIndex(headers, clazz);
//                if (CollectionUtils.isNotEmpty(excelEntity.getMissColumns())) {
//                    String missMessage = excelEntity.getMissColumns().stream().collect(Collectors.joining(","));
//                    if (StringUtils.isNotBlank(missMessage)) {
//                        throw new PoiSuiteException();
//                    }
//                }
//            } else {
//                T entity = readRow(row, rowIndex, clazz, excelEntity);
//                if (Objects.nonNull(entity)) {
//                    result.add(entity);
//                }
//            }
//            ++rowIndex;
//        }
////        ExcelEntity e = convertPropertyExcelIndex(headers, clazz);
////        if (CollectionUtils.isNotEmpty(e.getMissColumns())) {
////            String missMessage = e.getMissColumns().stream().collect(Collectors.joining(","));
////            CheckUtils.check(StringUtils.isBlank(missMessage), WelfareException.class,
////                    WelfareError.UPLOAD_FILE_MISS_COLUMNS, "缺少列" + missMessage);
////        }
//        // 读取数据
////        for (int i = ++index; i <= sheet.getLastRowNum(); ++i) {
////            T entity = readRow(sheet, i, clazz, e);
////            if (Objects.nonNull(entity)) {
////                result.add(entity);
////            }
////        }
//        return result;
//    }
//
//    /**
//     * 读取行数据
//     *
//     * @param row
//     * @param rowIndex
//     * @param clazz
//     * @param excelEntity
//     * @param <T>
//     * @return
//     * @throws IllegalAccessException
//     * @throws InstantiationException
//     * @throws InvocationTargetException
//     * @throws NoSuchMethodException
//     */
//    private static <T> T readRow(Row row, int rowIndex, Class<T> clazz, ExcelEntity excelEntity)
//            throws IllegalAccessException, InstantiationException, InvocationTargetException, NoSuchMethodException {
//        if (Objects.isNull(row)) {
//            return null;
//        }
//        T entity = clazz.newInstance();
//        if (rowIndex % 1000 == 0) {
//            log.info("haveReadExcel({})", rowIndex);
//        }
//        // 异常提示中，需要此字段
//        PropertyUtils.setProperty(entity, "rowNumber", rowIndex);
//        // 基于大多数情况下，字段应该是完整的
//        boolean isEmptyLine = true;     // 标记全行为空（用于跳过空白行）
//        List<Cell> emptyCell = Lists.newArrayList();
//        List<String> formatErrorCell = null;
//        for (Cell cell : excelEntity.getProperties()) {
//            try {
//                Object cellValue = getFormatCellValue(row, cell);
//                if (cell.isPrimary() && Objects.isNull(cellValue)) {
//                    // 不读取指定的行
//                    return null;
//                }
//                if (Objects.isNull(cellValue)) {
//                    if (cell.isNotNull()) {
//                        // 非空 需要记录异常
//                        emptyCell.add(cell);
//                    }
//                } else {
//                    isEmptyLine &= false;
//                    PropertyUtils.setProperty(entity, cell.getPropertyName(), cellValue);
//                }
//            } catch (OverLengthException e) {
//                // 格式异常
//                if (Objects.isNull(formatErrorCell)) {
//                    formatErrorCell = Lists.newArrayList();
//                }
//                formatErrorCell.add(String.format(FORMAT_OVER_LENGTH_COLUMN_TEMPLATE, cell.getFullColumnName(), cell.length,
//                        e.getCurrent()));
//            } catch (Exception e) {
//                // 格式异常
//                if (Objects.isNull(formatErrorCell)) {
//                    formatErrorCell = Lists.newArrayList();
//                }
//                formatErrorCell.add(String.format(FORMAT_ERROR_COLUMN_TEMPLATE, cell.getFullColumnName()));
//            }
//        }
//        for (CellGroup group : excelEntity.getGroups()) {
//            // propertyName -> objectValue
//            List<Cell> emptyGroupCell = null;
//            Map<String, Object> properties = new HashMap<>();
//            for (Cell cell : group.getProperties()) {
//                // 转换为适当的类型
//                try {
//                    Object cellValue = getFormatCellValue(row, cell);
//                    if (Objects.isNull(cellValue)) {
//                        if (cell.isNotNull()) {
//                            if (Objects.isNull(emptyGroupCell)) {
//                                emptyGroupCell = Lists.newArrayList();
//                            }
//                            emptyGroupCell.add(cell);
//                        }
//                    } else {
//                        properties.put(cell.getPropertyName(), cellValue);
//                    }
//                } catch (OverLengthException e) {
//                    // 格式异常
//                    if (Objects.isNull(formatErrorCell)) {
//                        formatErrorCell = Lists.newArrayList();
//                    }
//                    formatErrorCell.add(String.format(FORMAT_OVER_LENGTH_COLUMN_TEMPLATE, cell.getFullColumnName(), cell.length,
//                            e.getCurrent()));
//                } catch (Exception e) {
//                    // 格式异常
//                    if (Objects.isNull(formatErrorCell)) {
//                        formatErrorCell = Lists.newArrayList();
//                    }
//                    formatErrorCell.add(String.format(FORMAT_ERROR_COLUMN_TEMPLATE, cell.getFullColumnName()));
//                }
//            }
//            // 如果没有有效的数据，直接跳过
//            if (MapUtils.isEmpty(properties)) {
//                continue;
//            }
//            if (CollectionUtils.isNotEmpty(emptyGroupCell)) {
//                // 分组中存在空数据，需要单独标记
//                emptyCell.addAll(emptyGroupCell);
//                continue;
//            }
//            isEmptyLine &= false;
//            Object newOne = group.getClazz().newInstance();
//            PropertyUtils.setProperty(newOne, "type", group.getType());
//            PropertyUtils.setProperty(newOne, "groupName", group.getGroupName());
//            for (Map.Entry<String, Object> e : properties.entrySet()) {
//                PropertyUtils.setProperty(newOne, e.getKey(), e.getValue());
//            }
//            MethodUtils.invokeMethod(entity, "addDetail", newOne);
//        }
//        if (isEmptyLine) {
//            // 如果是空白行，直接跳过
//            return null;
//        }
//        boolean hasEmptyCell = CollectionUtils.isNotEmpty(emptyCell);
//        boolean hasFormatErrorCell = CollectionUtils.isNotEmpty(formatErrorCell);
//        if (hasEmptyCell || hasFormatErrorCell) {
//            List<String> errorMessage = Lists.newArrayList();
//            if (hasEmptyCell) {
//                String emptyErrorMessage = emptyCell.stream()
//                        .map(t -> String.format(EMPTY_COLUMN_TEMPLATE, t.getFullColumnName()))
//                        .collect(Collectors.joining(","));
//                if (StringUtils.isNotBlank(emptyErrorMessage)) {
//                    errorMessage.add(emptyErrorMessage);
//                }
//            }
//            if (hasFormatErrorCell) {
//                String formatErrorMessage = formatErrorCell.stream()
//                        .collect(Collectors.joining(","));
//                if (StringUtils.isNotBlank(formatErrorMessage)) {
//                    errorMessage.add(formatErrorMessage);
//                }
//            }
//            String finalErrorMessage = errorMessage.stream().collect(Collectors.joining(","));
//            PropertyUtils.setProperty(entity, "errorMessage", finalErrorMessage);
//            return entity;
//        }
//        return entity;
//    }
//
//    private static final String EMPTY_COLUMN_TEMPLATE = "%s不可为空";
//    private static final String FORMAT_ERROR_COLUMN_TEMPLATE = "%s数据格式错误";
//    private static final String FORMAT_OVER_LENGTH_COLUMN_TEMPLATE = "%s数据最大长度%d实际%d";
//
//    /**
//     * -> index
//     *
//     * @param headers
//     * @param clazz
//     * @param <T>
//     * @return
//     */
//    public static <T> ExcelEntity convertPropertyExcelIndex(Map<String, Integer> headers,
//                                                            Class<T> clazz) {
//        List<String> missColumns = Lists.newArrayList();
//        // 嵌套List字段
//        List<ListCell> listProperties = new LinkedList<>();
//        // columnName -> propertyName
//        Map<String, Cell> cellPropertyMap = new HashMap<>();
//        // CellInfo与CellListInfo排他出现
//        for (Field field : clazz.getDeclaredFields()) {
//            // 转换cellInfo
//            CellInfo cellInfo = field.getAnnotation(CellInfo.class);
//            if (Objects.nonNull(cellInfo)) {
//                if (!headers.containsKey(cellInfo.value())) {
//                    missColumns.add(cellInfo.value());
//                    continue;
//                } else {
//                    cellPropertyMap.put(cellInfo.value(), Cell.convert(field.getName(), field, cellInfo));
//                    continue;
//                }
//            }
//            // 转换CellListInfo
//            CellListInfo cellListInfo = field.getAnnotation(CellListInfo.class);
//            if (Objects.isNull(cellListInfo)) {
//                continue;
//            }
//            if (field.getType() == java.util.List.class) {
//                Type genericType = field.getGenericType();
//                if (Objects.isNull(genericType)) {
//                    continue;
//                }
//                // 如果是泛型参数的类型
//                if (genericType instanceof ParameterizedType) {
//                    ParameterizedType pt = (ParameterizedType) genericType;
//                    //得到泛型里的class类型对象
//                    Class<?> genericClazz = (Class<?>) pt.getActualTypeArguments()[0];
//                    listProperties.add(new ListCell(genericClazz, cellListInfo.regex(), cellListInfo.groupName(),
//                            cellListInfo.group(), cellListInfo.columnNameGroup()));
//                }
//            }
//        }
//        List<Cell> properties = new LinkedList<>();
//        // type(06) -> CellGroup
//        Map<String, CellGroup> groupMap = new HashMap<>();
//        // 遍历头文件
//        for (Map.Entry<String, Integer> e : headers.entrySet()) {
//            String fullColumnName = e.getKey();
//            // cell直接匹配的
//            Cell cell = cellPropertyMap.get(e.getKey());
//            if (Objects.nonNull(cell)) {
//                int excelIndex = e.getValue();
//                cell.setExcelIndex(excelIndex);
//                cell.setFullColumnName(fullColumnName);
//                properties.add(cell);
//                continue;
//            }
//            // 正则及嵌套的
//            for (ListCell listCell : listProperties) {
//                Pattern pattern = Pattern.compile(listCell.getRegex());
//                Matcher matcher = pattern.matcher(fullColumnName);
//                if (!matcher.matches()) {
//                    continue;
//                }
//                // (06)_XXX_(XX) 列聚合名称
//                String groupName = matcher.group(listCell.getGroupName());
//                String type = matcher.group(listCell.getGroup());
//                String columnName = matcher.group(listCell.getColumnNameGroup());
//                // columnName是否与预期匹配
//                Cell subCell = listCell.getProperties().get(columnName);
//                if (Objects.nonNull(subCell)) {
//                    int currentColumnIndex = e.getValue();
//                    CellGroup cellGroup = groupMap.get(type);
//                    if (Objects.isNull(cellGroup)) {
//                        cellGroup = new CellGroup(type, listCell.getClazz());
//                        cellGroup.setGroupName(groupName);
//                        groupMap.put(type, cellGroup);
//                    }
//                    subCell.setFullColumnName(fullColumnName);
//                    cellGroup.addCell(subCell, currentColumnIndex);
//                }
//            }
//        }
//        List<CellGroup> groups = Lists.newArrayList(groupMap.values());
//        return new ExcelEntity(properties, groups, missColumns);
//    }
//
//    private static class OverLengthException extends RuntimeException {
//
//        private int current;
//
//        public int getCurrent() {
//            return current;
//        }
//
//        private OverLengthException(int current) {
//            this.current = current;
//        }
//    }
//
//    /**
//     * 获取格式化后的字段
//     *
//     * @param row
//     * @param cell
//     * @return
//     */
//    private static Object getFormatCellValue(Row row, Cell cell) {
//        org.apache.poi.ss.usermodel.Cell currentCell = row.getCell(cell.getExcelIndex());
//        String cellValue = getCellValue(row.getCell(cell.getExcelIndex()));
//        if (StringUtils.isBlank(cellValue)) {
//            return null;
//        }
//        int length = StringUtils.length(cellValue);
//        if (length > cell.getLength()) {
//            // 超过限定长度
//            throw new OverLengthException(length);
//        }
//        switch (cell.getPropertyType()) {
//            case DATE:
//                String cellFormat = cell.getFormat();
//                if (currentCell.getCellTypeEnum() == CellType.STRING) {
//                    cellValue = cellValue.replace("/", "-");
//                }
//                SimpleDateFormat formatter = new SimpleDateFormat(cellFormat);
//                try {
//                    return formatter.parse(cellValue);
//                } catch (ParseException e) {
//                    e.printStackTrace();
//                    throw new RuntimeException();
//                }
//            case BIG_DECIMAL:
//                return new BigDecimal(cellValue).setScale(2, BigDecimal.ROUND_HALF_UP);
//            default:
//                return cellValue;
//        }
//    }
//
//    @Data
//    static class ListCell {
//        ListCell(Class<?> clazz, String regex, int groupName, int group, int columnNameGroup) {
//            this.clazz = clazz;
//            this.regex = regex;
//            this.groupName = groupName;
//            this.group = group;
//            this.columnNameGroup = columnNameGroup;
//            this.properties =
//                    Arrays.stream(clazz.getDeclaredFields())
//                            .filter(t -> t.isAnnotationPresent(CellInfo.class))
//                            .map(t -> {
//                                CellInfo cellInfo = t.getAnnotation(CellInfo.class);
//                                return Cell.convert(t.getName(), t, cellInfo);
//                            }).collect(Collectors.toMap(Cell::getColumnName, t -> t));
//        }
//
//        private Class<?> clazz;
//        private String regex;
//        private int groupName;
//        private int group;
//        private int columnNameGroup;
//        private Map<String, Cell> properties;
//    }
//
//    @Data
//    @NoArgsConstructor
//    static class Cell {
//        private boolean isPrimary;
//        private String propertyName;
//        private String columnName;
//        private String fullColumnName;
//        private int excelIndex;
//        private PropertyTypeEnum propertyType;
//        private String format;
//        private boolean notNull;
//        private int length;
//
//        /**
//         * 设置字段格式
//         *
//         * @param propertyName
//         * @param field
//         * @param cellInfo
//         * @return
//         */
//        static Cell convert(String propertyName, Field field, CellInfo cellInfo) {
//            String format = cellInfo.format();
//            // 是否为Date格式
//            PropertyTypeEnum propertyType = PropertyTypeEnum.STRING;
//            if (field.getType().isAssignableFrom(String.class)) {
//            } else if (field.getType().isAssignableFrom(Date.class)) {
//                propertyType = PropertyTypeEnum.DATE;
//                // 设置默认格式化数据
//                if (StringUtils.isBlank(format)) {
//                    format = DEFAULT_MONTH_PATTERN;
//                }
//            } else if (field.getType().isAssignableFrom(BigDecimal.class)) {
//                propertyType = PropertyTypeEnum.BIG_DECIMAL;
//            }
//            return new Cell(propertyName, cellInfo.value(), null, cellInfo.isPrimary(),
//                    cellInfo.notNull(), propertyType, format, cellInfo.length());
//        }
//
//        Cell(String propertyName,
//             String columnName, String fullColumnName, boolean isPrimary,
//             boolean notNull, PropertyTypeEnum propertyType, String format, int length) {
//            this.propertyName = propertyName;
//            this.isPrimary = isPrimary;
//            this.columnName = columnName;
//            this.fullColumnName = fullColumnName;
//            this.propertyType = propertyType;
//            this.format = format;
//            this.notNull = notNull;
//            this.length = length;
//        }
//    }
//
//    @Data
//    static class CellGroup {
//        private String type;    // 06
//        private String groupName;
//        private Class<?> clazz;
//        private List<Cell> properties = new LinkedList<>();
//
//        CellGroup(String type, Class<?> clazz) {
//            this.type = type;
//            this.clazz = clazz;
//        }
//
//        void addCell(Cell cell, int excelIndex) {
//            Cell newCell = new Cell(cell.getPropertyName(),
//                    cell.getColumnName(), cell.getFullColumnName(), cell.isPrimary(), cell.isNotNull(),
//                    cell.getPropertyType(), cell.getFormat(), cell.getLength());
//            newCell.setExcelIndex(excelIndex);
//            properties.add(newCell);
//        }
//    }
//
//    @Data
//    @NoArgsConstructor
//    @AllArgsConstructor
//    private static class ExcelEntity {
//        private List<Cell> properties;
//        private List<CellGroup> groups;
//        private List<String> missColumns;
//    }
//
//    /**
//     * 解析[excel]单元格数据
//     *
//     * @param cell
//     * @return
//     */
//    public static String getCellValue(org.apache.poi.ss.usermodel.Cell cell) {
//        String value = null;
//        if (cell == null) {
//            return null;
//        }
//        CellType cellType = cell.getCellTypeEnum();
//        String mark = null;
//        if (cellType == CellType.FORMULA) {
//            cellType = cell.getCachedFormulaResultTypeEnum();
//            if (cell instanceof StreamingCell && cell.getCachedFormulaResultTypeEnum() == CellType.FORMULA) {
//                mark = "\"";
//                cellType = CellType.STRING;
//            }
//        }
//        switch (cellType) {
//            case BOOLEAN:
//                value = cell.getBooleanCellValue() ? "true" : "false";
//                break;
//            case ERROR:
////                value = String.format("0x%02x", cell.getStringCellValue());
//                break;
//            case STRING:
//                value = cell.getStringCellValue();
//                if (Objects.nonNull(mark)) {
//                    value = value.replace(mark, "");
//                }
//                value = value.trim();
//                value = value.replaceAll("\u00A0", "");
//                break;
//            case NUMERIC:
//                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
//                    value = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(cell.getDateCellValue());
//                } else {
//                    Double valueInDouble = cell.getNumericCellValue();
//                    value = formatNumberToString(valueInDouble);
//                }
//                break;
//            case BLANK:
//                break;
//        }
//        return value;
//    }
//}
