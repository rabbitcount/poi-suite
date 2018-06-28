//package com.ocelot.poi;
//
//import com.ocelot.poi.annotation.ExportCellInfo;
//import lombok.Getter;
//import lombok.NoArgsConstructor;
//import lombok.Setter;
//import org.apache.commons.lang3.StringUtils;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.streaming.SXSSFWorkbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//import java.lang.reflect.Method;
//import java.math.BigDecimal;
//import java.text.SimpleDateFormat;
//import java.util.*;
//
///**
// * @description:
// * @author: 盗梦的虎猫
// * @date: 2018/6/28
// * @time: 6:34 PM
// * Copyright (C) 2018 Anocelot
// * All rights reserved
// */
//public final class ExcelExportUtils {
//
//    public static Workbook buildStreamWorkbook() {
//        return new SXSSFWorkbook();
//    }
//
//    public static Workbook buildXlsxWorkbook() {
//        return new XSSFWorkbook();
//    }
//
//    public static String getWorkbookType(Workbook workbook) {
//        if (workbook instanceof HSSFWorkbook) {
//            return "xls";
//        }
//        return "xlsx";
//    }
//
////    private ExcelExportUtil() {
////    }
//
//    Workbook buildWorkbook(List<SheetContext> sheetContextList) {
//        Workbook workbook = ExcelBaseUtil.buildStreamWorkbook();
//        for (SheetContext sheetContext : sheetContextList) {
//            buildWorkbookSheet(workbook, sheetContext);
//        }
//        return workbook;
//    }
//
//    private void buildWorkbookSheet(Workbook workbook, SheetContext<?, ?> sheetContext) {
//        Sheet sheet = workbook.createSheet(sheetContext.getSheetName());
//        buildWorkbookSheetHeader(sheet, sheetContext);
//        writeSheetData(sheet, sheetContext);
//        autoColumnWidth(workbook, sheet, sheetContext);
//    }
//
//    private void autoColumnWidth(Workbook workbook, Sheet sheet, SheetContext<?, ?> sheetContext) {
//        if (workbook instanceof SXSSFWorkbook) {
//            // 流式处理中，不能进行autosize
//            return;
//        }
//        for (MethodContext methodContext : sheetContext.getContext()) {
//            sheet.autoSizeColumn(methodContext.getIndex());
//        }
//    }
//
//    private void buildWorkbookSheetHeader(Sheet sheet, SheetContext<?, ?> sheetContext) {
//        Row headerRow = sheet.createRow(0);
//        List<MethodContext> methodContexts = sheetContext.getContext();
//        for (MethodContext methodContext : methodContexts) {
//            headerRow.createCell(methodContext.getIndex()).setCellValue(methodContext.columnName);
//        }
//    }
//
//    private void writeSheetData(Sheet sheet, SheetContext<?, ?> sheetContext) {
//        List<MethodContext> methodContexts = sheetContext.getContext();
//
//        int row = 0;
//        for (Object data : sheetContext.getData()) {
//            Row currentRow = sheet.createRow(++row);
//            for (MethodContext context : methodContexts) {
//                currentRow.createCell(context.getIndex())
//                        .setCellValue(getCellValue(data, context));
//            }
//        }
//    }
//
//    private static String getCellValue(Object data, MethodContext context) {
//        try {
//            Object retValue = context.getMethod().invoke(data);
//            if (Objects.isNull(retValue)) {
//                return StringUtils.EMPTY;
//            }
//            switch (context.getReturnType()) {
//                case DATE:
//                    SimpleDateFormat simpleDateFormat = new SimpleDateFormat(context.getFormat());
//                    return simpleDateFormat.format((Date) retValue);
//                case BIG_DECIMAL:
//                    // 默认保留两位
//                    BigDecimal value = ((BigDecimal) retValue).setScale(2, BigDecimal.ROUND_HALF_UP);
//                    return value.toString();
//                default:
//                    return String.valueOf(retValue);
//            }
//        } catch (Exception e) {
//            return "数值计算异常";
//        }
//    }
//
//    public static void main(String[] args) {
//        BigDecimal g = new BigDecimal("123");
//        g = g.setScale(2);
//        System.out.println(g);
//    }
//
//    public static class Builder {
//
//        private List<SheetContext> sheetContextList = Lists.newArrayList();
//
//        private Builder() {
//            // 隐藏
//        }
//
//        public static Builder init() {
//            return new Builder();
//        }
//
//        public <T, F extends T> Builder appendSheet(String sheetName, List<T> data, Class<F> clazz) {
//            SheetContext context = new SheetContext<T, F>(sheetName, data, buildMethodContext(clazz));
//            sheetContextList.add(context);
//            return this;
//        }
//
//        private <F> List<MethodContext> buildMethodContext(Class<F> clazz) {
//            Class<? super F> superClazz = clazz.getSuperclass();
//            List<MethodContext> context = new ArrayList<>();
//            for (Method method : clazz.getDeclaredMethods()) {
//                ExportCellInfo exportCellInfo = method.getAnnotation(ExportCellInfo.class);
//                if (Objects.isNull(exportCellInfo)) {
//                    continue;
//                }
//                if (StringUtils.equals(method.getReturnType().getName(), "void") || method.getParameterCount() > 0) {
//                    // 异常
//                    throw new RuntimeException("返回值为空或包含参数");
//                }
//                try {
//                    if (Objects.nonNull(superClazz)) {
//                        method = superClazz.getDeclaredMethod(method.getName());
//                    }
//                } catch (Exception e) {
//                    // 非Override的方法
//                }
//                context.add(new MethodContext(exportCellInfo, method));
//            }
//            // 基于index进行排序，并转成array
//            context.sort(Comparator.comparing(MethodContext::getIndex));
//            return context;
//        }
//
//        public Workbook build() {
//            ExcelExportUtil util = new ExcelExportUtil();
//            return util.buildWorkbook(this.sheetContextList);
//        }
//    }
//
//    @Setter
//    @Getter
//    static class SheetContext<T> {
//
//        SheetContext(String sheetName, List<T> data, List<MethodContext> context) {
//            this.sheetName = sheetName;
//            this.data = data;
//            this.context = context;
//        }
//
//        private String sheetName;
//        private List<T> data;
//        List<MethodContext> context;
//    }
//
//    public enum ReturnType {
//        DATE, BIG_DECIMAL, STRING;
//    }
//
//    @Setter
//    @Getter
//    @NoArgsConstructor
//    static class MethodContext {
//
//        MethodContext(ExportCellInfo exportCellInfo, Method method) {
//            this.index = exportCellInfo.index();
//            this.columnName = exportCellInfo.columnName();
//            this.format = exportCellInfo.format();
//            this.method = method;
//            this.returnType = ReturnType.STRING;
//            if (method.getReturnType().isAssignableFrom(Date.class)) {
//                this.returnType = ReturnType.DATE;
//                if (StringUtils.isBlank(this.format)) {
//                    this.format = "yyyy-MM";
//                }
//            } else if (method.getReturnType().isAssignableFrom(BigDecimal.class)) {
//                this.returnType = ReturnType.BIG_DECIMAL;
//            }
//        }
//
//        private int index;
//        private String columnName;
//        private String format;
//        private Method method;
//        private ReturnType returnType;
//    }
//}
