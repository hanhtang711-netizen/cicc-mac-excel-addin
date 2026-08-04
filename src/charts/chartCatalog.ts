import type { ChartKind } from "../core/types";

export const CHART_CATALOG = {
  column: { actionId: "createColumnChart", label: "普通柱形图", excelType: "columnClustered" },
  columnStacked: { actionId: "createStackedColumnChart", label: "堆积柱形图", excelType: "columnStacked" },
  line: { actionId: "createLineChart", label: "普通折线图", excelType: "line" },
  lineMarkers: { actionId: "createMarkedLineChart", label: "带标记折线图", excelType: "lineMarkers" },
  pie: { actionId: "createPieChart", label: "饼图", excelType: "pie" },
  bar: { actionId: "createBarChart", label: "条形图", excelType: "barClustered" },
  scatterTrend: { actionId: "createScatterTrendChart", label: "散点图＋趋势线", excelType: "xyscatter" },
  pieExploded: { actionId: "createExplodedPieChart", label: "分离饼图", excelType: "pieExploded" },
  columnStacked100: {
    actionId: "createPercentageStackedChart",
    label: "百分比堆积图",
    excelType: "columnStacked100",
  },
  lineStacked: { actionId: "createStackedLineChart", label: "堆积折线图", excelType: "lineStacked" },
} as const satisfies Record<ChartKind, { actionId: string; label: string; excelType: string }>;
