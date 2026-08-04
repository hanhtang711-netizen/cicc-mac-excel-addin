export interface RequirementChecker {
  isSetSupported(name: string, minimumVersion: string): boolean;
}

export interface AddinCapabilities {
  base: boolean;
  explodedPie: boolean;
  trendlines: boolean;
}

export function getCapabilities(requirements: RequirementChecker): AddinCapabilities {
  const excelApi19 = requirements.isSetSupported("ExcelApi", "1.9");
  return {
    base: excelApi19,
    explodedPie: excelApi19,
    trendlines: excelApi19,
  };
}
