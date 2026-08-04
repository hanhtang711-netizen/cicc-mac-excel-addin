export interface PackageReleaseOptions {
  baseUrl: string;
  outDir: string;
  projectRoot?: string;
}

export interface PackageReleaseResult {
  baseUrl: string;
  outDir: string;
  siteDir: string;
  manifestPath: string;
}

export function validateProductionBaseUrl(rawUrl: string): string;
export function packageRelease(options: PackageReleaseOptions): Promise<PackageReleaseResult>;
export function runReleaseCli(argv?: string[]): Promise<PackageReleaseResult>;
