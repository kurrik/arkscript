/**
 *  Types used by arksheets.
 */

interface ArkCell {
  value: any;
  frozen: boolean;
  hidden: boolean;
  comment: string;
  formula: string;
}

interface ArkSheetsColumn {
  sheet?: string;
  col?: number;
  rows: KeyedData;
}

interface ArkSheetsRow {
  sheet?: string;
  row?: number;
  columns: KeyedData;
}

type KeyedData = { [key: string]: ArkCell };