type FakeSelectionInput = {
  address: string;
  formula?: string;
  numberFormat?: string;
  sheetId?: string;
  sheetName?: string;
  value?: string | number | boolean | null;
  workbookId?: string;
};

export const createFakeOfficeContext = () => {
  let syncCount = 0;
  const selectionLoads: string[][] = [];
  const worksheetLoads: string[][] = [];
  const worksheet = {
    id: "sheet-1",
    load(properties: string[]) {
      worksheetLoads.push(properties);
    },
    name: "Sheet1",
  };
  const selection: {
    address: string;
    columnCount: number;
    formulas: unknown[][];
    load(properties: string[]): void;
    numberFormat: unknown[][];
    rowCount: number;
    values: unknown[][];
    worksheet: typeof worksheet;
  } = {
    address: "Sheet1!A1",
    columnCount: 1,
    formulas: [[""]],
    load(properties: string[]) {
      selectionLoads.push(properties);
    },
    numberFormat: [["General"]],
    rowCount: 1,
    values: [[null]],
    worksheet,
  };
  const context = {
    getSelection: () => selection,
    async sync() {
      syncCount += 1;
    },
    workbookId: "workbook-1",
  };
  const fake = {
    context,
    selectionLoads,
    get syncCount() {
      return syncCount;
    },
    withSelection(input: FakeSelectionInput) {
      context.workbookId = input.workbookId ?? "workbook-1";
      selection.address = input.address;
      selection.formulas = [[input.formula ?? input.value ?? ""]];
      selection.numberFormat = [[input.numberFormat ?? "General"]];
      selection.values = [[input.value ?? null]];
      worksheet.id = input.sheetId ?? "sheet-1";
      worksheet.name = input.sheetName ?? "DCF";
      return fake;
    },
    worksheetLoads,
  };
  return fake;
};
