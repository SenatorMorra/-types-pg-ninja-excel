import excel from "exceljs";

export default class ExcelExporter {
  max_width: number;
  wrap_words: boolean;

  constructor(max_width: number = 50, wrap_words: boolean = true) {
    this.max_width = max_width;
    this.wrap_words = wrap_words;
  }

  create_file_name(): string {
    const name = "postgresql-report";
    const dateStamp = new Date()
      .toLocaleString()
      .replace(/\./g, "-")
      .replace(/\//g, "-")
      .replace(", ", "T")
      .replace(/ /g, "");
    const randomString = (Math.random() + 1).toString(36).substring(7);
    const extension = ".xlsx";
    return `${name}-${dateStamp}-${randomString}${extension}`;
  }

  async pg_to_excel(rows: any[], path: string = "./"): Promise<boolean> {
    return new Promise((resolve, reject) => {
      const workbook = new excel.Workbook();

      workbook.creator = "pg-ninja-excel.ts";
      workbook.lastModifiedBy = "pg-ninja-excel.ts";
      workbook.created = new Date();
      workbook.modified = new Date();
      workbook.lastPrinted = new Date();

      const sheet = workbook.addWorksheet("PostgreSQL Result");
      const worksheet = workbook.getWorksheet("PostgreSQL Result")!;

      const header: { header: string; key: string }[] = [];
      let headerLength = 0;
      if (rows.length > 0) { // Check if rows array is not empty
        for (const key in rows[0]) {
          headerLength++;
          header.push({ header: key, key: key });
        }
      }


      worksheet.columns = header;
      worksheet.views = [{ state: "frozen", xSplit: headerLength, ySplit: 1 }];

      worksheet.addRows(rows);

      worksheet.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: 1, column: headerLength },
      };

      for (let i = 1; i <= headerLength; i++) {
        worksheet.getColumn(i).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        if (this.wrap_words) {
          worksheet.getColumn(i).alignment = { wrapText: true };
        }
      }

      worksheet.columns.forEach((column) => {
        let width = 8;

        column.eachCell((cell) => {
          const currentWidth = cell.value ? cell.value.toString().length : 8;
          if (currentWidth > width) {
            width = currentWidth;
          }
        });

        column.width = Math.min(this.max_width, width + 3);
      });

      const filePath = path.endsWith("/") ? path + this.create_file_name() : path;
      workbook.xlsx.writeFile(filePath)
        .then(() => resolve(true))
        .catch((err) => reject(false));
    });
  }
}
