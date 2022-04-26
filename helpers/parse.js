const {
	styleWorksheet,
	getCommonCellStyle,
	getLongestString,
	getNumberOfLines,
} = require("./styling");

function compileDataIntoWorkbook(workbook, data) {
	const length = data.length;

	// Sheets
	for (let sheetNumber = 0; sheetNumber < length; sheetNumber++) {
		const { name, sheetData, stylingLocation } = data[sheetNumber];
		const worksheet = workbook.addWorksheet(`${name}`);

		const sheetDataLength = sheetData.length + 1;
		// Rows
		for (let rowLocation = 1; rowLocation < sheetDataLength; rowLocation++) {
			const rowData = sheetData[rowLocation];
			// Columns
			for (let col in rowData) {
				const colLocation = col.toLowerCase().charCodeAt(0) - 96;
				const val = rowData[col];
				worksheet.cell(rowLocation, colLocation).string(val);
				worksheet
					.cell(rowLocation, colLocation)
					.style(getCommonCellStyle(workbook));

				const height = getNumberOfLines(val) * 14 + 70;
				if (worksheet.row(rowLocation).ht < height) {
					worksheet.row(rowLocation).setHeight(height);
				}

				let skip = false;
				const tablesLength = stylingLocation.tables.length;
				for (let tableIndex = 0; tableIndex < tablesLength; tableIndex++) {
					const { index } = stylingLocation.tables[tableIndex];

					if (index - 1 === rowLocation) {
						skip = true;
						break;
					}
				}

				if (skip) continue;

				const width = getLongestString(val);
				if (worksheet.column(colLocation).colWidth < width) {
					worksheet.column(colLocation).setWidth(width);
				}
			}
		}

		styleWorksheet(workbook, worksheet, stylingLocation);
	}

	return workbook;
}

module.exports = {
	compileDataIntoWorkbook,
};
