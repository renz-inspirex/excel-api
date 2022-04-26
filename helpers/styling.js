function styleWorksheet(workbook, worksheet, stylingLocation) {
	// STYLES
	const border = {
		left: {
			style: "thin",
			color: "#000000",
		},
		right: {
			style: "thin",
			color: "#000000",
		},
		top: {
			style: "thin",
			color: "#000000",
		},
		bottom: {
			style: "thin",
			color: "#000000",
		},
	};
	const tableHeaderCellStyle = workbook.createStyle({
		alignment: {
			horizontal: "center",
			vertical: "center",
			wrapText: true,
		},
		font: {
			bold: true,
			size: 15,
			vertAlign: "center",
		},
		border,
	});
	const tableCellStyle = workbook.createStyle({
		alignment: {
			horizontal: "left",
			vertical: "center",
			shrinkToFit: true,
			wrapText: true,
		},
		font: {
			size: 14,
			vertAlign: "center",
		},
		border,
	});

	const { tables } = stylingLocation;

	tables?.forEach((value) => {
		const { index, size } = value;

		// Rows
		for (let i = 0; i < size; i++) {
			const { cellRefs } = worksheet.row(index + i);
			if (!cellRefs) continue;
			// Columns
			cellRefs?.forEach((value) => {
				const regexStatement = /^([a-zA-Z]+)([0-9]+)$/;
				const matches = value.match(regexStatement);
				if (i == 0) {
					worksheet
						.cell(matches[2], matches[1].toLowerCase().charCodeAt(0) - 96)
						.style(tableHeaderCellStyle);
				} else {
					worksheet
						.cell(matches[2], matches[1].toLowerCase().charCodeAt(0) - 96)
						.style(tableCellStyle);
				}
			});
		}
	});
}

function getLongestString(val) {
	const values = val.split("\r\n");

	let maxLength = 1;
	values.forEach((value) => {
		if (value.length > maxLength) {
			maxLength = value.length;
		}
	});

	return maxLength;
}

function getNumberOfLines(val) {
	const numberOfLines = val.split("\r\n")?.length;
	if (!numberOfLines) {
		return 1;
	}
	return numberOfLines;
}

function getCommonCellStyle(workbook) {
	return workbook.createStyle({
		alignment: {
			horizontal: "left",
			vertical: "center",
			wrapText: false,
		},
		font: {
			size: 14,
		},
	});
}

const styles = (module.exports = {
	styleWorksheet,
	getLongestString,
	getNumberOfLines,
	getCommonCellStyle,
});
