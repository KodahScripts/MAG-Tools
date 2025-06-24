class Amount {
	public total: number;
	constructor(public auth: number, public sett: number, public fee: number) {
		this.total = sett - fee;
	}
}

class Row {
	public state: string;
	public brand: string;
	public amount: Amount;
	public invoice: string;
	public user: string;
	constructor(protected row: Array<string|number|boolean>) {
		this.state = String(row[1]);
		this.brand = String(row[2]);
		this.invoice = String(row[6]).length > 1 ? String(row[6]) : String(row[7]);

		const authAmount = this.getAmount(String(row[3]));
		const settAmount = this.getAmount(String(row[4]));
		const feeAmount = this.getAmount(String(row[5]));
		this.amount = new Amount(authAmount, settAmount, feeAmount);

		this.user = String(row[8]);
	}

	getAmount(cellValue: string) {
		if(cellValue === "") return 0;

		let amt = cellValue.split("$")[1];
		if (amt.length === 0) return 0;

		if (amt.includes(',')) {
			return Number(amt.replace(',', ''));
		}
		return Number(amt);
	}
}

function main(workbook: ExcelScript.Workbook) {
	const selectedSheet = workbook.getActiveWorksheet();
	selectedSheet.getRange("A:D").delete(ExcelScript.DeleteShiftDirection.left);
	selectedSheet.getRange("B:C").delete(ExcelScript.DeleteShiftDirection.left);
	selectedSheet.getRange("D:H").delete(ExcelScript.DeleteShiftDirection.left);
	selectedSheet.getRange("I:J").delete(ExcelScript.DeleteShiftDirection.left);
	selectedSheet.getRange("J:K").delete(ExcelScript.DeleteShiftDirection.left);

	const reportData = selectedSheet.getUsedRange().getValues();
	reportData.shift();
	const data = reportData.map((row, index) => {
		return new Row(row);
	});
	
	const approvedReceipts = data.filter(d => d.state === "APPROVAL");
	const settledReceipts = data.filter(d => d.state === "SETTLED");

	const approvedAX = approvedReceipts.filter(rec => rec.brand === "AMEX");
	const approvedOther = approvedReceipts.filter(rec => rec.brand !== "AMEX");
	const settledAX = settledReceipts.filter(rec => rec.brand === "AMEX");
	const settledOther = settledReceipts.filter(rec => rec.brand !== "AMEX");

	if (approvedAX.length > 0) addSheet(workbook, "Approved AX", approvedAX);
	if (approvedOther.length > 0) addSheet(workbook, "Approved Other", approvedOther);
	if (settledAX.length > 0) addSheet(workbook, "Settled AX", settledAX);
	if (settledOther.length > 0) addSheet(workbook, "Settled Other", settledOther);
}

function addSheet(workbook: ExcelScript.Workbook, sheetName: string, data: Row[]) {
	const sheet = workbook.addWorksheet(sheetName);
	sheet.getRange("A1:F1").setValues([["Auth Amount", "Settlement Amount", "Cardholder Surcharge", "Total", "Invoice Number", "User"]]);
	data.forEach((d, i) => {
		const row = i + 2;
		const invoice = d.invoice.length > 6 ? d.invoice.slice(-6) : d.invoice;
		sheet.getRange(`A${row}:F${row}`).setValues([[d.amount.auth, d.amount.sett, d.amount.fee, d.amount.total, invoice, d.user]]);
	});
	const range = sheet.getUsedRange();
	const table = sheet.addTable(range, true);
	const tableLen = table.getRowCount() + 1;
	table.addRow(null, [
		`=SUM(A2:A${tableLen})`,
		`=SUM(B2:B${tableLen})`,
		`=SUM(C2:C${tableLen})`,
		'','',''
	]);
	sheet.getRange("1:1").getFormat().autofitColumns();
}