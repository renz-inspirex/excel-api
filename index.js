const express = require("express");
app = express();
const path = require("path");
const { compileDataIntoWorkbook } = require("./helpers/parse");

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static("files"));
app.set("view engine", "ejs");

app.get("/", function (req, res) {
	res.send("This is a sample app");
});

app.use("/files", express.static(path.join(__dirname, "")));

app.post("/excel", async (request, response) => {
	//code to perform particular action.
	//To access POST variable use req.body()methods.

	const { body } = request;

	let xl = require("excel4node");

	let wb = new xl.Workbook();

	wb = compileDataIntoWorkbook(wb, body);

	wb.write("Acceptance Criteria.xlsx", response);
});

app.listen(process.env.PORT || 3000, () => {
	console.log("[EXPRESS] Web Server active on: " + 3000);
});
