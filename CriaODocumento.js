function CriarDocumento() {
	var Body = DocumentApp.create("Tabela de Trabalhos").getBody();
	var Range = SpreadsheetApp.openByUrl(
		"https://docs.google.com/spreadsheets/d/seuarquivo/edit#gid=0"
	)
		.getRange("planilhadesejada! range A2:H9")
		.getDisplayValues();
	for (let index = 0; index < Range.length; index++) {
		var Data = new Date(Range[index][0]);

		if (Data.getDay() == 0 || Data.getDay() == 6) {
			CriarTabeladoFinaldeSemana(Body, Range, index, Data);
		} else if (Data.getDay() == 3 || Data.getDay() == 2) {
			CriarTabelaMeiodeSemana(Body, Range, index, Data);
		}
	}
	var style12 = {};
	style12[DocumentApp.Attribute.LINE_SPACING] = 0.5;
	var Tabelas = Body.getTables();
	Tabelas[0].setAttributes(style12);
}

function CriarTabeladoFinaldeSemana(Body, Range, index, Data) {
	const option = {
		weekday: "long",
		day: "numeric",
		month: "long",
	};
	const locale = "pt-br";

	var TabeladoFinaldeSemana = [
		["Caminhão:", Range[index][7]],
		["Trator D:", Range[index][6]],
		[
			"Caminhonete 1, Caminhonete 2, Viagens e Organização:",
			Range[index][1].split(" ")[0] +
				"," +
				Range[index][2].split(" ")[0] +
				"," +
				Range[index][3].split(" ")[0] +
				"," +
				Range[index][4].split(" ")[0],
		],
	];
	var style = {};
	(style[DocumentApp.Attribute.FOREGROUND_COLOR] = "#0303b1"),
		(style[DocumentApp.Attribute.BORDER_COLOR] = "#0303b1"),
		(style[DocumentApp.Attribute.FONT_FAMILY] = "Arial"),
		(style[DocumentApp.Attribute.LINE_SPACING] = 0.5),
		(style[DocumentApp.Attribute.FONT_SIZE] = 9),
		(style[DocumentApp.Attribute.SPACING_AFTER] = 0.5),
		(style[DocumentApp.Attribute.SPACING_BEFORE] = 0.5);

	Body.appendParagraph(Data.toLocaleDateString(locale, option)).setAttributes(
		style
	);
	Body.appendTable(TabeladoFinaldeSemana).setAttributes(style);
}

function CriarTabelaMeiodeSemana(Body, Range, index, Data) {
	const option = {
		weekday: "long",
		day: "numeric",
		month: "long",
	};
	const locale = "pt-br";

	var TabeladaSemana = [
		["Caminhão:", Range[index][7]],
		["Trator C:", Range[index][5]],
		[
			"Caminhonete 1, Caminhonete 2, Viagens e Organização:",
			Range[index][1].split(" ")[0] +
				"," +
				Range[index][2].split(" ")[0] +
				"," +
				Range[index][3].split(" ")[0] +
				"," +
				Range[index][4].split(" ")[0],
		],
	];
	var style = {};
	(style[DocumentApp.Attribute.FOREGROUND_COLOR] = "#cc0000"),
		(style[DocumentApp.Attribute.BORDER_COLOR] = "#cc0000"),
		(style[DocumentApp.Attribute.FONT_FAMILY] = "Arial"),
		(style[DocumentApp.Attribute.LINE_SPACING] = 0.5),
		(style[DocumentApp.Attribute.FONT_SIZE] = 9);

	Body.appendParagraph(Data.toLocaleDateString(locale, option)).setAttributes(
		style
	);
	Body.appendTable(TabeladaSemana).setAttributes(style);
}
