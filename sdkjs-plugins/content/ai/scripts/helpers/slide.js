/*
 * (c) Copyright Ascensio System SIA 2010-2025
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

function getSlideFunctions() {
	let funcs = [];

	if (true) {
		let func = new RegisteredFunction();
		func.name = "addNewSlide";
		func.description = "Adds a new slide at the end of presentation using default layout from current slide's master";
		func.params = [];
		func.examples = ["if you need to add a new slide, respond with:\n" + "[functionCalling (addNewSlide)]: {}"];

		func.call = async function (params) {
			await Asc.Editor.callCommand(function () {
				let presentation = Api.GetPresentation();
				let currentSlide = presentation.GetCurrentSlide();
				if (!currentSlide) {
					currentSlide = presentation.GetSlideByIndex(0);
				}

				let curLayout = currentSlide.GetLayout();
				let master = curLayout.GetMaster();
				let layoutsCount = master.GetLayoutsCount();

				let layoutIndex = 1;
				if (layoutsCount <= 1) {
					layoutIndex = 0;
				}

				let layout = master.GetLayoutByType("obj");
				if (!layout && layoutsCount > 0) {
					layout = master.GetLayout(0);
				}

				let newSlide = Api.CreateSlide();

				if (layout && newSlide.ApplyLayout) {
					newSlide.ApplyLayout(layout);
				}

				presentation.AddSlide(newSlide);
			});
		};
		funcs.push(func);
	}

	if (true) {
		let func = new RegisteredFunction();
		func.name = "addShapeToSlide";
		func.description = "Adds a shape to the slide with optional text (139x42mm, centered, blue fill with dark border)";
		func.params = ["slideNumber (number): slide number to add shape to (optional, defaults to current)", "shapeType (string): shape type - rect, roundRect, ellipse, triangle, diamond, pentagon, hexagon, star5, plus, mathMinus, mathMultiply, mathEqual, mathNotEqual, heart, cloud, leftArrow, rightArrow, upArrow, downArrow, leftRightArrow, chevron, bentArrow, curvedRightArrow, blockArc, wedgeRectCallout, cloudCallout, ribbon, wave, can, cube, pie, donut, sun, moon, smileyFace, lightningBolt, noSmoking (optional, defaults to roundRect)", "text (string): text to add to the shape (optional)"];
		func.examples = ["if you need to add a rectangle with text on slide 2, respond with:\n" + "[functionCalling (addShapeToSlide)]: {\"slideNumber\": 2, \"shapeType\": \"rect\", \"text\": \"Important Point\"}", "if you need to add a star shape on current slide, respond with:\n" + "[functionCalling (addShapeToSlide)]: {\"shapeType\": \"star5\"}", "if you need to add a rounded rectangle with text, respond with:\n" + "[functionCalling (addShapeToSlide)]: {\"text\": \"Key Message\"}", "if you need to add a diamond shape with text, respond with:\n" + "[functionCalling (addShapeToSlide)]: {\"shapeType\": \"diamond\", \"text\": \"Decision Point\"}", "if you need to add a right arrow with text, respond with:\n" + "[functionCalling (addShapeToSlide)]: {\"shapeType\": \"rightArrow\", \"text\": \"Next Step\"}"];

		func.call = async function (params) {
			Asc.scope.params = params;

			await Asc.Editor.callCommand(function () {
				let presentation = Api.GetPresentation();
				let slide;

				if (Asc.scope.params.slideNumber) {
					slide = presentation.GetSlideByIndex(Asc.scope.params.slideNumber - 1);
				}
				else {
					slide = presentation.GetCurrentSlide();
				}

				if (!slide) return;

				let slideWidth = presentation.GetWidth();
				let slideHeight = presentation.GetHeight();

				let shapeType = Asc.scope.params.shapeType || "rect";
				let width = 2500000;
				let height = 2500000;
				let x = (slideWidth - width) / 2;
				let y = (slideHeight - height) / 2;

				let fill = Api.CreateSolidFill(Api.CreateRGBColor(91, 155, 213));
				let stroke = Api.CreateStroke(12700, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));

				let shape = Api.CreateShape(shapeType, width, height, fill, stroke);
				shape.SetPosition(x, y);


				if (Asc.scope.params.text) {
					let docContent = shape.GetDocContent();
					if (docContent) {
						docContent.RemoveAllElements();
						let paragraph = Api.CreateParagraph();
						paragraph.SetJc("center");
						let run = paragraph.AddText(Asc.scope.params.text);
						run.SetFontSize(32);
						run.SetBold(true);
						run.SetFill(Api.CreateSolidFill(Api.CreateRGBColor(255, 255, 255)));
						docContent.Push(paragraph);

						shape.SetVerticalTextAlign("center");
					}
				}

				slide.AddObject(shape);
			});
		};
		funcs.push(func);
	}
	if (true) {
		let func = new RegisteredFunction();
		func.name = "changeSlideBackground";
		func.params = ["slideNumber (number): the slide number to change background", "backgroundType (string): type of background - 'solid', 'gradient', 'image'", "color (string): hex color for solid background (e.g., '#FF5733')", "imageUrl (string): URL for image background", "gradientColors (array): array of hex colors for gradient"];
		func.examples = ["if you need to set blue background on slide 1, respond with:\n" + "[functionCalling (changeSlideBackground)]: {\"slideNumber\": 1, \"backgroundType\": \"solid\", \"color\": \"#0066CC\"}", "if you need to set gradient background, respond with:\n" + "[functionCalling (changeSlideBackground)]: {\"slideNumber\": 2, \"backgroundType\": \"gradient\", \"gradientColors\": [\"#FF0000\", \"#0000FF\"]}"];

		func.call = async function (params) {
			Asc.scope.params = params;

			await Asc.Editor.callCommand(function () {
				let presentation = Api.GetPresentation();
				let slide = presentation.GetSlideByIndex(Asc.scope.params.slideNumber - 1);
				if (!slide) return;

				let fill;

				switch (Asc.scope.params.backgroundType) {
					case "solid":
						if (Asc.scope.params.color) {
							let rgb = parseInt(Asc.scope.params.color.slice(1), 16);
							let r = (rgb >> 16) & 255;
							let g = (rgb >> 8) & 255;
							let b = rgb & 255;
							fill = Api.CreateSolidFill(Api.CreateRGBColor(r, g, b));
						}
						break;

					case "gradient":
						if (Asc.scope.params.gradientColors && Asc.scope.params.gradientColors.length >= 2) {
							let stops = [];
							let step = 100000 / (Asc.scope.params.gradientColors.length - 1);

							Asc.scope.params.gradientColors.forEach((color, index) => {
								let rgb = parseInt(color.slice(1), 16);
								let r = (rgb >> 16) & 255;
								let g = (rgb >> 8) & 255;
								let b = rgb & 255;
								let stop = Api.CreateGradientStop(Api.CreateRGBColor(r, g, b), index * step);
								stops.push(stop);
							});

							fill = Api.CreateLinearGradientFill(stops, 5400000);
						}
						break;

					case "image":
						if (Asc.scope.params.imageUrl) {
							fill = Api.CreateBlipFill(Asc.scope.params.imageUrl, "stretch");
						}
						break;
				}

				if (fill) {
					slide.SetBackground(fill);
				}
			});
		};
		funcs.push(func);
	}

	if (true) {
		let func = new RegisteredFunction();
		func.name = "addImageByDescription";
		func.params = ["slideNumber (number): the slide number to add generated image to", "description (string): text description of the image to generate", "x (number, optional): horizontal position in mm (default: 25)", "y (number, optional): vertical position in mm (default: 25)", "width (number, optional): image width in mm (default: 150)", "height (number, optional): image height in mm (default: 100)", "style (string, optional): image style (realistic, cartoon, abstract, etc.)"];
		func.examples = ["if you need to add an image of a sunset over mountains to slide 1, respond with:\n" + "[functionCalling (addImageByDescription)]: {\"slideNumber\": 1, \"description\": \"beautiful sunset over mountain range with orange and purple sky\"}", "if you need to add a cartoon style image of office workers, respond with:\n" + "[functionCalling (addImageByDescription)]: {\"slideNumber\": 2, \"description\": \"team of diverse office workers collaborating around a table\", \"style\": \"cartoon\", \"x\": 50, \"y\": 30}"];

		func.call = async function (params) {
			Asc.scope.slideNum = params.slideNumber - 1;
			Asc.scope.description = params.description;
			Asc.scope.x = (params.x || 25) * 36000;
			Asc.scope.y = (params.y || 25) * 36000;
			Asc.scope.width = (params.width || 150) * 36000;
			Asc.scope.height = (params.height || 100) * 36000;
			Asc.scope.style = params.style || "realistic";

			let requestEngine = AI.Request.create(AI.ActionType.ImageGeneration);
			if (!requestEngine) {
				console.error("Image generation is not available");
				return;
			}

			// Формируем промпт с учетом стиля
			let fullPrompt = Asc.scope.description;
			if (Asc.scope.style && Asc.scope.style !== "realistic") {
				fullPrompt = Asc.scope.style + " style, " + fullPrompt;
			}

			let imageUrl = await requestEngine.imageGenerationRequest(fullPrompt);
			debugger;
			if (imageUrl) {
				// Вставляем сгенерированное изображение на слайд
				Asc.scope.imageUrl = imageUrl;
				await Asc.Editor.callCommand(function () {
					let oPresentation = Api.GetPresentation();
					let oSlide = oPresentation.GetSlideByIndex(Asc.scope.slideNum);
					if (!oSlide) return;

					let oImage = Api.CreateImage(Asc.scope.imageUrl, Asc.scope.width, Asc.scope.height);
					oImage.SetPosition(Asc.scope.x, Asc.scope.y);
					oSlide.AddObject(oImage);
				});
			}
		};
		funcs.push(func);
	}

	if (true) {
		let func = new RegisteredFunction();
		func.name = "addTableToSlide";
		func.description = "Adds a table to the slide (194x97mm, centered)";
		func.params = ["slideNumber (number): slide number to add table to (optional, defaults to current)", "rows (number): number of rows (optional, defaults to 3)", "columns (number): number of columns (optional, defaults to 3)", "data (array): 2D array of cell values - rows x columns (optional)"];
		func.examples = ["if you need to add a 3x3 table on slide 2, respond with:\n" + "[functionCalling (addTableToSlide)]: {\"slideNumber\": 2, \"rows\": 3, \"columns\": 3}", "if you need to add a table with data on current slide, respond with:\n" + "[functionCalling (addTableToSlide)]: {\"data\": [[\"Name\", \"Age\", \"City\"], [\"John\", \"30\", \"New York\"], [\"Jane\", \"25\", \"London\"]]}", "if you need to add a simple 2x4 table, respond with:\n" + "[functionCalling (addTableToSlide)]: {\"rows\": 2, \"columns\": 4}"];

		func.call = async function (params) {
			Asc.scope.params = params;

			await Asc.Editor.callCommand(function () {
				let presentation = Api.GetPresentation();
				let slide;

				if (Asc.scope.params.slideNumber) {
					slide = presentation.GetSlideByIndex(Asc.scope.params.slideNumber - 1);
				}
				else {
					slide = presentation.GetCurrentSlide();
				}

				if (!slide) return;

				let slideWidth = presentation.GetWidth();
				let slideHeight = presentation.GetHeight();

				let data = Asc.scope.params.data;
				let rows = Asc.scope.params.rows || 3;
				let columns = Asc.scope.params.columns || 3;

				if (data && Array.isArray(data) && data.length > 0) {
					rows = data.length;
					if (data[0] && Array.isArray(data[0])) {
						columns = data[0].length;
					}
				}

				let tableWidth = 7000000;
				let tableHeight = 3500000;
				let x = (slideWidth - tableWidth) / 2;
				let y = (slideHeight - tableHeight) / 2;

				let table = Api.CreateTable(columns, rows);

				if (table) {
					table.SetPosition(x, y);
					table.SetSize(tableWidth, tableHeight);
					let rowHeight = tableHeight / rows;
					if (data && Array.isArray(data)) {
						for (let row = 0; row < Math.min(data.length, rows); row++) {
							if (Array.isArray(data[row])) {
								for (let col = 0; col < Math.min(data[row].length, columns); col++) {
									let cell = table.GetCell(row, col);
									if (cell && cell.GetContent) {
										let cellContent = cell.GetContent();
										if (cellContent) {
											cellContent.RemoveAllElements();
											let paragraph = Api.CreateParagraph();
											let value = data[row][col];
											if (value !== null && value !== undefined) {
												paragraph.AddText(value);
												cellContent.Push(paragraph);
											}
										}
									}
								}
							}
						}
					}

					slide.AddObject(table);
				}
			});
		};
		funcs.push(func);
	}

	if (true) {
		let func = new RegisteredFunction();
		func.name = "deleteSlide";
		func.params = ["slideNumber (number): the slide number to delete"];
		func.examples = ["if you need to delete slide 5, respond with:\n" + "[functionCalling (deleteSlide)]: {\"slideNumber\": 5}"];

		func.call = async function (params) {
			Asc.scope.slideNum = params.slideNumber;

			await Asc.Editor.callCommand(function () {
				let presentation = Api.GetPresentation();
				let slide = presentation.GetSlideByIndex(Asc.scope.slideNum - 1);
				if (slide) {
					slide.Delete();
				}
			});
		};
		funcs.push(func);
	}

	if (true) {
		let func = new RegisteredFunction();
		func.name = "duplicateSlide";
		func.params = ["slideNumber (number): the slide number to duplicate"];
		func.examples = ["if you need to duplicate slide 3, respond with:\n" + "[functionCalling (duplicateSlide)]: {\"slideNumber\": 3}"];

		func.call = async function (params) {
			Asc.scope.slideNum = params.slideNumber;

			await Asc.Editor.callCommand(function () {
				let presentation = Api.GetPresentation();
				let slide = presentation.GetSlideByIndex(Asc.scope.slideNum - 1);
				if (slide) {
					let newSlide = slide.Duplicate();
					presentation.AddSlide(newSlide, Asc.scope.slideNum);
				}
			});
		};
		funcs.push(func);
	}

	if (true) {
		let func = new RegisteredFunction();
		func.name = "addChartToSlide";
		func.params = ["slideNumber (number): slide number to add chart to (optional, defaults to current)", "chartType (string): type of chart - bar, barStacked, barStackedPercent, bar3D, barStacked3D, barStackedPercent3D, barStackedPercent3DPerspective, horizontalBar, horizontalBarStacked, horizontalBarStackedPercent, horizontalBar3D, horizontalBarStacked3D, horizontalBarStackedPercent3D, lineNormal, lineStacked, lineStackedPercent, line3D, pie, pie3D, doughnut, scatter, stock, area, areaStacked, areaStackedPercent, comboBarLine, comboBarLineSecondary, comboCustom", "data (array): 2D array of numeric data values - all sub-arrays must have same length, number of arrays must match series count", "series (array): array of series names - must have same length as data arrays count", "categories (array): array of category names - must have same length as each data array", "x (number): x position in EMUs (optional, defaults to center)", "y (number): y position in EMUs (optional, defaults to center)", "width (number): width in EMUs (optional, defaults to 5472000 = 152mm)", "height (number): height in EMUs (optional, defaults to 3204000 = 89mm)", "prompt (string): description of what kind of data to generate for the chart (optional)"];
		func.examples = ["if you need to add a bar chart showing sales data on slide 2, respond with:\n" + "[functionCalling (addChartToSlide)]: {\"slideNumber\": 2, \"chartType\": \"bar3D\", \"data\": [[100, 120, 140], [90, 110, 130]], \"series\": [\"Product A\", \"Product B\"], \"categories\": [\"Q1\", \"Q2\", \"Q3\"]}", "if you need to add a pie chart on current slide, respond with:\n" + "[functionCalling (addChartToSlide)]: {\"chartType\": \"pie\", \"data\": [[30, 25, 20, 15, 10]], \"series\": [\"Market Share\"], \"categories\": [\"Company A\", \"Company B\", \"Company C\", \"Company D\", \"Others\"]}", "if you need to add a line chart with 3 series and 4 data points, respond with:\n" + "[functionCalling (addChartToSlide)]: {\"chartType\": \"lineNormal\", \"data\": [[10, 20, 30, 40], [15, 25, 35, 45], [12, 22, 32, 42]], \"series\": [\"Series 1\", \"Series 2\", \"Series 3\"], \"categories\": [\"Jan\", \"Feb\", \"Mar\", \"Apr\"]}", "if you need AI to generate chart data, respond with:\n" + "[functionCalling (addChartToSlide)]: {\"slideNumber\": 3, \"chartType\": \"lineNormal\", \"prompt\": \"Create monthly revenue data for 2024 showing steady growth from $50k to $120k\"}"];

		func.call = async function (params) {
			Asc.scope.params = params;

			if (params.prompt && !params.data) {
				let requestEngine = AI.Request.create(AI.ActionType.Chat);
				if (!requestEngine) return;

				let chartPrompt = "Generate chart data for the following request: " + params.prompt + "\n\nReturn ONLY a JSON object in this exact format (no other text):\n" + "{\n" + "  \"data\": [[number, number, ...], [number, number, ...]],\n" + "  \"series\": [\"Series1\", \"Series2\", ...],\n" + "  \"categories\": [\"Category1\", \"Category2\", ...]\n" + "}\n\n" + "IMPORTANT RULES:\n" + "1. The number of arrays in 'data' MUST equal the number of items in 'series'\n" + "2. ALL arrays in 'data' MUST have exactly the same length\n" + "3. The number of items in 'categories' MUST equal the length of each data array\n" + "Example: if data=[[10,20,30],[40,50,60]], then series must have 2 names and categories must have 3 names";

				let generatedData = await requestEngine.chatRequest(chartPrompt, false);

				try {
					let parsedData = JSON.parse(generatedData);
					Asc.scope.params.data = parsedData.data;
					Asc.scope.params.series = parsedData.series;
					Asc.scope.params.categories = parsedData.categories;

					let dataLength = Asc.scope.params.data.length;
					let seriesLength = Asc.scope.params.series.length;
					let pointsLength = Asc.scope.params.data[0] ? Asc.scope.params.data[0].length : 0;
					let categoriesLength = Asc.scope.params.categories.length;

					for (let i = 1; i < Asc.scope.params.data.length; i++) {
						if (Asc.scope.params.data[i].length !== pointsLength) {
							while (Asc.scope.params.data[i].length < pointsLength) {
								Asc.scope.params.data[i].push(0);
							}
							Asc.scope.params.data[i] = Asc.scope.params.data[i].slice(0, pointsLength);
						}
					}

					if (dataLength !== seriesLength) {
						while (Asc.scope.params.series.length < dataLength) {
							Asc.scope.params.series.push("Series " + (Asc.scope.params.series.length + 1));
						}
						Asc.scope.params.series = Asc.scope.params.series.slice(0, dataLength);
					}

					if (pointsLength !== categoriesLength) {
						while (Asc.scope.params.categories.length < pointsLength) {
							Asc.scope.params.categories.push("Cat " + (Asc.scope.params.categories.length + 1));
						}
						Asc.scope.params.categories = Asc.scope.params.categories.slice(0, pointsLength);
					}
				} catch (e) {
					Asc.scope.params.data = [[100, 120, 140], [90, 110, 130]];
					Asc.scope.params.series = ["Series 1", "Series 2"];
					Asc.scope.params.categories = ["Cat 1", "Cat 2", "Cat 3"];
				}
			}

			await Asc.Editor.callCommand(function () {
				let presentation = Api.GetPresentation();
				let slide;

				if (Asc.scope.params.slideNumber) {
					slide = presentation.GetSlideByIndex(Asc.scope.params.slideNumber - 1);
				}
				else {
					slide = presentation.GetCurrentSlide();
				}

				if (!slide) return;

				// Значения по умолчанию
				let chartType = Asc.scope.params.chartType || "bar3D";
				let data = Asc.scope.params.data || [[100, 120, 140], [90, 110, 130]];
				let series = Asc.scope.params.series || ["Series 1", "Series 2"];
				let categories = Asc.scope.params.categories || ["Category 1", "Category 2", "Category 3"];

				if (!data || data.length === 0 || !data[0] || data[0].length === 0) {
					data = [[100, 120, 140], [90, 110, 130]];
					series = ["Series 1", "Series 2"];
					categories = ["Category 1", "Category 2", "Category 3"];
				}

				if (data.length > 0 && data[0].length > 0) {
					let dataLength = data.length;
					let pointsLength = data[0].length;

					for (let i = 1; i < data.length; i++) {
						if (data[i].length !== pointsLength) {
							while (data[i].length < pointsLength) {
								data[i].push(0);
							}
							data[i] = data[i].slice(0, pointsLength);
						}
					}

					if (series.length !== dataLength) {
						while (series.length < dataLength) {
							series.push("Series " + (series.length + 1));
						}
						series = series.slice(0, dataLength);
					}

					if (categories.length !== pointsLength) {
						while (categories.length < pointsLength) {
							categories.push("Category " + (categories.length + 1));
						}
						categories = categories.slice(0, pointsLength);
					}
				}

				// Размеры диаграммы: 152x89 мм (в EMU: 1мм = 36000 EMU)

				let slideWidth = presentation.GetWidth();
				let slideHeight = presentation.GetHeight();
				let width = Asc.scope.params.width || 5472000;  // 152 мм
				let height = Asc.scope.params.height || 3204000; // 89 мм

				// Центрирование на слайде
				let x = Asc.scope.params.x || (slideWidth - width) / 2;   // центр по горизонтали
				let y = Asc.scope.params.y || (slideHeight - height) / 2; // центр по вертикали

				// Создаем диаграмму
				let chart = Api.CreateChart(chartType, data, series, categories, width, height, 24);

				if (chart) {
					// Устанавливаем позицию
					chart.SetPosition(x, y);

					// Добавляем диаграмму на слайд
					slide.AddObject(chart);

				}
			});
		};
		funcs.push(func);
	}

	return funcs;
}
