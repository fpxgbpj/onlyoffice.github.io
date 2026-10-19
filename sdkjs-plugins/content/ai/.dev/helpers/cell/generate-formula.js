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

(function () {

	let func = new RegisteredFunction({
		"name": "generateFormula",
		"text": "Generate Formula",
		"description": "Creates a new formula based on the natural language request.",
		"parameters": {
			"type": "object",
			"properties": {
				"dataRange": {
					"type": "string",
					"description": "Range of the data to be used in the formula (e.g., 'A1:D10'). If omitted, uses the entire sheet. If the range is larger than 1000 cells, an error will be returned. If the table is too large, choose a meaningful sample of the data to provide context for the formula generation."
				},
				"targetRange": {
					"type": "string",
					"description": "The target cell or range where the generated formula will be placed (e.g., 'B1:B10', 'E1'). If omitted, the formula will be returned without placing it in a cell. If target not empty, error will be returned. If target range is larger than 10 cells, error will be returned."
				},
				"request": {
					"type": "string",
					"description": "The request or instruction for LLM to generate the formula."
				}
			},
			"required": []
		},
		"examples": [
			{
				"prompt": "Generate a formula to calculate the standard deviation of column B",
				"arguments": { "dataRange": "B1:B10", "request": "Generate a formula to calculate the standard deviation of column B." }
			},
			{
				"prompt": "Add a new row at the bottom of the sheet and generate averages for every numeric column to fill it",
				"arguments": { "dataRange": "A1:D10", "targetRange": "A11:E11", "request": "Generate averages for every numeric column to fill the new row." }
			},
			{
				"prompt": "Generate a formula to find the minimum value of column A",
				"arguments": { "dataRange": "A1:D10", "request": "Generate a formula to find the minimum value of column A." }
			},
			{
				"prompt": "Suggest and generate formulas to measure the distribution of numeric values in the entire sheet.",
				"arguments": { "request": "Suggest and generate formulas to measure the distribution of numeric values in the entire sheet." }
			}
		]
	});

	func.call = async function (params) {
		if (params.dataRange !== undefined && typeof params.dataRange !== 'string') {
			throw new window.AgentState.ToolError(
				'Parameter "dataRange" must be a string like "A1:D100". Got: ' + JSON.stringify(params.dataRange)
			);
		}
		if (params.targetRange !== undefined && typeof params.targetRange !== 'string') {
			throw new window.AgentState.ToolError(
				'Parameter "targetRange" must be a string like "A1:D100". Got: ' + JSON.stringify(params.targetRange)
			);
		}
		console.log("generateFormula called with params: ", params);

		Asc.scope.dataRange = params.dataRange;
		Asc.scope.targetRange = params.targetRange;
		Asc.scope.request = params.request;

		let rangeData = await Asc.Editor.callCommand(function () {
			let ws = Api.GetActiveSheet();
			let _range;
			let _targetRange;

			if (!Asc.scope.dataRange) {
				_range = ws.GetUsedRange();
			} else {
				_range = ws.GetRange(Asc.scope.dataRange);
				if (!_range)
					return { error: "Invalid range \"" + Asc.scope.dataRange + "\". Please provide a valid Excel range like 'A1:D10'." };
			}

			if (!_range)
				return null;

			if (_range.GetCount() > 1000) {
				return { error: "Target range \"" + Asc.scope.targetRange + "\" is too large. Please provide a smaller range (max 1000 cells)." };
			}

			console.log("Data range: ", _range.GetAddress(true, true, "xlA1", false).replaceAll("$", ""));

			let targetRangeCells = [];
			if (Asc.scope.targetRange) {
				_targetRange = ws.GetRange(Asc.scope.targetRange);
				if (!_targetRange)
					return { error: "Invalid target range \"" + Asc.scope.targetRange + "\". Please provide a valid Excel range like 'A1:D10'." };
				if (_targetRange.GetCount() > 10) {
					return { error: "Target range \"" + Asc.scope.targetRange + "\" is too large. Please provide a smaller range (max 10 cells)." };
				}
				let targetRangeError = null;
				_targetRange.ForEach(function (cell) {
					let text = cell.GetText();
					if (text && text.trim() !== '') {
						targetRangeError = "Target range \"" + Asc.scope.targetRange + "\" is not empty. Please provide an empty range.";
						return;
					}
					targetRangeCells.push(cell.GetAddress(true, true, "xlA1", false).replaceAll("$", ""));
				});
				if (targetRangeError) {
					return { error: targetRangeError };
				}
			}
			console.log("Target range: ", _targetRange ? _targetRange.GetAddress(true, true, "xlA1", false).replaceAll("$", "") : "None");
			// Use ForEach to iterate through the first 5 rows and 5 columns of data to get a sample of the data for context.
			let sampleData = {};

			let rowCount = Math.min(_range.RowCount, 5);
			let colCount = Math.min(_range.ColCount, 5);

			_range.ForEach(function (cell) {
				// Make sure to only sample the first 5 rows and 5 columns of the range for context.
				if (cell.Row - _range.Row < rowCount || cell.Col - _range.Col < colCount) {
					sampleData[cell.GetAddress(true, true, "xlA1", false).replaceAll("$", "")] = cell.GetText();
				}
			});

			console.log("Sample data: ", sampleData);
			if (Object.keys(sampleData).length === 0)
				return { error: "No data found in the specified range." };

			return {
				sampleData: sampleData,
				targetRangeCells: targetRangeCells
			};
		});

		if (rangeData && rangeData.error)
			throw new window.AgentState.ToolError(rangeData.error);

		if (!rangeData)
			return;
		let responseFormat = `Please generate a valid Excel formula that fulfills the user's request. 
			If the request is ambiguous or cannot be fulfilled, provide a clear explanation instead of a formula.
			Do not include any additional text or explanations in your response. Only provide the formula or the explanation.`

		if (Asc.scope.targetRange) {
			responseFormat = `Please provide the generated formula(s) in a JSON array format, where each key corresponds to a cell in the target range "${rangeData.targetRangeCells.join(", ")}".
			And each value should be a string representing the formula for that cell. example: { "A1": "=SUM(B1:B10)", "A2": "=AVERAGE(C1:C10)" }.
			The order of the formulas should match the order of the cells in the target range, from top-left to bottom-right.
			If a formula cannot be generated for a specific cell, please provide an explanation instead of a formula for that cell.
			Do not include any additional text or explanations in your response. Only provide the formula(s) or the explanation(s) as per the user's request. As the response will be parsed as JSON`;
		}

		let argPromptContext = `
			You are an AI assistant that generates Excel formulas based on user requests. The user has made the following request: ${Asc.scope.request}. 
			Provided is the following sample data from their spreadsheet: ${JSON.stringify(rangeData.sampleData)}.
			Do not make any assumptions about the data beyond what is provided in the sample.
			The response should be in the following format:
			${responseFormat}
		`;
		console.log("argPromptContext: ", argPromptContext);

		let requestEngine = AI.Request.create(AI.ActionType.Chat);
		if (!requestEngine)
			return;

		let isSendedEndLongAction = false;
		async function checkEndAction() {
			if (!isSendedEndLongAction) {
				await Asc.Editor.callMethod("EndAction", ["Block", "AI (" + requestEngine.modelUI.name + ")"]);
				isSendedEndLongAction = true;
			}
		}

		await Asc.Editor.callMethod("StartAction", ["Block", "AI (" + requestEngine.modelUI.name + ")"]);

		try {
			await Asc.Editor.callMethod("StartAction", ["GroupActions"]);
		}
		catch (error) {
			await checkEndAction();
			throw new window.AgentState.ToolError("Failed to start group actions. Error: " + error.message);
		}
		let aiResult;
		try {
			aiResult = await requestEngine.chatRequest(argPromptContext, false, async function (data) {
				if (!data)
					return;
				await checkEndAction();
			});
		}
		catch (error) {
			// If the request fails, we still want to end the action and group actions before throwing the error.
			await checkEndAction();
			await Asc.Editor.callMethod("EndAction", ["GroupActions"]);
			throw new window.AgentState.ToolError("Failed to generate notes for the slide. Error: " + error.message);
		}
		await checkEndAction();
		await Asc.Editor.callMethod("EndAction", ["GroupActions"]);

		if (aiResult) {
			console.log("aiResult: ", aiResult);
			try {
				if (!Asc.scope.targetRange && typeof aiResult === 'string') {
					return aiResult.trim();
				}

				let formulas = JSON.parse(aiResult.trim());
				if (!formulas || typeof formulas !== 'object' || Array.isArray(formulas)) {
					throw new window.AgentState.ToolError("Expected a JSON object keyed by target cell address.");
				}

				Asc.scope.formulas = formulas;

				await Asc.Editor.callCommand(function () {
					let ws = Api.GetActiveSheet();

					for (let cell in Asc.scope.formulas) {
						let formula = Asc.scope.formulas[cell];

						if (formula && typeof formula === 'string') {
							let cellObj = ws.GetRange(cell);
							if (cellObj && cellObj.GetCount() === 1) {
								cellObj.SetValue(formula);
							}
						}
					}
				});
			} catch (error) {
				if (!Asc.scope.targetRange && typeof aiResult === 'string') {
					return aiResult.trim();
				}
				throw new window.AgentState.ToolError("Failed to parse AI response for generated formulas: " + error.message);
			}
		}
	};

	return func;
})();
