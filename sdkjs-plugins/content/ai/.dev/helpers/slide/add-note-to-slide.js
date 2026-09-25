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
		"name": "addNoteToSlide",
		"text": "Insert Note",
		"description": `Adds a note to the slide. This function is particularly useful for adding speaker notes or additional context to a slide to aid in presentations.
		If slide number is not specified, the note will be added to the current slide.
		If intent is passed in the text parameter, precise text is added to the notes.
		If intent is passed in the request parameters, it interpreted as an LLM prompt. 
		If the request is meant to target multiple slides, call this function once for each slide.`,
		"parameters": {
			"type": "object",
			"properties": {
				"slideNumber": {
					"type": "number",
					"description": "the slide number to add text to (optional, default current slide)",
					"minimum": 1
				},
				"text": {
					"type": "string",
					"description": "Precise text to add to the note"
				},
				"request": {
					"type": "string",
					"description": "LLM prompt describing what the user intends to add to the notes"
				}
			},
			"required": []
		},
		"examples": [
			{
				"prompt": "write a note with the following content to slide 3: Hello, world!",
				"arguments": { "slideNumber": 3, "text": "Hello, world!" }
			},
			{
				"prompt": "add talking points to slide 2",
				"arguments": { "slideNumber": 2, "request": "add talking points to slide 2" }
			},
			{
				"prompt": "make a note with AI content",
				"arguments": { "request": "make a note with AI content" }
			},
			{
				"prompt": "Write speaker notes",
				"arguments": { "request": "write speaker notes" }
			},
			{
				"prompt": "create a speaking script.",
				"arguments": { "request": "create a speaking script." }
			},
			{
				"prompt": "generate talking points for this slide",
				"arguments": { "request": "generate talking points for this slide" }
			},
		]
	});

	func.call = async function (params) {
		Asc.scope.params = params;
		if (!Asc.scope.params.text && !Asc.scope.params.request) {
			throw new window.AgentState.ToolError("No text was passed to the addNoteToSlide")
		}
		if (Asc.scope.params.text && Asc.scope.params.request) {
			throw new window.AgentState.ToolError("Only one of 'text' or 'request' should be provided")
		}
		if ('slideNumber' in Asc.scope.params && ( !Number.isInteger(Asc.scope.params.slideNumber) || Asc.scope.params.slideNumber < 1)) {
			throw new window.AgentState.ToolError("Invalid slide number. Please provide a positive integer slide number.");
		}

		// Read, compute and validate parameters
		let callResult = await Asc.Editor.callCommand(function () {

			let presentation = Api.GetPresentation();
			let slide;
			if (Asc.scope.params.slideNumber) {
				slide = presentation.GetSlideByIndex(Asc.scope.params.slideNumber - 1);
				if (!slide) return { error: "slide_not_found", slidesCount: presentation.GetSlidesCount() };
			}
			else {
				slide = presentation.GetCurrentSlide();
			}

			if (!slide) return { error: "no_current_slide" };


			// If text is passed, return early here.
			if (Asc.scope.params.text) {
				if (!slide.AddNotesText(Asc.scope.params.text)) {
					return { error: "failed_to_add_note", text: Asc.scope.params.text, slideNumber: slide.GetSlideIndex() }
				}
				return;

			}

			let slideContent;
			// Fetch slide content for LLM case
			if (Asc.scope.params.request) {

				// Get slide content. Tolerate errors.
				let shapesContent = [];
				let shapes = slide.GetAllShapes();
				for (let i = 0; i < shapes.length; i++) {
					let shape = shapes[i];
					let shapeText = "";
					try {
						let content = shape.GetDocContent();
						if (content) {
							let count = content.GetElementsCount();
							let parts = [];
							for (let j = 0; j < count; j++) {
								let el = content.GetElement(j);
								if (el && el.GetText) {
									parts.push(el.GetText());
								}
							}
							shapeText = parts.join("\n");
						}
					}
					// Tolerate failures reading slide content
					catch (e) {
					}
					if (shapeText) shapesContent.push(shapeText);
				}

				let shapesResult = shapesContent.join("\n\n");

				// Get slide content from tables. Tolerate errors.
				let tableResults = [];
				try {
					let aTables = slide.GetAllTables();
					for (let i = 0; i < aTables.length; i++) {
						let table = aTables[i];
						let rows = [];
						let k = 0;
						let rowObj = table.GetRow(k++);
						while (rowObj) {
							let row = [];
							let nCols = rowObj.GetCellsCount();
							for (let c = 0; c < nCols; c++) {
								let cell = rowObj && rowObj.GetCell ? rowObj.GetCell(c) : null;
								let cellText = cell && cell.GetText ? cell.GetText() : "";
								row.push(cellText);
							}
							rows.push(row);
							rowObj = table.GetRow(k++);
						}
						tableResults.push(rows);
					}
				}
				catch (e) {
				}
				let tableJsonContents = JSON.stringify(tableResults);

				slideContent = "Plain text of the slide: " + shapesResult + "\n\n" + "Contents of tables on the slide: " + tableJsonContents;
			}
			return {
				slideContentObj: slideContent,
				slideNumber: slide.GetSlideIndex()
			};
		});
		if (callResult && callResult.error === "failed_to_add_note") {
			throw new window.AgentState.ToolError("failed to add note. Parametes: Text: " + callResult.text + ", slideNumber: " + callResult.slideNumber);
		}
		if (callResult && callResult.error === "slide_not_found") {
			throw new window.AgentState.ToolError("Slide " + params.slideNumber + " does not exist! The presentation has " + callResult.slidesCount + " slides.");
		}
		if (callResult && callResult.error === "no_current_slide") {
			throw new window.AgentState.ToolError("No current slide available.");
		}

		// If text is passed, we are done. Otherwise, we continue to LLM case.
		if (Asc.scope.params.text) return;

		var text = '';

		// Create LLM request
		let llmPrompt =
			`You are an AI chatbot. You are tasked to generate notes to a specific slide of a presentation. 
					To do that, you should primarily follow the user's request which is: ${Asc.scope.params.request}
					To enrich your output, you should use the slide's content: ${callResult.slideContentObj}
					Note that the slide's content and tables may be empty. 
					Do note make stuff up. If there is not enough context to generate notes, simply return "Not enough content"
					If the request and presentation are not in english try to detect the language and match it in your output. 
					`
		let requestEngine = AI.Request.create(AI.ActionType.Chat);
		if (!requestEngine)
			throw new window.AgentState.ToolError("Request engine is not available for the action type: " + AI.ActionType.Chat);

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

		try {
			text = await requestEngine.chatRequest(llmPrompt, false, async function (data) {
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

		Asc.scope.addNotesResult = text;
		Asc.scope.addNotesResolvedSlideNumber = callResult.slideNumber;
		callResult = await Asc.Editor.callCommand(function () {
			// Push result to notes
			let text = Asc.scope.addNotesResult;
			let slideNumber = Asc.scope.addNotesResolvedSlideNumber;
			let presentation = Api.GetPresentation();
			let slide;
			slide = presentation.GetSlideByIndex(slideNumber);
			if (!slide) return { error: "slide_not_found", slidesCount: presentation.GetSlidesCount() };

			if (!slide.AddNotesText(text)) {
				return { error: "failed_to_add_note", text: text, slideNumber: slide.GetSlideIndex() }
			}
		})

		if (callResult && callResult.error === "failed_to_add_note") {
			throw new window.AgentState.ToolError("failed to add note. Parametes: Text: " + callResult.text + ", slideNumber: " + callResult.slideNumber);
		}
		if (callResult && callResult.error === "slide_not_found") {
			throw new window.AgentState.ToolError("Slide " + Asc.scope.addNotesResolvedSlideNumber + " does not exist. It was removed from the presentation before the operation could be completed. The presentation has " + callResult.slidesCount + " slides.");
		}
	};

	return func;
})();
