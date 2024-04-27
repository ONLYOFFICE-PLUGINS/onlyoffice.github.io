/**
 *
 * (c) Copyright Ascensio System SIA 2020
 *
 * Licensed under the Apache License, Version 2.0 (the 'License');
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *	 http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an 'AS IS' BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */
 // todo поправить readme  и собранный плагин и с синонимами поправить тоже
(function(window, undefined){
	const plugin = window.Asc.plugin;
	const localStorageKey = 'OpenAISettings';
	let pluginSettings = {
		apiKey: '',
		model: 'gpt-3.5-turbo',
		maxTokens: 4000
	};

	let loadingPhrase = 'Loading...';
	let thesaurusCounter = 0;
	let pluginWindows = {
		settingsWindow : null,
		chatWindow : null,
		customReqWindow : null,
		translatorWindow : null,
		summarizationWindow : null,
		textToImageWindow : null,
		keyWordsWindow : null,
		oleHistoryWindow : null,
		linkWindow: null
	};
	let imgsize = null;
	let link = null;
	let oleData = null;
	let bCreateMenu = true;

	plugin.init = function() {
		createToolbarMenu();
	};

	function checkApiKey() {
		try {
			let value = localStorage.getItem(localStorageKey);
			if (value)
				pluginSettings = JSON.parse(value);
		} catch (error) {
			// we have some problem with saved message. we should remove it from localStorage
			localStorage.removeItem(localStorageKey);
		}
	};

	function createToolbarMenu() {
		const items = getPannelWithItems(plugin.info.editorType);
		if (items)
			plugin.executeMethod('AddToolbarMenuItem', [items]);
	};

	function getPannelWithItems(editorType) {
		checkApiKey();
		let items = {guid: plugin.guid, tabs:[]};
		if (pluginSettings.apiKey) {
			items = {
				guid : plugin.guid,
				tabs: [
					{
						id : "open_ai_plugin_tab",
						text: 'AI',
						items: [
							{
								id: 'ai_test',
								type: 'big-button',
								text:  'Test item',
								hint: 'test hint',
								data: 'TEST_MAIN',
								lockInViewMode: false,
								icons: 'resources/%theme-type%(light)/%state%(normal)translate%scale%(100%).%extension%(png)',
								items : [
									{ 
										id : "toolC1", 
										text : "Explain text in comment", 
										data : "Hello",
										icons: 'resources/dark/icon.png'
									},
									{
										id : "toolC2",
										text : "Fix spelling & grammar",
										data: "TEST_HELLO",
										icons: 'resources/%theme-type%(light)/%state%(normal)translate%scale%(100%).%extension%(png)'
									},
									{ id : "toolC3", text : "Make longer" },
									{ id : "toolC4", text : "Make shorter" }
								]
							},
							{
								id: 'ai_setting',
								type: 'big-button',
								text:  'Settings',
								hint: 'Enter an Api key for OpenAI.',
								icons: 'resources/%theme-type%(light)/%state%(normal)translate%scale%(100%).%extension%(png)',
								lockInViewMode: false,
								enableToggle: true
							},
							{
								id: 'ai_translator',
								type: 'big-button',
								text:  'AI tranlator',
								hint: 'Translate your text to selected language.',
								icons: 'resources/%theme-type%(light)/%state%(normal)translate%scale%(100%).%extension%(png)',
								lockInViewMode: false,
								enableToggle: true,
								separator: true,
							},
							{
								id: 'ai_check_text',
								type: 'button',   
								text:  "Check text",
								icons: 'resources/%theme-type%(light)/%state%(normal)checktext%scale%(100%).%extension%(png)',
								hint: "Check text",
								lockInViewMode: true,
								enableToggle: true
							},
							{
								id: 'ai_ask_ai',
								type: 'button',   
								text:  "Ask AI",
								icons: 'resources/%theme-type%(light)/%state%(normal)askai%scale%(100%).%extension%(png)',
								hint: "Ask AI",
								lockInViewMode: false,
								enableToggle: true
							},
							{
								id: 'ai_text_to_image',
								type: 'button',
								text:  "Text to image",
								icons: 'resources/%theme-type%(light)/%state%(normal)texttoimage%scale%(100%).%extension%(png)',
								hint: "Text to image",
								separator: true,
								lockInViewMode: true,
								enableToggle: true
							},
							{
								id: 'ai_image_to_image',
								type: 'button',
								text:  "Image to image",
								icons: 'resources/%theme-type%(light)/%state%(normal)imagetoimage%scale%(100%).%extension%(png)',
								hint: "Image to image",
								lockInViewMode: true,
								enableToggle: true
							}
						]
					}
				]
			};

			switch (editorType) {
				case 'word':
					items.tabs[0].items.push(
						{
							id: 'ai_summarization',
							type: 'button',
							text:  "Summatization",
							icons: 'resources/%theme-type%(light)/%state%(normal)summarization%scale%(100%).%extension%(png)',
							hint: "Summatization",
							lockInViewMode: true,
							separator: true,
							enableToggle: true
						},
						{
							id: 'ai_beautify_text',
							type: 'button',
							text:  "Beautify text",
							icons: 'resources/%theme-type%(light)/%state%(normal)beautifytext%scale%(100%).%extension%(png)',
							hint: "Beautify text",
							lockInViewMode: true,
							enableToggle: true
						},
						{
							id: 'ai_key_words',
							type: 'button',
							text:  "Key words",
							icons: 'resources/%theme-type%(light)/%state%(normal)keywords%scale%(100%).%extension%(png)',
							hint: "Key words",
							separator: true,
							lockInViewMode: true,
							enableToggle: true
						},
						{
							id: 'ai_heading_generator',
							type: 'button',
							text:  "Heading Generator",
							icons: 'resources/%theme-type%(light)/%state%(normal)headinggenerator%scale%(100%).%extension%(png)',
							hint: "Heading Generator",
							lockInViewMode: true,
							enableToggle: true
						}
					);	
					break;
				case 'slide':
					items.tabs[0].items.push(
						{
							id: 'ai_summarization',
							type: 'button',
							text:  "Summatization",
							icons: 'resources/%theme-type%(light)/%state%(normal)summarization%scale%(100%).%extension%(png)',
							hint: "Summatization",
							lockInViewMode: true,
							enableToggle: true
						},
						{
							id: 'ai_beautify_text',
							type: 'button',
							text:  "Beautify text",
							icons: 'resources/%theme-type%(light)/%state%(normal)beautifytext%scale%(100%).%extension%(png)',
							hint: "Beautify text",
							lockInViewMode: true,
							enableToggle: true
						},
						{
							id: 'ai_heading_generator',
							type: 'button',
							text:  "Heading Generator",
							icons: 'resources/%theme-type%(light)/%state%(normal)headinggenerator%scale%(100%).%extension%(png)',
							hint: "Heading Generator",
							lockInViewMode: true,
							separator: true,
							enableToggle: true
						}
					);
					break;
				case 'cell':
					
					break;
			
			}
			// start test block
			plugin.attachToolbarMenuClickEvent("ai_test", function(data) {
				console.log("Test item: " + data);
			});
			plugin.attachToolbarMenuClickEvent("toolC1", function(data) {
				console.log("Explain text in comment: " + data);
			});
			// end test block

			
			plugin.attachToolbarMenuClickEvent("ai_setting", createSettingsWindow);
			plugin.attachToolbarMenuClickEvent("ai_translator", createTranslatorWindow);
			plugin.attachToolbarMenuClickEvent("ai_check_text", function() {
				// createSpellingWindow();
			});
			plugin.attachToolbarMenuClickEvent("ai_ask_ai", function() {
				console.log("Ask AI");
			});
			plugin.attachToolbarMenuClickEvent("ai_summarization", makeSummarization);
			plugin.attachToolbarMenuClickEvent("ai_beautify_text", function() {
				console.log("Beautify text");
			});
			plugin.attachToolbarMenuClickEvent("ai_key_words", createKeyWordsWindow);
			plugin.attachToolbarMenuClickEvent("ai_heading_generator", function() {
				console.log("Heading Generator");
			});
			plugin.attachToolbarMenuClickEvent("ai_text_to_image", createTextToImageWindow);
			plugin.attachToolbarMenuClickEvent("ai_image_to_image", function() {
				console.log("Image to image");
			});
		}
		return items;
	};

	function getContextMenuItems(options) {
		link = null;
		checkApiKey();
		let settings = {
			guid: plugin.info.guid,
			items: [
				{
					id : 'ChatGPT',
					text : generateText('ChatGPTnew'),
					items : []
				}
			]
		};

		if (pluginSettings.apiKey) {
			switch (options.type) {
				case 'Target': {
					if (Asc.plugin.info.editorType === 'word') {
						settings.items[0].items.push({
							id : 'onMeaningT',
							text : generateText('Explain text in comment')
						});
					}

					break;
				}
				case 'Selection': {
					if (Asc.plugin.info.editorType === 'word') {
						settings.items[0].items.push(
							{
								id : 'onFixSpelling',
								text : generateText('Fix spelling & grammar')
							},
							{
								id : 'onRewrite',
								text : generateText('Rewrite differently')
							},
							{
								id : 'onMakeLonger',
								text : generateText('Make longer')
							},
							{
								id : 'onMakeShorter',
								text : generateText('Make shorter')
							},
							{
								id : 'onMakeSimple',
								text : generateText('Make simpler')
							},
							{
								id : 'TextAnalysis',
								text : generateText('Text analysis'),
								separator: true,
								items : [
									{
										id : 'onSummarize',
										text : generateText('Summarize')
									},
									{
										id : 'onKeyWords',
										text : generateText('Keywords')
									},
									{
										id : 'onHeading',
										text : generateText('Heading generator')
									},
								]
							},
							{
								id : 'Tex Meaning',
								text : generateText('Word analysis'),
								items : [
									{
										id : 'onMeaningS',
										text : generateText('Explain text in comment'),
									},
									{
										id : 'onMeaningLinkS',
										text : generateText('Explain text in hyperlink')
									}
								]
							},
							{
								id : 'TranslateText',
								text : generateText('Translate'),
								items : [
									{
										id : 'onTranslate',
										text : generateText('Translate to English'),
										data : 'English'
									},
									{
										id : 'onTranslate',
										text : generateText('Translate to French'),
										data : 'French'
									},
									{
										id : 'onTranslate',
										text : generateText('Translate to German'),
										data : 'German'
									},
									{
										id : 'onTranslate',
										text : generateText('Translate to Chinese'),
										data : 'Chinise'
									},
									{
										id : 'onTranslate',
										text : generateText('Translate to Japanese'),
										data : 'Japanese'
									},
									{
										id : 'onTranslate',
										text : generateText('Translate to Russian'),
										data : 'Russian'
									},
									{
										id : 'onTranslate',
										text : generateText('Translate to Korean'),
										data : 'Korean'
									},
									{
										id : 'onTranslate',
										text : generateText('Translate to Spanish'),
										data : 'Spanish'
									},
									{
										id : 'onTranslate',
										text : generateText('Translate to Italian'),
										data : 'Italian'
									},
								]
							},
							{
								id : 'OnGenerateImageList',
								text : generateText('Generate image from text'),
								items : [
									{
										id : 'OnGenerateImage',
										text : generateText('256x256'),
										data : 256
									},
									{
										id : 'OnGenerateImage',
										text : generateText('512x512'),
										data : 512
									},
									{
										id : 'OnGenerateImage',
										text : generateText('1024x1024'),
										data : 1024
									}
								]
							}
						);
					}
					break;
				}
				case 'Image':
				case 'Shape': {
					settings.items[0].items.push({
						id : 'onImgVar',
						text : generateText('Generate image variation')
					});

					break;
				}
				case 'OleObject': {
					plugin.executeMethod('GetSelectedOleObjects', null, function(arr) {
						oleData = null;
						let oleArr = arr.filter(function(obj) {
							return obj.guid === plugin.info.guid;
						});
						if (oleArr.length) {
							oleData = oleArr[0];
							settings.items[0].items.unshift({
								id : 'onOleHistory',
								text : generateText('Show history')
							});
						}
					});
					break;
				}
				case 'Hyperlink':{
					settings.items[0].items.push({
						id : 'onHyperlink',
						text : generateText('Show hyperlink content')
					});
					link = options.value;
					break;
				}

				default:
					break;
			}

			// Add it only if this window is closed
			if (Asc.plugin.info.editorType === 'word' && options.type === 'Selection' && !pluginWindows.translatorWindow) {
				settings.items[0].items[7].items.push({
					id : 'onTranslate',
					text : generateText('Open advanced settings'),
					separator: true,
					data : 'Settings'
				});
			}

			settings.items[0].items.push(
				{
					id : 'onChat',
					text : generateText('Chat'),
					separator: true
				},
				{
					id : 'onCustomReq',
					text : generateText('Custom request')
				}
			);
		}

		settings.items[0].items.push({
			id : 'onSettings',
			text : generateText('Settings'),
			separator: true
		});

		return settings;
	};

	plugin.attachEvent('onContextMenuShow', function(options) {
		// todo: change key validation
		if (!options)
			return;

		this.executeMethod('AddContextMenuItem', [getContextMenuItems(options)]);

		if (pluginSettings.apiKey && options.type === "Target")
		{
			plugin.executeMethod('GetCurrentWord', null, function(text) {
				if (!isEmpyText(text, true)) {
					thesaurusCounter++;
					let tokens = window.Asc.OpenAIEncode(text);
					createSettings(text, tokens, 9, true);
				}
			});
		}
	});

	function generateText(text) {
		let lang = plugin.info.lang.substring(0,2);
		let result = { en: text	};
		if (lang !== "en")
			result[lang] = plugin.tr(text);

		return result;
	};

	plugin.attachContextMenuClickEvent('onSettings', createSettingsWindow);

	plugin.attachContextMenuClickEvent('onCustomReq', function() {
		// default settings for modal window (I created separate settings, because we have many unnecessary field in plugin variations)
		let variation = {
			url : makeFileUrl('custom.html'),
			description : plugin.tr('OpenAI'),
			isVisual : true,
			isModal : true,
			EditorsSupport : ["word", "slide", "cell"],
			size : [ 343, 434 ],
			buttons : [
				{
					text : "Submit",
					primary : true,
					textLocale : {
						"ru"    : "Отправить",
						"fr"    : "Envoyer",
						"es"    : "Enviar",
						"pt-BR" : "Enviar",
						"de"    : "Senden",
						"cs"    : "Odeslat",
						"zh"    : "提交"
					}
				},
				{
					text : "Clear",
					textLocale : {
						"ru"    : "Очистить",
						"fr"    : "Nettoyer",
						"es"    : "Limpiar",
						"pt-BR" : "Claro",
						"de"    : "Reinigen",
						"cs"    : "Vyčistit",
						"zh"    : "清除"
					}
				}
			]
		};

		if (!pluginWindows.customReqWindow)
			pluginWindows.customReqWindow = createWindow();

		pluginWindows.customReqWindow.show(variation);
	});

	plugin.attachContextMenuClickEvent('onChat', function() {
		// default settings for modal window (I created separate settings, because we have many unnecessary field in plugin variations)
		let variation = {
			url : makeFileUrl('chat.html'),
			description : plugin.tr('ChatGPT'),
			isVisual : true,
			buttons : [],
			isModal : false,
			EditorsSupport : ["word", "slide", "cell"],
			size : [ 343, 486 ]
		};

		if (!pluginWindows.chatWindow)
			pluginWindows.chatWindow = createWindow();

		pluginWindows.chatWindow.show(variation);
	});

	plugin.attachContextMenuClickEvent('onHyperlink', function(data) {
		// default settings for modal window (I created separate settings, because we have many unnecessary field in plugin variations)
		let variation = {
			url : makeFileUrl('hyperlink.html'),
			description : plugin.tr('Hyperlink'),
			isVisual : true,
			buttons : [],
			isModal : false,
			EditorsSupport : ["word"],
			size : [ 1000, 1000 ]
		};

		if (!pluginWindows.linkWindow)
			pluginWindows.linkWindow = createWindow();

		pluginWindows.linkWindow.show(variation);
		// setTimeout(function() {
		// 	pluginWindows.linkWindow.command('onTest', link);
		// },500)
	});

	plugin.attachContextMenuClickEvent('onOleHistory', creatOleVersionWindow);

	plugin.attachContextMenuClickEvent('onMeaningT', function() {
		plugin.executeMethod('GetCurrentWord', null, function(text) {
			if (!isEmpyText(text)) {
				let tokens = window.Asc.OpenAIEncode(text);
				createSettings(text, tokens, 8);
			}
		});
	});

	plugin.attachContextMenuClickEvent('onSummarize', makeSummarization);

	plugin.attachContextMenuClickEvent('onKeyWords', function() {
		plugin.executeMethod('GetSelectedText', null, function(text) {
			if (!isEmpyText(text)) {
				let tokens = window.Asc.OpenAIEncode(text);
				createSettings(text, tokens, 2);
			}
		});
	});

	plugin.attachContextMenuClickEvent('onMeaningS', function() {
		plugin.executeMethod('GetSelectedText', null, function(text) {
			if (!isEmpyText(text)) {
				let tokens = window.Asc.OpenAIEncode(text);
				createSettings(text, tokens, 3);
			}
		});
	});

	plugin.attachContextMenuClickEvent('onMeaningLinkS', function() {
		plugin.executeMethod('GetSelectedText', null, function(text) {
			if (!isEmpyText(text)) {
				let tokens = window.Asc.OpenAIEncode(text);
				createSettings(text, tokens, 4);
			}
		});
	});

	plugin.attachContextMenuClickEvent('onTranslate', function(data) {
		if (data === 'Settings') {
			if (!pluginWindows.translatorWindow)
				createTranslatorWindow();
		} else {
			plugin.executeMethod('GetSelectedText', null, function(text) {
				if (!isEmpyText(text)) {
					let tokens = window.Asc.OpenAIEncode(text);
					let message = `Translate to ${data}: ${text}`;
					createSettings(message, tokens, 6);
				}
			});
		}
	});

	plugin.attachContextMenuClickEvent('OnGenerateImage', function(data) {
		let size = Number(data);
		imgsize = {width: size, height: size};
		plugin.executeMethod('GetSelectedText', null, function(text) {
			if (!isEmpyText(text)) {
				let tokens = window.Asc.OpenAIEncode(text);
				createSettings(text, tokens, 7);
			}
		});
	});

	plugin.attachContextMenuClickEvent('onThesaurus', function(data) {
		plugin.executeMethod('ReplaceCurrentWord', [data]);
	});

	plugin.attachContextMenuClickEvent('onImgVar', function() {
		plugin.executeMethod('GetImageDataFromSelection', null, function(data) {
			createSettings(data, 0, 10);
		});
	});

	plugin.attachContextMenuClickEvent('onFixSpelling', function() {
		plugin.executeMethod('GetSelectedText', null, function(text) {
			if (!isEmpyText(text)) {
				let tokens = window.Asc.OpenAIEncode(text);
				createSettings(text, tokens, 11);
			}
		});
	});

	plugin.attachContextMenuClickEvent('onRewrite', function() {
		plugin.executeMethod('GetSelectedText', null, function(text) {
			if (!isEmpyText(text)) {
				let tokens = window.Asc.OpenAIEncode(text);
				createSettings(text, tokens, 12);
			}
		});
	});

	plugin.attachContextMenuClickEvent('onMakeLonger', function() {
		plugin.executeMethod('GetSelectedText', null, function(text) {
			if (!isEmpyText(text)) {
				let tokens = window.Asc.OpenAIEncode(text);
				createSettings(text, tokens, 13);
			}
		});
	});

	plugin.attachContextMenuClickEvent('onMakeShorter', function() {
		plugin.executeMethod('GetSelectedText', null, function(text) {
			if (!isEmpyText(text)) {
				let tokens = window.Asc.OpenAIEncode(text);
				createSettings(text, tokens, 14);
			}
		});
	});

	plugin.attachContextMenuClickEvent('onMakeSimple', function() {
		plugin.executeMethod('GetSelectedText', null, function(text) {
			if (!isEmpyText(text)) {
				let tokens = window.Asc.OpenAIEncode(text);
				createSettings(text, tokens, 15);
			}
		});
	});

	plugin.attachContextMenuClickEvent('onHeading', function() {
		plugin.executeMethod('GetSelectedText', null, function(text) {
			if (!isEmpyText(text)) {
				let tokens = window.Asc.OpenAIEncode(text);
				createSettings(text, tokens, 16);
			}
		});
	});

	function createSettings(text, tokens, type, isNoBlockedAction) {
		let url;
		let settings = {
			model : pluginSettings.model,
			max_tokens : pluginSettings.maxTokens - tokens.length
		};

		if (settings.max_tokens < 100) {
			// todo add visual error
			console.error(new Error('This request is too big!'));
			return;
		}

		plugin.executeMethod('StartAction', [isNoBlockedAction ? 'Information' : 'Block', 'ChatGPT: ' + loadingPhrase]);

		switch (type) {
			case 1:
				settings.messages = [ { role: 'user', content: `Summarize this text: "${text}"` } ];
				url = 'https://api.openai.com/v1/chat/completions';
				break;

			case 2:
				settings.messages = [ { role: 'user', content: `Get Key words from this text: "${text}"` } ];
				url = 'https://api.openai.com/v1/chat/completions';
				break;

			case 3:
				settings.messages = [ { role: 'user', content: `What does it mean "${text}"?` } ];
				url = 'https://api.openai.com/v1/chat/completions';
				break;

			case 4:
				settings.messages = [ { role: 'user', content: `Give a link to the explanation of the word "${text}"` } ];
				url = 'https://api.openai.com/v1/chat/completions';
				break;

			case 5:
				settings.messages = [ { role: 'user', content: text } ];
				url = 'https://api.openai.com/v1/chat/completions';
				break;

			case 6:
				settings.messages = [ { role: 'user', content: text } ];
				url = 'https://api.openai.com/v1/chat/completions';
				break;

			case 7:
				delete settings.model;
				delete settings.max_tokens;
				settings.prompt = `Generate image:"${text}"`;
				settings.n = 1;
				settings.size = `${imgsize.width}x${imgsize.height}`;
				settings.response_format = 'b64_json';
				url = 'https://api.openai.com/v1/images/generations';
				break;

			case 8:
				settings.messages = [ { role: 'user', content: `What does it mean "${text}"?` } ];
				url = 'https://api.openai.com/v1/chat/completions';
				break;

			case 9:
				settings.messages = [ { role: 'user', content: `Give synonyms for the word "${text}" as javascript array` } ];
				url = 'https://api.openai.com/v1/chat/completions';
				break;

			case 10:
				imageToBlob(text).then(function(obj) {
					url = 'https://api.openai.com/v1/images/edits';
					url = 'https://api.openai.com/v1/images/variations';
					const formdata = new FormData();
					// formdata.append('prompt', 'add a big red cat');
					formdata.append('image', obj.blob);
					formdata.append('size', obj.size.str);
					formdata.append('n', 1);// Number.parseInt(elements.inpTopSl.value));
					formdata.append('response_format', "b64_json");
					fetchData(formdata, url, type, isNoBlockedAction);
				});
				break;

			case 11:
				settings.messages = [ { role: 'user', content: `Сorrect the errors in this text: ${text}`} ];
				url = 'https://api.openai.com/v1/chat/completions';
				break;

			case 12:
				settings.messages = [ { role: 'user', content: `Rewrite differently and give result on the same language: ${text}`} ];
				url = 'https://api.openai.com/v1/chat/completions';
				break;
			
			case 13:
				settings.messages = [ { role: 'user', content: `Make this text longer and give result on the same language: ${text}`} ];
				url = 'https://api.openai.com/v1/chat/completions';
				break;

			case 14:
				settings.messages = [ { role: 'user', content: `Make this text simpler and give result on the same language: ${text}`} ];
				url = 'https://api.openai.com/v1/chat/completions';
				break;

			case 15:
				settings.messages = [ { role: 'user', content: `Make this text shorter and save language: ${text}`} ];
				url = 'https://api.openai.com/v1/chat/completions';
				break;
			case 16:
				settings.messages = [ { role: 'user', content: `Create a heading for this text: ${text}`} ];
				url = 'https://api.openai.com/v1/chat/completions';
				break;
		}
		if (type !== 10)
			fetchData(settings, url, type, isNoBlockedAction);
	};

	function fetchData(settings, url, type, isNoBlockedAction) {
		let header = {
			'Authorization': 'Bearer ' + pluginSettings.apiKey
		};
		if (type !== 10) {
			header['Content-Type'] = 'application/json';
		}
		fetch(url, {
				method: 'POST',
				headers: header,
				body: (type !== 10 ? JSON.stringify(settings) : settings),
			})
			.then(function(response) {
				return response.json()
			})
			.then(function(data) {
				if (data.error)
					throw data.error

				processResult(data, type, isNoBlockedAction);
			})
			.catch(function(error) {
				if (type == 9)
					thesaurusCounter--;

				// todo add visual error
				console.error(error);
				plugin.executeMethod('EndAction', [isNoBlockedAction ? 'Information' : 'Block', 'ChatGPT: ' + loadingPhrase]);
			});
	};

	function processResult(data, type, isNoBlockedAction) {
		plugin.executeMethod('EndAction', [isNoBlockedAction ? 'Information' : 'Block', 'ChatGPT: ' + loadingPhrase]);
		let text, start, end, img;
		Asc.scope = {};
		switch (type) {
			case 1:
				Asc.scope.summarizationResult = data.choices[0].message.content.split('\n\n'); //data.choices[0].text.split('\n\n');
				Asc.scope.summarizationTitle = plugin.tr('Summarize selected text:') + ' ';
				createSummarizationWindow();
				break;

			case 2:
				Asc.scope.data = data.choices[0].message.content.split('\n\n'); //data.choices[0].text.split('\n\n');
				plugin.callCommand(function() {
					let oDocument = Api.GetDocument();
					for(let ind = 0; ind < Asc.scope.data.length; ind++) {
						let text = Asc.scope.data[ind];
						if (text.length) {
							let oParagraph = Api.CreateParagraph();
							oParagraph.AddText(text);
							oDocument.Push(oParagraph);
						}
					}
				}, false);
				break;

			case 3:
				text = data.choices[0].message.content; //data.choices[0].text;
				Asc.scope.comment = text.startsWith('\n\n') ? text.substring(2) : text;
				plugin.callCommand(function() {
					let oDocument = Api.GetDocument();
					let oRange = oDocument.GetRangeBySelect();
					oRange.AddComment(Asc.scope.comment, 'OpenAI');
				}, false);
				break;

			case 4:
				text = data.choices[0].message.content; //data.choices[0].text;
				start = text.indexOf('htt');
				end = text.indexOf(' ', start);
				if (end == -1) {
					end = text.length;
				}
				Asc.scope.link = text.slice(start, end);
				if (Asc.scope.link) {
					plugin.callCommand(function() {
						let oDocument = Api.GetDocument();
						let oRange = oDocument.GetRangeBySelect();
						oRange.AddHyperlink(Asc.scope.link, 'Meaning of the word');
					}, false);
				}
				break;

			case 5:
				text = data.choices[0].message.content; //data.choices[0].text;
				start = text.indexOf('<img');
				end = text.indexOf('/>', start);
				if (end == -1) {
					end = text.length;
				}
				let imgUrl = text.slice(start, end);
				if (imgUrl) {
					plugin.executeMethod('PasteHtml', [imgUrl])
				}
				break;

			case 6:
				text = data.choices[0].message.content.startsWith('\n\n') ? data.choices[0].message.content.substring(2) : data.choices[0].message.content;
				// text = data.choices[0].text.startsWith('\n\n') ? data.choices[0].text.substring(2) : data.choices[0].text;
				plugin.executeMethod('PasteText', [text]);
				break;

			case 7:
				let url = (data.data && data.data[0]) ? data.data[0].b64_json : null;
				if (url) {
					Asc.scope.url = /^data\:image\/png\;base64/.test(url) ? url : 'data:image/png;base64,' + url + '';
					Asc.scope.imgsize = imgsize;
					imgsize = null;
					plugin.callCommand(function() {
						let oDocument = Api.GetDocument();
						let oParagraph = Api.CreateParagraph();
						let width = Asc.scope.imgsize.width * (25.4 / 96.0) * 36000;
						let height = Asc.scope.imgsize.height * (25.4 / 96.0) * 36000;
						let oDrawing = Api.CreateImage(Asc.scope.url, width, height);
						oParagraph.AddDrawing(oDrawing);
						oDocument.Push(oParagraph);
					}, false);

					// let oImageData = {
					// 	"src": /^data\:image\/png\;base64/.test(url) ? url : `data:image/png;base64,${url}`,
					// 	"width": imgsize.width,
					// 	"height": imgsize.height
					// };
					// imgsize = null;
					// plugin.executeMethod ("PutImageDataToSelection", [oImageData]);
				}
				break;

			case 8:
				text = data.choices[0].message.content; //data.choices[0].text;
				Asc.scope.comment = text.startsWith('\n\n') ? text.substring(2) : text;
				plugin.callCommand(function() {
					var oDocument = Api.GetDocument();
					Api.AddComment(oDocument, Asc.scope.comment, 'OpenAI');
				}, false);
				break;

			case 9:
				thesaurusCounter--;
				if (0 < thesaurusCounter)
					return;

				text = data.choices[0].message.content; //data.choices[0].text;
				let startPos = text.indexOf("[");
				let endPos = text.indexOf("]");

				if (-1 === startPos || -1 === endPos || startPos > endPos)
					return;

				text = text.substring(startPos, endPos + 1);
				let arrayWords = eval(text);

				let items = getContextMenuItems({ type : "Target" });

				let itemNew = {
					id : "onThesaurusList",
					text : generateText("Thesaurus"),
					items : []
				};

				for (let i = 0; i < arrayWords.length; i++)
				{
					itemNew.items.push({
							id : 'onThesaurus',
							data : arrayWords[i],
							text : arrayWords[i]
						}
					);
				}

				items.items[0].items.unshift(itemNew);
				plugin.executeMethod('UpdateContextMenuItem', [items]);
				break;

			case 10:
				img = (data.data && data.data[0]) ? data.data[0].b64_json : null;
				if (img) {
					let sImageSrc = /^data\:image\/png\;base64/.test(img) ? img : 'data:image/png;base64,' + img + '';
					let oImageData = {
						"src": sImageSrc,
						"width": imgsize.width,
						"height": imgsize.height
					};
					imgsize = null;
					plugin.executeMethod("PutImageDataToSelection", [oImageData]);
				}
				break;

			case 11:
				text = data.choices[0].message.content.split('\n\n'); //data.choices[0].text.split('\n\n');
				if (text !== 'The text is correct, there are no errors in it.')
					plugin.executeMethod('ReplaceTextSmart', [text]);
				else
					console.log('The text is correct, there are no errors in it.');
				break;

			case 12:
			case 13:
			case 14:
			case 15:
				text = data.choices[0].message.content.replace(/\n\n/g, '\n'); //data.choices[0].text.split('\n\n');
				plugin.executeMethod('PasteText', [text]);
				break;
			case 16:
				Asc.scope.data = data.choices[0].message.content.split('\n\n'); //data.choices[0].text.split('\n\n');
				plugin.callCommand(function() {
					let oDocument = Api.GetDocument();
					var oNewDocumentStyle = oDocument.GetStyle("Heading 4");
					var oParagraph = Api.CreateParagraph();
					oParagraph.SetStyle(oNewDocumentStyle);
					oParagraph.AddText(Asc.scope.data[0]);
					let range = oDocument.GetRangeBySelect();
					let firstPos = range.Paragraphs[0].GetPosInParent();
					oDocument.AddElement(firstPos, oParagraph);
				}, false);
				break;
		}
	};

	plugin.button = function(id, windowId) {
		if (pluginWindows.customReqWindow && pluginWindows.customReqWindow.id === windowId) {
			switch (id) {
				case 0:
					pluginWindows.customReqWindow.command('onSubmitBtn');
					break;
				case 1:
					pluginWindows.customReqWindow.command('onClearBtn');
					break;
			
				default:
					pluginWindows.customReqWindow.close();
					pluginWindows.customReqWindow = null;
					break;
			}
		} else if (pluginWindows.settingsWindow && pluginWindows.settingsWindow.id === windowId) {
			switch (id) {
				case 0:
					pluginWindows.settingsWindow.command('onSaveBtn');
					break;
			
				default:
					pluginWindows.settingsWindow.close();
					pluginWindows.settingsWindow = null;
					break;
			}
		} else if (pluginWindows.translatorWindow && pluginWindows.translatorWindow.id === windowId) {
			pluginWindows.translatorWindow.command('onSaveSetting');
			pluginWindows.translatorWindow.close();
			pluginWindows.translatorWindow = null;
		} else if (pluginWindows.summarizationWindow && pluginWindows.summarizationWindow.id === windowId) {
			pluginWindows.summarizationWindow.close();
			pluginWindows.summarizationWindow = null;
		} else if (pluginWindows.textToImageWindow && pluginWindows.textToImageWindow.id === windowId) {
			pluginWindows.textToImageWindow.command('onSaveSetting');
			pluginWindows.textToImageWindow.close();
			pluginWindows.textToImageWindow = null;
		} else if (pluginWindows.keyWordsWindow && pluginWindows.keyWordsWindow.id === windowId) {
			pluginWindows.keyWordsWindow.close();
			pluginWindows.keyWordsWindow = null;
		} else if (pluginWindows.oleHistoryWindow && pluginWindows.oleHistoryWindow.id === windowId) {
			switch (id) {
				case 0:
					pluginWindows.oleHistoryWindow.command('onRestore');
					break;
			
				default:
					pluginWindows.oleHistoryWindow.close();
					pluginWindows.oleHistoryWindow = null;
					break;
			}
		} else if (windowId) {
			switch (id) {
				case -1:
				default:
					plugin.executeMethod('CloseWindow', [windowId]);
			}
		}
	};

	plugin.onTranslate = function() {
		loadingPhrase = plugin.tr(loadingPhrase);
	};

	function imageToBlob(img) {
		return new Promise(function(resolve) {
			const image = new Image();
			image.onload = function() {
				const img_size = {width: image.width, height: image.height};
				const canvas_size = normalizeImageSize(img_size);
				const draw_size = canvas_size.width > image.width ? img_size : canvas_size;
				let canvas = document.createElement('canvas');
				canvas.width = canvas_size.width;
				canvas.height = canvas_size.height;
				canvas.getContext('2d').drawImage(image, 0, 0, draw_size.width, draw_size.height*image.height/image.width);
				imgsize = img_size;
				canvas.toBlob(function(blob) {resolve({blob: blob, size: canvas_size})}, 'image/png');
			};
			image.src = img.src;
		});
	};

	function normalizeImageSize (size) {
		let width = 0, height = 0;
		if ( size.width > 750 || size.height > 750 )
			width = height = 1024;
		else if ( size.width > 375 || size.height > 350 )
			width = height = 512;
		else width = height = 256;

		return {width: width, height: height, str: width + 'x' + height}
	};

	function messageHandler(modal, message) {
		switch (message.type) {
			case 'onWindowReady':
				modal.command('onApiKey', JSON.stringify(pluginSettings));
				break;

			case 'onBtnReview':
				pasteSummarization(true);
				break;

			case 'onBtnBelow':
				pasteSummarization();
				break;

			case 'onGetLink':
				modal.command('onSetLink', link);
				break;
			
			case 'onGetOleData':
				modal.command('onData', JSON.stringify(oleData));
				break;

			case 'onRestoreOle':
				oleData = JSON.parse(message.data)
				updateOleObject(message.active)
				break;

			case 'onRemoveApiKey':
				localStorage.removeItem(localStorageKey);
				if (pluginSettings.apiKey)
					bCreateMenu = false;
				break;

			case 'onRemoveApiKeyAndCLose':
				localStorage.removeItem(localStorageKey);
				pluginSettings.apiKey = '';
				closeAllActiveWindow();
				plugin.executeMethod('AddToolbarMenuItem', [{guid: plugin.guid, tabs:[]}]);
				bCreateMenu = true;
				break;

			case 'onAddApiKey':
				pluginSettings = JSON.parse(message.settings);
				localStorage.setItem(localStorageKey, message.settings);
				pluginWindows.settingsWindow.close();
				pluginWindows.settingsWindow = null;
				if (bCreateMenu)
					createToolbarMenu();

				bCreateMenu = false;
				updateActiveWindowSettings();
				break;
	}
	};

	function createWindow() {
		let newWindow = new window.Asc.PluginWindow();
		newWindow.attachEvent("onWindowMessage", function(message) {
			messageHandler(newWindow, message);
		});
		return newWindow;
	};

	function createTranslatorWindow() {
		// default settings for modal window (I created separate settings, because we have many unnecessary field in plugin variations)
		let variation = {
			url : makeFileUrl('translate.html'),
			description : plugin.tr('AI translator'),
			isVisual : true,
			buttons : [],
			type: 'panel',
			EditorsSupport : ["word", "slide", "cell"],
		};

		if (pluginWindows.translatorWindow) {
			pluginWindows.translatorWindow.close();
			pluginWindows.translatorWindow = null;
		} else {
			pluginWindows.translatorWindow = createWindow();
			pluginWindows.translatorWindow.show(variation);
		}
	};

	function createSettingsWindow() {
		// default settings for modal window (I created separate settings, because we have many unnecessary field in plugin variations)
		let variation = {
			url : makeFileUrl('settings.html'),
			description : plugin.tr('AI translator'),
			isVisual : true,
			type: 'window',
			EditorsSupport : ["word", "slide", "cell"],size : [343, 122],
			buttons : [
				{
					"text" : "Save",
					"primary" : true,
					"textLocale" : {
						"ru"    : "Сохранить",
						"fr"    : "Enregistrer",
						"es"    : "Guardar",
						"pt-BR" : "Salvar",
						"de"    : "Speichern",
						"cs"    : "Uložit",
						"zh"    : "保存"
					}
				}
			]
		};

		if (!pluginWindows.settingsWindow)
			pluginWindows.settingsWindow = createWindow();

		pluginWindows.settingsWindow.show(variation);
	};

	function createTextToImageWindow() {
		// default settings for modal window (I created separate settings, because we have many unnecessary field in plugin variations)
		let variation = {
			url : makeFileUrl('textToImage.html'),
			description : plugin.tr('Text to image'),
			isVisual : true,
			buttons : [],
			type: 'panel',
			EditorsSupport : ["word", "slide", "cell"],
		};

		if (pluginWindows.textToImageWindow) {
			pluginWindows.textToImageWindow.close();
			pluginWindows.textToImageWindow = null;
		} else {
			pluginWindows.textToImageWindow = createWindow();
			pluginWindows.textToImageWindow.show(variation);
		}
	};

	function createKeyWordsWindow() {
		// default settings for modal window (I created separate settings, because we have many unnecessary field in plugin variations)
		let variation = {
			url : makeFileUrl('keyWords.html'),
			description : plugin.tr('Key words'),
			isVisual : true,
			buttons : [],
			type: 'panel',
			EditorsSupport : ["word", "slide", "cell"],
		};

		if (pluginWindows.keyWordsWindow) {
			pluginWindows.keyWordsWindow.close();
			pluginWindows.keyWordsWindow = null;
		} else {
			pluginWindows.keyWordsWindow = createWindow();
			pluginWindows.keyWordsWindow.show(variation);
		}
	};

	function creatOleVersionWindow() {
		// default settings for modal window (I created separate settings, because we have many unnecessary field in plugin variations)
		let variation = {
			url : makeFileUrl('oleHistory.html'),
			description : plugin.tr('Verion history'),
			isVisual : true,
			type: 'window',
			EditorsSupport : ["word", "slide", "cell"],
			size: [689, 428],
			buttons : [
				{
					text : "Restore",
					primary : true,
					textLocale : {
						"ru"    : "Восстановить",
						"fr"    : "Restaurer",
						"es"    : "Restaurar",
						"pt-BR" : "Restaurar",
						"de"    : "Wiederherstellen",
						"cs"    : "Obnovit",
						"zh"    : "还原",
						"ja"    : "復元"
					}
				},
				{
					text : "Cancel",
					textLocale : {
						"ru"    : "Отменить",
						"fr"    : "Annuler",
						"es"    : "Cancelar",
						"pt-BR" : "Cancelar",
						"de"    : "Stornieren",
						"cs"    : "Zrušit",
						"zh"    : "取消",
						"ja"    : "キャンセル"
					}
				}
			]
		};


		pluginWindows.oleHistoryWindow = createWindow();
		pluginWindows.oleHistoryWindow.show(variation);
	};

	function createSummarizationWindow() {
		// default settings for modal window (I created separate settings, because we have many unnecessary field in plugin variations)
		let variation = {
			url : makeFileUrl('summarization.html'),
			description : plugin.tr('Success!'),
			isVisual : true,
			type: 'window',
			EditorsSupport : ["word", "slide", "cell"],
			size : [328, 140],
			buttons : []
		};

		if (!pluginWindows.summarizationWindow)
			pluginWindows.summarizationWindow = createWindow();

		pluginWindows.summarizationWindow.show(variation);
	};

	function makeSummarization() {
		plugin.executeMethod('GetSelectedText', null, function(text) {
			if (!isEmpyText(text)) {
				let tokens = window.Asc.OpenAIEncode(text);
				createSettings(text, tokens, 1);
			}
		});
	};

	function pasteSummarization(bReview) {
		if (bReview) {
			// todo add review
			plugin.executeMethod('PasteText', [Asc.scope.summarizationResult.join(' ')]);
		} else {
			plugin.callCommand(function() {
				let oDocument = Api.GetDocument();
				let sumPar = Api.CreateParagraph();
				sumPar.AddText(Asc.scope.summarizationTitle);
				sumPar.SetBold(true);
				oDocument.Push(sumPar);
				for(let ind = 0; ind < Asc.scope.summarizationResult.length; ind++) {
					let text = Asc.scope.summarizationResult[ind];
					if (text.length) {
						let oParagraph = Api.CreateParagraph();
						oParagraph.AddText(text);
						oDocument.Push(oParagraph);
					}
				}
			}, false);
		}
		pluginWindows.summarizationWindow.close();
		pluginWindows.summarizationWindow = null;
	};

	function updateOleObject(active) {
		if (active) {
			let props = oleData.data[active];
			let info = plugin.info;
			let data = JSON.stringify(oleData.data);
			const param = {
				guid : info.guid,
				widthPix : props.width,
				heightPix : props.height,
				width : props.width / info.mmToPx,
				height : props.height / info.mmToPx,
				imgSrc : props.src,
				data : data,
				objectId : oleData.objectId,
				resize : true,
				recalculate: true
			};
			plugin.executeMethod("EditOleObject", [param]);
			pluginWindows.oleHistoryWindow.close();
			pluginWindows.oleHistoryWindow = null;
		}
	};

	function closeAllActiveWindow() {
		for (const winName in pluginWindows) {
			if (Object.hasOwnProperty.call(pluginWindows, winName)) {
				const win = pluginWindows[winName];
				if (win) {
					win.close();
					pluginWindows[winName] = null;
				}
			}
		}
	};

	function updateActiveWindowSettings() {
		for (const winName in pluginWindows) {
			if (Object.hasOwnProperty.call(pluginWindows, winName)) {
				const win = pluginWindows[winName];
				if (win) {
					win.command('onApiKey', JSON.stringify(pluginSettings));
				}
			}
		}
	};

	function isEmpyText(text, bDonShowErr) {
		if (text.trim() === '') {
			if (!bDonShowErr)
				console.error('No word in this position or nothing is selected.');

			return true;
		}
		return false;
	};

	function makeFileUrl(name) {
		let location  = window.location;
		let start = location.pathname.lastIndexOf('/') + 1;
		let file = location.pathname.substring(start);
		return location.href.replace(file, name)
	};

})(window, undefined);
