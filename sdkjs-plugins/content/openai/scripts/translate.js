/**
 *
 * (c) Copyright Ascensio System SIA 2020
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */
(function(window, undefined) {

	let	loader,
		bCreateLoader = true,
		elements = {},
		errTimeout = null,
		errMessage = 'Invalid Api key or this service doesn\'t work in your region.',
		translationSetting = {
			langFrom: 'English',
			langTo: 'English',
			type: 'Word to word',
			audience: 'Average',
			tone: 'Official'
		},
		pluginSettings = {
			apiKey: '',
			model: 'gpt-3.5-turbo',
			maxTokens: 4000
		};

	const url = 'https://api.openai.com/v1/chat/completions';
	const localStorageKey = 'AI_translator_settings';
	const langs = [
		{id: "English", text: "English"},
		{id: "French", text: "French"},
		{id: "Germam", text: "Germam"},
		{id: "Chinese", text: "Chinese"},
		{id: "Japanese", text: "Japanese"},
		{id: "Russian", text: "Russian"},
		{id: "Korean", text: "Korean"},
		{id: "Spanish", text: "Spanish"},
		{id: "Italian", text: "Italian"},
		{id: "Portuguese", text: "Portuguese"}
	];

	const tones = [
		{id: "Official", text: "Official"},
		{id: "Professional", text: "Professional"},
		{id: "Casual", text: "Casual"},
		{id: "Friendly", text: "Friendly"},
		{id: "Sophisticated", text: "Sophisticated"},
		{id: "Simple", text: "Simple"}
	];

	const audiences = [
		{id: "Average", text: "Average"},
		{id: "Children", text: "Children"},
		{id: "Teenagers", text: "Teenagers"},
		{id: "Education", text: "Education"},
		{id: "Technology", text: "Technology"},
		{id: "Finance", text: "Finance"},
		{id: "Government", text: "Government"},
		{id: "Medicine", text: "Medicine"},
		{id: "Sales", text: "Sales"}
	];

	const types = [
		{id: "Word to word", text: "Word to word"},
		{id: "Summary", text: "Summary"}
	];
	
	window.Asc.plugin.init = function() {
		restoreSettings();
		sendPluginMessage({type: 'onWindowReady'});
		bCreateLoader = true;
		createLoader();
		initElements();
		createSelects();
		updateMaxHeight();

		elements.btnShowSettins.onclick = function() {
			elements.divParams.classList.toggle('hidden');
			elements.arrow.classList.toggle('arrow_down');
			elements.arrow.classList.toggle('arrow_up');
			updateMaxHeight();
		};

		elements.btnSwitch.onclick = function() {
			makeTranslation();
		};

		elements.btnCopy.onclick = function() {
			if (window.getSelection && document.createRange) { //Browser compatibility
				let range = document.createRange(); //range object
				range.selectNodeContents(elements.resultField);  //sets Range
				sel = window.getSelection();
				sel.removeAllRanges(); //remove all ranges from selection
				sel.addRange(range); //add Range to a Selection.
				document.execCommand("copy"); //copy
				sel.removeAllRanges(); //remove all ranges from selection
			}
		};

		elements.btnPaste.onclick = function() {
			// todo включить рецензирование и после вставки выключить его
			window.Asc.plugin.executeMethod("PasteText", [elements.resultField.innerText]);
		};
	};

	window.Asc.plugin.onTranslate = function() {
		errMessage = window.Asc.plugin.tr(errMessage);
		if (bCreateLoader)
			createLoader();
		
		let elements = document.querySelectorAll('.i18n');
		bCreateLoader = true;

		elements.forEach(function(element) {
			element.innerHTML = getTranslated(element.innerHTML);
		});
	};

	function createLoader() {
		$('#loader-container').removeClass( "hidden" );
		loader && (loader.remove ? loader.remove() : $('#loader-container')[0].removeChild(loader));
		loader = showLoader($('#loader-container')[0], getTranslated('Loading...'));
	};

	function destroyLoader() {
		$('#loader-container').addClass( "hidden" )
		loader && (loader.remove ? loader.remove() : $('#loader-container')[0].removeChild(loader));
		loader = undefined;
	};

	function setError(error, bConst, bQuotaErr) {
		if (bQuotaErr) {
			elements.mainErrorLb.innerHTML = '<label style="martgin-bottom:5px">' + getTranslated('You exceeded your current quota, please check your plan and billing details.') + '</label>' +
			'<p></p> <label>' + getTranslated('For more information on this error, read the docs:') + '</label>' +
			'<br> <a target="_blank" class="text-link" href="https://platform.openai.com/docs/guides/error-codes/api-errors.">https://platform.openai.com/docs/guides/error-codes/api-errors.</a>';
		} else {
			elements.mainErrorLb.innerHTML = '<label>' + getTranslated(error) + '</label>';
		}
		elements.mainError.classList.remove('hidden');
		if (errTimeout) {
			clearTimeout(errTimeout);
			errTimeout = null;
		}
		if (!bConst)
			errTimeout = setTimeout(clearMainError, 5000);
	};

	function clearMainError() {
		elements.mainError.classList.add('hidden');
		elements.mainErrorLb.innerHTML = '';
	};

	function initElements() {
		elements.mainError      = document.getElementById('div_err');
		elements.mainErrorLb    = document.getElementById('lb_err');
		elements.divParams      = document.getElementById('div_parametrs');
		elements.btnShowSettins = document.getElementById('div_show_settings');
		elements.btnSwitch      = document.getElementById('btn_switch');
		elements.arrow          = document.getElementById('arrow');
		elements.resultField    = document.getElementById('div_result');
		elements.translatedCont = document.getElementById('translated_container');
		elements.btnCopy        = document.getElementById('btn_copy');
		elements.btnPaste       = document.getElementById('btn_paste');
	};

	function createSelects() {
		$('#sel_source_lang').select2({
			data : langs,
		}).on('select2:select', function(e) {
			translationSetting.langFrom = $('#sel_source_lang :selected').text();
		});
		$('#sel_source_lang').val(translationSetting.langFrom).trigger("change");


		$('#sel_target_lang').select2({
			data : langs,
			// sorter: function(data){ return data.sort(function(a, b){ return a.text.localeCompare(b.text ) } ) }
		}).on('select2:select', function(e) {
			translationSetting.langTo = $('#sel_target_lang :selected').text();
		});
		$('#sel_target_lang').val(translationSetting.langTo).trigger("change");


		$('#sel_translation_tone').select2({
			data : tones,
		}).on('select2:select', function(e) {
			translationSetting.tone = $('#sel_translation_tone :selected').text();
		});
		$('#sel_translation_tone').val(translationSetting.tone).trigger("change");


		$('#sel_translation_type').select2({
			data : types,
		}).on('select2:select', function(e) {
			translationSetting.type = $('#sel_translation_type :selected').text();
		});
		$('#sel_translation_type').val(translationSetting.type).trigger("change");


		$('#sel_target_audience').select2({
			data : audiences,
		}).on('select2:select', function(e) {
			translationSetting.audience = $('#sel_target_audience :selected').text();
		});
		$('#sel_target_audience').val(translationSetting.audience).trigger("change");

	};

	function updateMaxHeight() {
        var contentTopOffset = elements.resultField.getBoundingClientRect().top;
        var windowHeight = window.innerHeight;
        var contentMaxHeight = windowHeight - contentTopOffset;
        elements.resultField.style.maxHeight = elements.resultField.style.minHeight = contentMaxHeight - 46 + 'px';

    };

	function makeTranslation() {
		window.Asc.plugin.executeMethod('GetSelectedText', null, function(text) {
			if (!isEmpyText(text)) {
				createLoader();
				let message  = `Translate this text from ${translationSetting.langFrom} to ${translationSetting.langTo} with some settings:tone of tranlation: ${translationSetting.tone}, target audience: ${translationSetting.audience}, type of translation: ${translationSetting.type}.  Text: ${text}`;
				let tokens = window.Asc.OpenAIEncode(message);
				let settings = {
					model : pluginSettings.model,
					max_tokens : pluginSettings.maxTokens - tokens.length,
					messages: [ { role: 'user', content: message } ]
				};
				if (settings.max_tokens < 200) {
					setError('Too many tokens in selection. Limit is: ' + keyWordsSettings.maxTokens, false, false);
					return;
				}
				fetch(url, {
					method: 'POST',
					headers: {
						'Authorization': 'Bearer ' + pluginSettings.apiKey,
						'Content-Type' : 'application/json'
					},
					body: JSON.stringify(settings),
				})
				.then(function(response) {
					return response.json()
				})
				.then(function(data) {
					if (data.error)
						throw data.error
	
					destroyLoader();
					text = data.choices[0].message.content.startsWith('\n\n') ? data.choices[0].message.content.substring(2) : data.choices[0].message.content;
					elements.resultField.innerText = text;
					elements.translatedCont.classList.remove('invisible');
				})
				.catch(function(error) {
					destroyLoader();
					let bQuotaErr = error.type === 'insufficient_quota';
					setError(error.message, false, bQuotaErr);
					console.error(error);
				});
			}
		});
	}

	window.onresize = function() {
		updateMaxHeight();
	};

	function getTranslated(key) {
		return window.Asc.plugin.tr(key);
	};

	function sendPluginMessage(message) {
		window.Asc.plugin.sendToPlugin("onWindowMessage", message);
	};

	function isEmpyText(text) {
		if (text.trim() === '') {
			let err = new Error('No word in this position or nothing is selected.');
			setError(err.message)
			console.error(err);
			return true;
		}
		return false;
	};

	function restoreSettings() {
		try {
			let value = localStorage.getItem(localStorageKey);
			if (value)
				translationSetting = JSON.parse(value);
		} catch (error) {
			// we have some problem with saved message. we should remove it from localStorage
			localStorage.removeItem(localStorageKey);
		}
	};

	window.Asc.plugin.attachEvent("onApiKey", function(settings) {
		pluginSettings = JSON.parse(settings);
		bCreateLoader = false;
		destroyLoader();
		if (!pluginSettings.apiKey) {
			setError('Api key not found, please make enter it into the settings window.', false, false)
		}
	});

	window.Asc.plugin.attachEvent('onSaveSetting', function() {
		localStorage.setItem(localStorageKey, JSON.stringify(translationSetting));
	});

	window.Asc.plugin.onThemeChanged = function(theme) {
		// todo в окна плагина не приходит сообщение о смене темы, пока единственный вариант - посылать это событие с темой из основного плагина (но на старте оно приходит)
		console.log('onThemeChanged');
		window.Asc.plugin.onThemeChangedBase(theme);

		let rule = ".select2-container--default.select2-container--open .select2-selection__arrow b { border-color : " + window.Asc.plugin.theme["text-normal"] + " !important; }";
		rule += '\n .arrow { border-color : ' + theme['text-normal'] + ' !important; }';
		rule += '\n #div_result::-webkit-scrollbar-track, .err_background { background: ' + theme['background-toolbar'] + ' !important; }';
		let styleTheme = document.createElement('style');
		styleTheme.type = 'text/css';
		styleTheme.innerHTML = rule;
		document.getElementsByTagName('head')[0].appendChild(styleTheme);
	};


})(window, undefined);
