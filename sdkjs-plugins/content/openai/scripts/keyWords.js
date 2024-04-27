/**
 *
 * (c) Copyright Ascensio System SIA 2020
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *	 http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */
(function (window, undefined) {
	let loader;
	let elements = {};
	let errTimeout = null;
	let keyWordsSettings = {};
	const url = 'https://api.openai.com/v1/chat/completions';
	
	window.Asc.plugin.init = function() {
		sendPluginMessage({type: 'onWindowReady'});
		initElements();
		
		$('#sel_count').select2({});

		elements.btnGenerate.onclick = function(e) {
			window.Asc.plugin.executeMethod('GetSelectedText', null, function(text) {
				text = text.trim();
				if (text) {
					let tokens = window.Asc.OpenAIEncode(text);
					let count = Number( $('#sel_count').val() );
					let settings = createSettings(text, tokens, count);
					fetchData(settings);
				} else {
					setError('Select some text for key words generating.', false, false);
				}
			});
		};

		elements.btnInsert.onclick = function(e) {
			let selected = document.getElementsByClassName('word_selected');
			let selectedString = ' ';
			for (let index = 0; index < selected.length; index++) {
				selectedString += selected[index].innerText + (index !== selected.length - 1 ? ', ' : '');
			}
			Asc.scope.data = [getTranslated('Key Words:'), selectedString];
			window.Asc.plugin.callCommand(function() {
				let oDocument = Api.GetDocument();
				let oParagraph = Api.CreateParagraph();
				var oRun = Api.CreateRun();
				oRun.AddText(Asc.scope.data[0]);
				oRun.SetBold(true);
				oParagraph.AddElement(oRun);
				oParagraph.AddText(Asc.scope.data[1]);
				oDocument.Push(oParagraph);
			}, false);
		};
	};

	function initElements() {
		elements.btnGenerate     = document.getElementById('btn_generate');
		elements.btnInsert       = document.getElementById('btn_insert');
		elements.selectCount     = document.getElementById('sel_count');
		elements.divResult       = document.getElementById('div_result');
		elements.divWords        = document.getElementById('div_words');
		elements.mainError       = document.getElementById('div_err');
		elements.mainErrorLb     = document.getElementById('lb_err');
	};

	function createSettings(text, tokens, count) {
		if ( tokens > (keyWordsSettings.maxTokens - 100) ) {
			setError('Too many tokens in selection. Limit is: ' + keyWordsSettings.maxTokens, false, false);
		} else {
			let settings = {
				messages : [ { role: 'user', content: `Get ${count} Key words from this text as javascript array: "${text}"` } ],
				model: keyWordsSettings.model,
				max_tokens: 100
			};
			return settings;
		}
	};

	function createLoader() {
		$('#loader-container').removeClass( 'hidden' );
		loader && (loader.remove ? loader.remove() : $('#loader-container')[0].removeChild(loader));
		loader = showLoader($('#loader-container')[0], getTranslated('Loading...'));
	};

	function destroyLoader() {
		$('#loader-container').addClass( 'hidden' )
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

	function getTranslated(key) {
		return window.Asc.plugin.tr(key);
	};

	function sendPluginMessage(message) {
		window.Asc.plugin.sendToPlugin('onWindowMessage', message);
	};

	function fetchData(settings) {
		createLoader();
		let header = {
			'Authorization': 'Bearer ' + keyWordsSettings.apiKey,
			'Content-Type': 'application/json'
		};
		fetch(url, {
				method: 'POST',
				headers: header,
				body: JSON.stringify(settings),
			})
			.then(function(response) {
				return response.json()
			})
			.then(function(data) {
				if (data.error)
					throw data.error

				let words = JSON.parse(data.choices[0].message.content);
				createWords(words);
			})
			.catch(function(error) {
				elements.divResult.classList.add('hidden');
				let bQuotaErr = error.type === 'insufficient_quota';
				setError(error.message, false, bQuotaErr);
				console.error('Error:', error);
			})
			.finally(function(){
				destroyLoader();
			});
	};

	function createWords(words) {
		elements.divWords.innerHTML = '';
		words.forEach(function(text) {
			let word = document.createElement('span');
			word.classList.add('word_bg', 'word');
			word.innerText = text;
			word.onclick = function(e) {
				e.target.classList.toggle('word_selected');
			};
			elements.divWords.appendChild(word);
		});
		elements.divResult.classList.remove('hidden');
	};

	window.Asc.plugin.onTranslate = function() {
		let elements = document.querySelectorAll('.i18n');

		for (let index = 0; index < elements.length; index++) {
			let element = elements[index];
			if (element.attributes['placeholder']) element.attributes['placeholder'].value = getTranslated(element.attributes['placeholder'].value);
			element.innerText = getTranslated(element.innerText);
		}
	};

	window.Asc.plugin.onThemeChanged = function(theme) {
		// todo в окна плагина не приходит сообщение о смене темы, пока единственный вариант - посылать это событие с темой из основного плагина (но на старте оно приходит)
		window.Asc.plugin.onThemeChangedBase(theme);

		let rule = '.select2-container--default.select2-container--open .select2-selection__arrow b { border-color : ' + window.Asc.plugin.theme['text-normal'] + ' !important; }';
		rule += '\n .white_bg { background-color: ' + theme['background-normal'] + ' !important; }';
		rule += '\n .gray_bg { background-color: ' + theme['background-toolbar'] + ' !important; }';
		rule += '\n .word_bg { background-color: ' + theme['border-toolbar-button-hover'] + ' !important; }';
		rule += '\n .word_selected { background-color: #a5a5a5 !important; }';
		let styleTheme = document.createElement('style');
		styleTheme.type = 'text/css';
		styleTheme.innerHTML = rule;
		document.getElementsByTagName('head')[0].appendChild(styleTheme);
	};

	window.Asc.plugin.attachEvent('onApiKey', function(settings) {
		keyWordsSettings = JSON.parse(settings);
	});

})(window, undefined);
