/**
 *
 * (c) Copyright Ascensio System SIA 2020
 *
 * Licensed under the Apache License, Version 2.0 (the 'License');
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an 'AS IS' BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */
(function(window, undefined) {
	let loader = null;
	let errMessage = 'Invalid Api key or this service doesn\'t work in your region.';
	let loadMessage = 'Loading...';
	const localStorageKey = 'OpenAISettings';
	let settings = {
		apiKey: '',
		model: '',
		maxTokens: ''
	};

	window.Asc.plugin.init = function() {
		// createLoader();
		// window.Asc.plugin.resizeWindow(343, 122, 343, 122, 343, 122);
		restoreSettings();
	};

	function createError(error) {
		document.getElementById('inp_key').classList.add('error_border');
		document.getElementById('err_message').innerText = errMessage;
		console.error(error.message || errMessage);
	};

	function createLoader() {
		if (!window.Asc.plugin.theme)
			window.Asc.plugin.theme = {type: 'light'};
		$('#loader-container').removeClass( 'hidden' );
		loader && (loader.remove ? loader.remove() : $('#loader-container')[0].removeChild(loader));
		loader = showLoader($('#loader-container')[0], loadMessage);
	};

	function destroyLoader() {
		$('#loader-container').addClass( 'hidden' )
		loader && (loader.remove ? loader.remove() : $('#loader-container')[0].removeChild(loader));
		loader = null;
	};

	function getMaxTokens(model) {
		let result = 4000;
		let arr = model.split('-');
		let length = arr.find(function(el){return (el.slice(-1) == 'k' && el.length <= 3)});
		if (length) {
			result = Number(length.slice(0,-1)) * 1000;
		}
		return result;
	};

	function restoreSettings() {
		try {
			let value = localStorage.getItem(localStorageKey);
			if (value)
				settings = JSON.parse(value);

			document.getElementById('inp_key').value = settings.apiKey;
		} catch (error) {
			// we have some problem with saved message. we should remove it from localStorage
			sendPluginMessage({type: 'onRemoveApiKeyAndCLose'});
		}
	};

	function sendPluginMessage(message) {
		window.Asc.plugin.sendToPlugin('onWindowMessage', message);
	}

	window.Asc.plugin.onTranslate = function() {
		errMessage = window.Asc.plugin.tr(errMessage);
		loadMessage = window.Asc.plugin.tr(loadMessage);
		let elements = document.querySelectorAll('.i18n');
		elements.forEach(function(element) {
			element.innerText = window.Asc.plugin.tr(element.innerText);
		})
	};

	
	window.Asc.plugin.onThemeChanged = function(theme) {
		// todo в окна плагина не приходит сообщение о смене темы, пока единственный вариант - посылать это событие с темой из основного плагина (но на старте оно приходит)
		window.Asc.plugin.onThemeChangedBase(theme);
		let rule = '\n .loader-bg { background-color: ' + theme['background-normal'] + ' !important; }';
		rule += '.text-link, .text-link:hover, .text-link:active, .text-link:visited{color: ' + theme['text-normal'] + ' !important;}';
		let styleTheme = document.createElement('style');
		styleTheme.type = 'text/css';
		styleTheme.innerHTML = rule;
		document.getElementsByTagName('head')[0].appendChild(styleTheme);
	};

	window.Asc.plugin.attachEvent("onSaveBtn", function() {
		document.getElementById('inp_key').classList.remove('error_border');
		document.getElementById('err_message').innerText = '';
		document.getElementById('success_message').classList.add('hidden');
		let key = document.getElementById('inp_key').value.trim();
		if (key.length) {
			sendPluginMessage({type: 'onRemoveApiKey'});
			createLoader();
			// check api key by fetching models
			fetch('https://api.openai.com/v1/models', {
				method: 'GET',
				headers: {
					'Authorization': 'Bearer ' + key
				}
			}).
			then(function(response) {
				if (response.ok) {
					response.json().then(function(data) {
						let firsVar = {model: null, maxTokens: 0};
						let secondVar = {model: null, maxTokens: 0};
						let thirdVar = {model: null, maxTokens: 0};
						let fourthVar = {model: null, maxTokens: 0};
						for (let index = 0; index < data.data.length; index++) {
							let cur = data.data[index].id;
							if (cur === 'gpt-4') {
								firsVar.model = cur;
								firsVar.maxTokens = 8000;
							}
							
							if (cur.includes('gpt-4')) {
								secondVar.model = cur;
								secondVar.maxTokens = getMaxTokens(cur);
							}

							if (cur === 'gpt-3.5-turbo-16k') {
								thirdVar.model = cur;
								thirdVar.maxTokens = 16000;
							}

							if (cur.includes('gpt-3.5')) {
								fourthVar.model = cur;
								fourthVar.maxTokens = getMaxTokens(cur);
							}
						}
						settings.model = (firsVar.model || secondVar.model || thirdVar.model || fourthVar.model);
						settings.maxTokens = (firsVar.maxTokens || secondVar.maxTokens || thirdVar.maxTokens  || fourthVar.maxTokens);
						settings.apiKey = key;
						if (settings.model) {
							// todo add extra check (if token is expired). Maybe send request '2+2?' or 'say hi';
							// because we can receive list of model if our token is expired.

							// create AI tab in toolbar after it
							sendPluginMessage({type: 'onAddApiKey', settings: JSON.stringify(settings)});
							document.getElementById('success_message').classList.remove('hidden');
						} else {
							createError(new Error(errMessage));
						}
					});
				} else {
					response.json().then(function(data) {
						let message = data.error && data.error.message ? data.error.message : errMessage;
						createError(new Error(message));
					});
				}
			})
			.catch(function(error) {
				createError(error);
				sendPluginMessage({type: 'onRemoveApiKeyAndCLose'});
			})
			.finally(function(){
				destroyLoader();
			});
		} else {
			sendPluginMessage({type: 'onRemoveApiKeyAndCLose'});
			// createError(new Error(errMessage));
		}
	});

})(window, undefined);
